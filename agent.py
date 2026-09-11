"""本地代理（學生電腦上執行的東西）。

只做三件事：
  1. 向雲端取得工具清單，交給 Claude Desktop
  2. 把工具呼叫原封不動轉給雲端
  3. 執行雲端傳回的動作，把結果送回去，直到雲端說完成

不含任何業務邏輯。逆向出來只會看到「怎麼讀寫 Aspen 節點」與這段轉發迴圈。

執行：D:\\Aspen_MCP\\.venv\\Scripts\\python.exe agent.py
（需要 pywin32 與 mcp 套件，Aspen venv 兩者都有）
"""
from __future__ import annotations

import atexit
import json
import os
import sys
import time
import urllib.error
import urllib.request
import uuid

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bridge import Bridge  # noqa: E402

CLOUD = os.environ.get("ASPEN_CLOUD", "http://127.0.0.1:8787")
TOKEN = os.environ.get("ASPEN_TOKEN", "demo-token")
MAX_ROUNDS = 500         # 防止雲端邏輯有 bug 時無限迴圈
                         # 執行一次模擬會輪詢數十次，上限不能訂得比它低


# 雲端重啟中的那幾秒，TCP 會直接拒絕（WinError 10061）。以前一次拒絕就
# 整個工具呼叫失敗，而且訊息只有一串 errno；重啟是常態，不該讓學生端
# 為此整場崩掉。這裡短暫重試，仍然連不上再清楚地講出來。
RETRY_S = (0.5, 1.0, 2.0, 3.0, 5.0)


def post(path: str, payload: dict, timeout: float = 60.0) -> dict:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        CLOUD + path, data=body,
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer " + TOKEN})
    last = None
    for wait in RETRY_S + (None,):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.URLError as exc:
            reason = getattr(exc, "reason", exc)
            refused = (isinstance(reason, ConnectionRefusedError)
                       or "10061" in str(reason)
                       or "refused" in str(reason).lower())
            if not refused or wait is None:
                if refused:
                    raise ConnectionError(
                        "連不上雲端 {}：連線被拒絕，已重試 {} 秒。伺服器可能"
                        "沒在跑或正在重啟；請稍後再試，或請管理者確認 8787 "
                        "有在監聽。".format(CLOUD, int(sum(RETRY_S)))) from exc
                raise
            last = exc
            time.sleep(wait)
    raise last  # 理論上到不了這裡


class Agent:
    def __init__(self) -> None:
        self.bridge = Bridge()
        self.session = uuid.uuid4().hex[:12]
        self.rounds = 0
        self.ops_sent = 0
        # 這支程式被 Claude Desktop 收掉時沒有人會呼叫 close_project，
        # 那一份 AspenPlus.exe 就留下來鎖著 .bkp。行程結束前補一次關閉。
        atexit.register(self.shutdown)

    def list_tools(self) -> list:
        """工具清單從雲端來，不寫死在本地。

        這樣雲端隨時可以增減工具、改參數，學生端不用重裝。
        """
        return post("/tools", {"token": TOKEN})["tools"]

    def call(self, tool: str, args: dict) -> dict:
        """轉發呼叫，然後執行雲端指揮的動作，直到雲端說完成。

        本地永遠不知道下一步是什麼 —— 配方不落地，這是保護的關鍵。
        """
        reply = post("/call", {"session": self.session, "tool": tool,
                               "args": args, "token": TOKEN})
        self.rounds = 0
        while reply.get("type") == "ops":
            self.rounds += 1
            if self.rounds > MAX_ROUNDS:
                return {"error": "超過 {} 回合，雲端可能有迴圈".format(MAX_ROUNDS)}
            results = []
            for item in reply["ops"]:
                self.ops_sent += 1
                results.append(self.bridge.execute(item.get("op"),
                                                   item.get("args")))
            reply = post("/result", {"session": self.session,
                                     "results": results, "token": TOKEN})
        return reply.get("result", {})

    def shutdown(self) -> None:
        try:
            self.bridge.execute("close")
        except Exception:
            pass


# ── MCP stdio server ────────────────────────────────────────────────
def serve_mcp() -> None:
    """把雲端的工具清單，以 MCP 介面提供給 Claude Desktop。"""
    import asyncio
    from concurrent.futures import ThreadPoolExecutor

    from mcp.server import Server
    from mcp.server.stdio import stdio_server
    from mcp.types import TextContent, Tool

    agent = Agent()
    app = Server("aspen-agent")

    # **所有對 Aspen 的操作都必須在同一個執行緒上。**
    #
    # COM 物件屬於建立它的執行緒。原本這裡用 asyncio.to_thread，
    # 那是一個多執行緒的池子 —— connect 可能落在 A 執行緒、
    # 後面的 open_file 落在 B 執行緒，於是開檔失敗。
    #
    # 直接呼叫 Agent 類別測不出這個問題（那時只有一個執行緒），
    # 只有走真正的 MCP 路徑才會現形。max_workers=1 就是為了這個。
    def _com_init():
        """在工作執行緒上初始化 COM。

        主執行緒是隱式初始化的，所以直接呼叫 Agent 時不會有問題 ——
        這就是為什麼先前的測試通通看不到這個錯。走 MCP 之後工作
        落到別的執行緒上，沒有這一行就是
        「CoInitialize 尚未被呼叫」，而且訊息指不到真正的原因。
        """
        import pythoncom
        pythoncom.CoInitialize()

    pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="aspen",
                              initializer=_com_init)

    @app.list_tools()
    async def _tools():
        loop = asyncio.get_running_loop()
        tools = await loop.run_in_executor(pool, agent.list_tools)
        return [Tool(name=t["name"], description=t["description"],
                     inputSchema=t["inputSchema"]) for t in tools]

    @app.call_tool()
    async def _call(name: str, arguments: dict):
        loop = asyncio.get_running_loop()
        result = await loop.run_in_executor(
            pool, agent.call, name, arguments or {})
        return [TextContent(type="text",
                            text=json.dumps(result, ensure_ascii=False, indent=2))]

    async def main():
        async with stdio_server() as (r, w):
            await app.run(r, w, app.create_initialization_options())

    try:
        asyncio.run(main())
    finally:
        # 關閉也要在同一個執行緒上做，否則 COM 物件會在錯的 apartment 被釋放
        try:
            pool.submit(agent.shutdown).result(timeout=30)
        except Exception:
            pass
        pool.shutdown(wait=False)


# ── 不經 MCP 的直接測試 ─────────────────────────────────────────────
def selftest(bkp: str) -> int:
    """繞過 MCP，直接驗證「代理 ↔ 雲端 ↔ 橋接 ↔ Aspen」整條迴圈。"""
    import subprocess
    import time

    for image in ("AspenPlus.exe", "AspenProperties.exe"):
        subprocess.run(["taskkill", "/F", "/IM", image],
                       capture_output=True, check=False)

    agent = Agent()
    try:
        tools = agent.list_tools()
        print("[1] 從雲端取得 {} 個工具：{}".format(
            len(tools), ", ".join(t["name"] for t in tools)))

        def step(n, tool, args, show=None):
            t0 = time.time()
            r = agent.call(tool, args)
            print("[{}] {:<14} {} 回合 {:.1f}s  {}".format(
                n, tool, agent.rounds, time.time() - t0,
                show(r) if show else r))
            return r

        r = step(2, "open_project", {"path": bkp},
                 lambda r: "opened={} 單位表 {} 筆".format(
                     r.get("opened"), r.get("units_loaded")))
        if not r.get("opened"):
            print("    ", r.get("error"))
            return 1

        step(3, "list_model", {},
             lambda r: "區塊 {} 物流 {}".format(r.get("blocks"), r.get("streams")))

        r = step(4, "get_specs",
                 {"target": "block", "name": "COL1", "only_set": True},
                 lambda r: "{} 個已設定的欄位".format(r.get("count")))
        for k in ("NSTAGE", "BASIS_RR", "PRES1", "BASIS_D"):
            if k in (r.get("specs") or {}):
                v = r["specs"][k]
                print("      {:<10} {:<10} {}".format(
                    k, v["value"], v["unit"] or ""))

        step(5, "set_specs",
             {"target": "block", "name": "COL1",
              "specs": {"NSTAGE": 25, "PRES1": {"value": 1.5, "unit": "bar"}}},
             lambda r: "寫入 {} 個，退回 {} 個 {}".format(
                 r.get("n_applied"), r.get("n_rejected"), r.get("rejected") or ""))

        step(6, "set_specs",
             {"target": "stream", "name": "FEED", "specs": {"TEMP": 60}},
             lambda r: "退回 {} 個：{}".format(
                 r.get("n_rejected"),
                 [v.get("detail") for v in (r.get("rejected") or {}).values()]))

        print("\n本次共送出 {} 個動作".format(agent.ops_sent))
        print("學生端從頭到尾不知道：'bar' 對應哪個索引、寫入後要複查、"
              "漏單位要擋下來 —— 那些都在雲端。")
        return 0
    finally:
        agent.shutdown()
        for image in ("AspenPlus.exe", "AspenProperties.exe"):
            subprocess.run(["taskkill", "/F", "/IM", image],
                           capture_output=True, check=False)


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--selftest":
        raise SystemExit(selftest(sys.argv[2]))
    serve_mcp()
