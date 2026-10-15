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
from bridge import Bridge, _aspen_pids  # noqa: E402

CLOUD = os.environ.get("ASPEN_CLOUD", "http://127.0.0.1:8787")
TOKEN = os.environ.get("ASPEN_TOKEN", "demo-token")
MAX_ROUNDS = 500         # 防止雲端邏輯有 bug 時無限迴圈
                         # 執行一次模擬會輪詢數十次，上限不能訂得比它低


# 雲端重啟中的那幾秒，TCP 會直接拒絕（WinError 10061）。以前一次拒絕就
# 整個工具呼叫失敗，而且訊息只有一串 errno；重啟是常態，不該讓學生端
# 為此整場崩掉。這裡短暫重試，仍然連不上再清楚地講出來。
RETRY_S = (0.5, 1.0, 2.0, 3.0, 5.0)

# 一次 MCP 工具呼叫最多讓用戶端等幾秒（見 CallGate）。Codex 預設 60 秒就判
# 逾時，Claude Desktop 也有自己的上限；留一點餘裕。0 表示不限制。
CALL_BUDGET_S = float(os.environ.get("ASPEN_CALL_BUDGET_S") or 45)


def _cloud_error(exc: urllib.error.HTTPError) -> RuntimeError:
    """把雲端的錯誤回應翻成學生看得懂的一句話。

    不處理的話 MCP 只會顯示「HTTP Error 429: Too Many Requests」—— 分不出
    額度用完還是叫太快，也看不到雲端寫好的重置時間。
    """
    try:
        body = json.loads(exc.read().decode("utf-8") or "{}")
    except Exception:
        body = {}
    said = body.get("error") if isinstance(body, dict) else None
    if body.get("code") == "quota_exceeded":
        return RuntimeError(
            "這個帳號本期的工具呼叫額度已經用完（{}）。重試不會成功，這也不是"
            "模型或 Aspen 的問題；請等重置時間過後再用，需要更多額度請聯絡課程"
            "助教。".format(said or "用量已達上限"))
    if exc.code == 429:
        return RuntimeError("{}。約一分鐘後再呼叫。".format(
            said or "呼叫太頻繁，請稍候再試"))
    if exc.code == 401:
        return RuntimeError("{}（請確認 ASPEN_TOKEN 沒有打錯）".format(
            said or "憑證無效或帳號已停權"))
    return RuntimeError("雲端回應 HTTP {}：{}".format(exc.code, said or exc.reason))


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
        except urllib.error.HTTPError as exc:
            # 雲端有回應、只是拒絕 —— 重試沒有意義，把它的說明轉出來
            raise _cloud_error(exc) from None
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


def _call_key(tool: str, args: dict) -> str:
    return tool + " " + json.dumps(args or {}, ensure_ascii=False, sort_keys=True,
                                   default=str)


class CallGate:
    """每次 MCP 呼叫在 budget 秒內一定回話；Aspen 一次只做一件事。

    2026-10-15 上線測試：「逾時，然後模型一直重新開 Aspen 檔案」。一次
    run_simulation 可以跑上幾分鐘（上限 timeout_s=1800），open_project、讀大
    模型的結果也常超過一分鐘；用戶端（Codex 預設 60 秒）先判逾時，模型以為
    壞了，就重開專案。可是舊的呼叫還在唯一的 Aspen 執行緒上跑，新的呼叫排在
    它後面 —— 又逾時、又重開，越排越長，每一個排到的 open_project 都把 Aspen
    關掉重開一次。

    這裡的做法：
      - 呼叫超過 budget 秒還沒完：先回 still_running，工作本身不中斷。
      - 之後用一模一樣的參數再呼叫一次：接著等，完成就交出那一次的結果。
      - 還沒完就呼叫別的工具：等到 budget 秒仍沒空，回 busy，**不執行、不排隊**。
      - 沒人來領的結果，附在下一次呼叫的回覆裡（earlier_call），不會不見。
    """

    def __init__(self, submit, budget: float = CALL_BUDGET_S) -> None:
        self.submit = submit          # (tool, args) -> asyncio.Future
        self.budget = budget
        self.inflight = None          # 正在跑的那一次
        self.unclaimed = None         # 跑完了、還沒交給任何人的那一次

    def _finished(self, rec: dict, fut) -> None:
        rec["done_at"] = time.monotonic()
        if self.inflight is rec:
            self.inflight = None
        if not rec.get("claimed"):
            self.unclaimed = rec

    async def _wait(self, fut, deadline):
        """等到 fut 完成回 True；過了 deadline 回 False（fut 照樣繼續跑）。"""
        import asyncio
        if self.budget <= 0:
            await asyncio.shield(fut)
            return True
        try:
            await asyncio.wait_for(asyncio.shield(fut), max(0.0, deadline - time.monotonic()))
            return True
        except asyncio.TimeoutError:
            return False

    @staticmethod
    def _outcome(rec: dict):
        """跑完的那一次 → 它的結果；那一次丟了例外就照樣丟出去。"""
        rec["claimed"] = True
        exc = rec["fut"].exception()
        if exc is not None:
            raise exc
        return rec["fut"].result()

    def _still_running(self, rec: dict) -> dict:
        took = int(time.monotonic() - rec["t0"])
        return {
            "still_running": True, "tool": rec["tool"], "elapsed_s": took,
            "note": ("{0} 還在 Aspen 裡進行（已 {1} 秒），沒有中斷。為了不讓用戶端判逾時，"
                     "每次呼叫最多等 {2:.0f} 秒就先回話。用**一模一樣的參數**再呼叫一次 "
                     "{0}，就會接著等、完成後拿到結果。不要重開或關閉專案 —— 那不會讓它"
                     "比較快，只會排在它後面。".format(rec["tool"], took, self.budget))}

    def _busy(self, rec: dict, tool: str) -> dict:
        took = int(time.monotonic() - rec["t0"])
        return {
            "busy": True, "running": rec["tool"], "elapsed_s": took,
            "error": ("上一個呼叫 {0} 還在 Aspen 裡進行（已 {1} 秒）。Aspen 一次只能做一件"
                      "事，這次的 {2} **沒有執行**。先用原本的參數再呼叫一次 {0} 拿到它的"
                      "結果，再做 {2}；不要重開專案。".format(rec["tool"], took, tool))}

    async def call(self, tool: str, args: dict):
        key = _call_key(tool, args)
        deadline = time.monotonic() + self.budget
        rec = self.inflight
        if rec is not None:
            if not await self._wait(rec["fut"], deadline):
                return self._still_running(rec) if rec["key"] == key else self._busy(rec, tool)
        earlier = self.unclaimed
        self.unclaimed = None
        if earlier is not None and earlier["key"] == key:
            # 就是上次沒等到的那一次：交出它的結果，不要再做一遍
            return self._outcome(earlier)

        rec = {"key": key, "tool": tool, "t0": time.monotonic()}
        rec["fut"] = self.submit(tool, args)
        self.inflight = rec
        rec["fut"].add_done_callback(lambda f, r=rec: self._finished(r, f))
        if not await self._wait(rec["fut"], deadline):
            return self._still_running(rec)
        result = self._outcome(rec)
        if self.unclaimed is rec:
            self.unclaimed = None
        if earlier is not None and isinstance(result, dict):
            try:
                said = self._outcome(earlier)
            except Exception as exc:                          # noqa: BLE001
                said = {"error": "{}: {}".format(type(exc).__name__, exc)}
            result = dict(result)
            result["earlier_call"] = {
                "tool": earlier["tool"],
                "note": "先前沒等到結果的那次 {} 已經完成，這是它的回覆".format(earlier["tool"]),
                "result": said}
        return result


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

    def _submit(name, arguments):
        return asyncio.get_running_loop().run_in_executor(
            pool, agent.call, name, arguments)

    gate = CallGate(_submit)

    @app.list_tools()
    async def _tools():
        # 只是向雲端要清單，不碰 Aspen —— 不放進 Aspen 執行緒，免得排在一次
        # 長的模擬後面，整個工具清單跟著逾時。
        tools = await asyncio.to_thread(agent.list_tools)
        return [Tool(name=t["name"], description=t["description"],
                     inputSchema=t["inputSchema"]) for t in tools]

    @app.call_tool()
    async def _call(name: str, arguments: dict):
        result = await gate.call(name, arguments or {})
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
    """繞過 MCP，直接驗證「代理 ↔ 雲端 ↔ 橋接 ↔ Aspen」整條迴圈。

    只開自己的一份 Aspen，結束時只關自己的那一份（shutdown → Bridge.close
    會確認那個行程真的結束）。這裡以前開頭和結尾各用映像檔名砍一次
    AspenPlus.exe —— 機器上每一份都中，包括使用者正開著、還沒存檔的模型。
    """
    import time

    others = _aspen_pids()
    if others:
        print("注意：已經有 {} 份 Aspen 開著（行程 {}）。自我測試會另外開一份，"
              "結束時只關自己開的那一份，不會動到它們。要測的檔案如果正開在"
              "其中一份裡，開檔會失敗 —— 那時請先把它關掉再測。".format(
                  len(others), "、".join(str(p) for p in others)))

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


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--selftest":
        raise SystemExit(selftest(sys.argv[2]))
    serve_mcp()
