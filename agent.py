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
import hashlib
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path

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


# ── 工作日誌 ────────────────────────────────────────────────────────
# MCP 走 stdin/stdout，這支程式印出來的東西沒有人看得到；用戶端說「逾時」
# 的時候，事後完全查不出是哪一支工具、卡在哪一個動作、等了多久。每一次
# 工具呼叫、每一個 Aspen 動作、每一次重試都記一行在本機。
#
# 一天一個檔（logs/agent-YYYY-MM-DD.log），保留 LOG_KEEP_DAYS 天。同一台電腦
# 可能同時跑好幾個 agent（每開一個對話就一個），每一行都帶行程編號。
# 不記 token。寫日誌失敗絕不影響工具呼叫。ASPEN_LOG=0 關閉。
LOG_DIR = Path(os.environ.get("ASPEN_LOG_DIR")
               or Path(os.path.dirname(os.path.abspath(__file__))) / "logs")
LOG_KEEP_DAYS = 14
_LOG_LOCK = threading.Lock()


def _fmt(value, limit: int = 300) -> str:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return "{:.2f}".format(value) if isinstance(value, float) else str(value)
    plain = isinstance(value, str)
    text = value if plain else json.dumps(value, ensure_ascii=False, default=str)
    text = text.replace("\r", " ").replace("\n", " ")
    if TOKEN and TOKEN in text:
        text = text.replace(TOKEN, "<token>")
    if len(text) > limit:
        text = text[:limit] + "...({} characters in total)".format(len(text))
    # 有空白的字串加引號，一行裡的 key=value 才分得開
    return json.dumps(text, ensure_ascii=False) if plain and (" " in text or not text) else text


def log(event: str, **fields) -> None:
    if os.environ.get("ASPEN_LOG") == "0":
        return
    try:
        now = time.time()
        line = "{}.{:03d} pid={} {}{}\n".format(
            time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(now)),
            int(now * 1000) % 1000, os.getpid(), event,
            "".join(" {}={}".format(k, _fmt(v)) for k, v in fields.items()
                    if v is not None))
        path = LOG_DIR / "agent-{}.log".format(time.strftime("%Y-%m-%d", time.localtime(now)))
        with _LOG_LOCK:
            LOG_DIR.mkdir(parents=True, exist_ok=True)
            with open(path, "a", encoding="utf-8") as fh:
                fh.write(line)
    except Exception:                                         # noqa: BLE001
        pass


def _prune_logs() -> None:
    """刪掉超過 LOG_KEEP_DAYS 天的日誌（只動 agent-*.log）。"""
    try:
        cutoff = time.time() - LOG_KEEP_DAYS * 86400
        for p in LOG_DIR.glob("agent-*.log"):
            if p.stat().st_mtime < cutoff:
                p.unlink()
    except Exception:                                         # noqa: BLE001
        pass


# ── 自己檢查有沒有新版，經學生同意才更新 ──────────────────────────────────
# 伺服器手上有一份 client_manifest.json（每個學生端檔案的指紋）。agent.py 啟動時、也隨時能問
# （MCP 工具 client_status）：本機的檔案和課程目前的版本一樣嗎。不一樣的話，Claude 先問學生要不要更新，
# 學生同意了才呼叫 client_update(confirm=true)：
#   - 這個資料夾是 git 複本、沒有本機修改 → git pull --ff-only；
#   - 否則從課程的公開倉庫下載有變的檔案；
#   兩條路都用伺服器 manifest 的指紋核對，核對不過就整個不換（倉庫還沒跟上課程版本時會這樣）。
# 換完要完全關閉 Claude 再開：正在跑的 agent.py 不會自己重新載入。
HERE = os.path.dirname(os.path.abspath(__file__))
RAW_BASE = os.environ.get("ASPEN_CLIENT_RAW_BASE",
                          "https://raw.githubusercontent.com/conlinkang/AspenPlusMCP_client/master/")
UPDATE_STATE: dict = {}


def _fp(path: str) -> str | None:
    """檔案的 sha1（完整 40 碼，換行字元統一成 LF）；檔案不在回 None。"""
    try:
        with open(path, "rb") as fh:
            return hashlib.sha1(fh.read().replace(b"\r\n", b"\n")).hexdigest()
    except OSError:
        return None


def fetch_manifest(timeout: float = 8.0) -> dict:
    """向課程伺服器要目前的版本表。"""
    return post("/client", {"token": TOKEN}, timeout=timeout, retry=False)


def client_status() -> dict:
    """本機的檔案和課程目前的版本一樣嗎。只查不改。"""
    try:
        m = fetch_manifest()
    except Exception as exc:                                  # noqa: BLE001
        return {"checked": False,
                "error": "Could not ask the course server for the current version: {}".format(exc)}
    files = m.get("files") or {"agent.py": m.get("agent"), "bridge.py": m.get("bridge")}
    changed = [name for name, want in files.items() if want and _fp(os.path.join(HERE, name)) != want]
    is_git = os.path.isdir(os.path.join(HERE, ".git"))
    out = {"checked": True, "up_to_date": not changed, "changed_files": changed,
           "yours": CLIENT_ID, "latest": "{}.{}".format(str(m.get("agent"))[:10], str(m.get("bridge"))[:10]),
           "git_clone": is_git}
    if changed:
        out["next"] = ("Ask the student whether to update now. If they agree, call client_update with confirm=true; "
                       "it {} and checks the result against the course version. They then close Claude completely "
                       "and reopen it.".format("runs git pull in this folder" if is_git else "downloads the changed files"))
    return out


def _run_git(*args, timeout: float = 60.0):
    return subprocess.run(["git", "-C", HERE, *args], capture_output=True, text=True, timeout=timeout,
                          creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def client_update(confirm=False) -> dict:
    """把本機的檔案換成課程目前的版本。一定要學生同意（confirm=true）。"""
    if confirm is not True:
        return {"updated": False,
                "error": "Not done: ask the student first, then call client_update with confirm=true."}
    try:
        m = fetch_manifest()
    except Exception as exc:                                  # noqa: BLE001
        return {"updated": False, "error": "Could not ask the course server for the current version: {}".format(exc)}
    files = m.get("files") or {"agent.py": m.get("agent"), "bridge.py": m.get("bridge")}
    changed = [name for name, want in files.items() if want and _fp(os.path.join(HERE, name)) != want]
    if not changed:
        return {"updated": False, "up_to_date": True, "note": "Already the course version; nothing to do."}
    needs_pip = "requirements.txt" in changed
    tail = {"restart_required": True,
            "next": "Tell the student to close Claude completely and open it again" + (
                "; requirements.txt changed, so also run: python -m pip install -r requirements.txt" if needs_pip else "")}

    if os.path.isdir(os.path.join(HERE, ".git")):
        try:
            dirty = _run_git("status", "--porcelain", "--untracked-files=no").stdout.strip()
            if dirty:
                return {"updated": False, "method": "git pull",
                        "error": "This folder is a git copy with local changes, so git pull was not run (it could overwrite "
                                 "them). Save or discard the changes ({}), then try again.".format(dirty.replace("\n", "; ")[:200])}
            old = _run_git("rev-parse", "HEAD").stdout.strip()
            pulled = _run_git("pull", "--ff-only")
            if pulled.returncode != 0:
                return {"updated": False, "method": "git pull", "error": "git pull failed: {}".format(
                    (pulled.stderr or pulled.stdout).strip()[:300])}
            still = [n for n, want in files.items() if want and _fp(os.path.join(HERE, n)) != want]
            if still:
                if old:
                    _run_git("reset", "--hard", old)
                return {"updated": False, "method": "git pull",
                        "error": "After git pull these files still differ from the course version: {}. The repository "
                                 "is not up to date with the course server yet; nothing was changed. Ask the "
                                 "administrator.".format(", ".join(still))}
            return dict({"updated": True, "method": "git pull", "files": changed}, **tail)
        except Exception as exc:                              # noqa: BLE001
            return {"updated": False, "method": "git pull", "error": "{}: {}".format(type(exc).__name__, exc)}

    # git 複本以外：下載有變的檔案，全部核對過才換
    import tempfile
    got = {}
    try:
        for name in changed:
            req = urllib.request.Request(RAW_BASE + name, headers={"User-Agent": "aspen-agent"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = resp.read()
            if hashlib.sha1(data.replace(b"\r\n", b"\n")).hexdigest() != files[name]:
                return {"updated": False, "method": "download",
                        "error": "{} from the repository does not match the course version, so nothing was changed. The "
                                 "repository is not up to date with the course server yet; ask the "
                                 "administrator.".format(name)}
            got[name] = data
    except Exception as exc:                                  # noqa: BLE001
        return {"updated": False, "method": "download", "error": "Download failed: {}: {}".format(type(exc).__name__, exc)}
    backup = os.path.join(HERE, "client_backup", time.strftime("%Y%m%d-%H%M%S"))
    os.makedirs(backup, exist_ok=True)
    for name, data in got.items():
        target = os.path.join(HERE, name)
        if os.path.exists(target):
            shutil.copy2(target, os.path.join(backup, name))
        tmp = target + ".new"
        with open(tmp, "wb") as fh:
            fh.write(data)
        os.replace(tmp, target)
    return dict({"updated": True, "method": "download", "files": changed, "backup": backup}, **tail)


def _startup_check() -> None:
    """agent 一啟動就在背景問一次，結果寫進日誌；查不到不影響任何事。"""
    try:
        UPDATE_STATE.update(client_status())
        log("client_status", **{k: v for k, v in UPDATE_STATE.items() if k != "next"})
    except Exception as exc:                                  # noqa: BLE001
        log("client_status_failed", exception="{}: {}".format(type(exc).__name__, exc))


def _client_version() -> str | None:
    """學生端資料夾的 git 提交（用 git pull 更新的話才有）。"""
    try:
        r = subprocess.run(
            ["git", "-C", os.path.dirname(os.path.abspath(__file__)), "log", "-1",
             "--format=%h %cd", "--date=format:%Y-%m-%d %H:%M"],
            capture_output=True, text=True, timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return r.stdout.strip() or None
    except Exception:                                         # noqa: BLE001
        return None


def _summary(result) -> dict:
    """工具回覆裡值得記的幾個欄位，加上整份的大小。"""
    if not isinstance(result, dict):
        return {"type": type(result).__name__}
    keep = ("error", "status", "ran", "converged", "opened", "still_running",
            "busy", "running", "n_applied", "n_rejected", "complete", "interrupted",
            "engine_stuck")
    out = {k: result[k] for k in keep if k in result}
    if "earlier_call" in result:
        out["earlier_call"] = (result["earlier_call"] or {}).get("tool")
    out["chars"] = len(json.dumps(result, ensure_ascii=False, default=str))
    return out


def _op_brief(op: str, args) -> str:
    args = args or {}
    if op == "batch":
        return "{} operations".format(len(args.get("ops") or []))
    return str(args.get("path") or args.get("action") or "")[:160]


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
            "This account's tool call quota for this period is used up "
            "({}). Retrying will not succeed, and this is not a problem "
            "with the model or Aspen; wait until the reset time, and "
            "contact the course teaching assistant if more quota is "
            "needed.".format(said or "usage limit reached"))
    if exc.code == 429:
        return RuntimeError("{}. Call again in about a minute.".format(
            said or "too many calls; wait a moment and try again"))
    if exc.code == 401:
        return RuntimeError("{} (check that ASPEN_TOKEN is typed correctly)".format(
            said or "invalid credentials or the account is suspended"))
    return RuntimeError("The cloud server responded HTTP {}: {}".format(exc.code, said or exc.reason))


def client_id() -> str:
    """這份 agent.py 和旁邊的 bridge.py 的指紋（sha1 前 10 碼，換行字元統一成 LF 再算）。
    每次請求都帶給雲端，雲端用它認得學生端的檔案是不是課程目前的版本，太舊就提醒去下載。"""
    here = os.path.dirname(os.path.abspath(__file__))
    out = []
    for name in ("agent.py", "bridge.py"):
        try:
            with open(os.path.join(here, name), "rb") as fh:
                out.append(hashlib.sha1(fh.read().replace(b"\r\n", b"\n")).hexdigest()[:10])
        except OSError:
            out.append("none")
    return ".".join(out)


CLIENT_ID = client_id()


def post(path: str, payload: dict, timeout: float = 60.0, retry: bool = True) -> dict:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        CLOUD + path, data=body,
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer " + TOKEN,
                 "X-Aspen-Client": CLIENT_ID})
    last = None
    for wait in (RETRY_S if retry else ()) + (None,):
        t0 = time.monotonic()
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                out = json.loads(resp.read().decode("utf-8"))
            took = time.monotonic() - t0
            if took > 5:
                log("cloud_slow", path=path, s=took)
            return out
        except urllib.error.HTTPError as exc:
            # 雲端有回應、只是拒絕 —— 重試沒有意義，把它的說明轉出來
            err = _cloud_error(exc)
            log("cloud_http_error", path=path, code=exc.code, said=str(err))
            raise err from None
        except urllib.error.URLError as exc:
            reason = getattr(exc, "reason", exc)
            refused = (isinstance(reason, ConnectionRefusedError)
                       or "10061" in str(reason)
                       or "refused" in str(reason).lower())
            if not refused or wait is None:
                log("cloud_unreachable", path=path, refused=refused,
                    reason=str(reason), s=time.monotonic() - t0)
                if refused:
                    raise ConnectionError(
                        "Cannot reach the cloud server {}: the "
                        "connection was refused, retried for {} seconds. "
                        "The server may not be running or may be "
                        "restarting; try again later, or ask the "
                        "administrator to check that port 8787 is "
                        "listening.".format(CLOUD, int(sum(RETRY_S)))) from exc
                raise
            log("cloud_refused_retry", path=path, wait_s=wait)
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
        t0 = time.monotonic()
        n_ops = 0
        self.rounds = 0
        log("call_start", session=self.session, tool=tool, args=args)
        try:
            reply = post("/call", {"session": self.session, "tool": tool,
                                   "args": args, "token": TOKEN})
            while reply.get("type") == "ops":
                self.rounds += 1
                if self.rounds > MAX_ROUNDS:
                    log("call_end", tool=tool, s=time.monotonic() - t0,
                        rounds=self.rounds, ops=n_ops, error="too_many_rounds")
                    return {"error": "More than {} rounds; the cloud server may be looping".format(MAX_ROUNDS)}
                results, names, t_round = [], [], time.monotonic()
                for item in reply["ops"]:
                    self.ops_sent += 1
                    n_ops += 1
                    op, op_args = item.get("op"), item.get("args")
                    names.append(op)
                    t_op = time.monotonic()
                    res = self.bridge.execute(op, op_args)
                    results.append(res)
                    took = time.monotonic() - t_op
                    ok = isinstance(res, dict) and res.get("ok")
                    # 每個動作都記會太長：慢的、失敗的才單獨記一行
                    if took >= 0.5 or not ok:
                        log("op", tool=tool, round=self.rounds, op=op,
                            target=_op_brief(op, op_args), s=took, ok=bool(ok),
                            error=None if ok else (res or {}).get("error"))
                log("round", tool=tool, round=self.rounds, ops=",".join(map(str, names)),
                    s=time.monotonic() - t_round)
                reply = post("/result", {"session": self.session,
                                         "results": results, "token": TOKEN})
        except Exception as exc:
            log("call_failed", tool=tool, s=time.monotonic() - t0, rounds=self.rounds,
                ops=n_ops, exception="{}: {}".format(type(exc).__name__, exc))
            raise
        result = reply.get("result", {})
        log("call_end", tool=tool, s=time.monotonic() - t0, rounds=self.rounds,
            ops=n_ops, result=_summary(result))
        return result

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
            if rec.get("abandoned"):        # 用戶端已經拿到 still_running、走掉了
                log("gate_finished_unclaimed", tool=rec["tool"],
                    s=rec["done_at"] - rec["t0"])

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
        rec["abandoned"] = True
        log("gate_still_running", tool=rec["tool"], elapsed_s=took)
        return {
            "still_running": True, "tool": rec["tool"], "elapsed_s": took,
            "note": ("{0} is still running in Aspen ({1} s so far) and "
                     "has not been interrupted. So that the client does "
                     "not time out, each call waits at most {2:.0f} "
                     "seconds before replying. Call {0} again with "
                     "**exactly the same arguments** to keep waiting and "
                     "get the result when it finishes. Do not reopen or "
                     "close the project: that does not make it faster, "
                     "it only queues behind it.".format(rec["tool"], took, self.budget))}

    def _busy(self, rec: dict, tool: str) -> dict:
        took = int(time.monotonic() - rec["t0"])
        log("gate_busy", refused=tool, running=rec["tool"], elapsed_s=took)
        return {
            "busy": True, "running": rec["tool"], "elapsed_s": took,
            "error": ("The previous call {0} is still running in Aspen "
                      "({1} s so far). Aspen can do only one thing at a "
                      "time, so this {2} **was not executed**. First "
                      "call {0} again with its original arguments to get "
                      "its result, then do {2}; do not reopen the "
                      "project.".format(rec["tool"], took, tool))}

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
            log("gate_delivered_earlier", tool=tool)
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
            log("gate_attached_earlier", tool=tool, earlier=earlier["tool"])
            result["earlier_call"] = {
                "tool": earlier["tool"],
                "note": "The earlier call to {} whose result was not awaited has finished; this is its reply".format(earlier["tool"]),
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

    _prune_logs()
    threading.Thread(target=_startup_check, name="client-check", daemon=True).start()
    log("agent_start", cloud=CLOUD, client=_client_version(),
        python=sys.version.split()[0], budget_s=CALL_BUDGET_S,
        parent_pid=os.getppid())
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
        try:
            tools = await asyncio.to_thread(agent.list_tools)
        except Exception as exc:
            log("list_tools_failed", exception="{}: {}".format(type(exc).__name__, exc))
            raise
        log("list_tools", n=len(tools))
        return [Tool(name=t["name"], description=t["description"],
                     inputSchema=t["inputSchema"]) for t in tools] + [
            Tool(name=t["name"], description=t["description"], inputSchema=t["inputSchema"]) for t in LOCAL_TOOLS]

    @app.call_tool()
    async def _call(name: str, arguments: dict):
        # 用戶端實際等了多久：和 call_end 的秒數不同時，差的就是排隊或 still_running
        t0 = time.monotonic()
        if name in LOCAL_HANDLERS:
            # 不經雲端、不碰 Aspen 的本地工具：不排在 Aspen 執行緒後面（長模擬跑著也問得到有沒有新版）
            result = await asyncio.to_thread(LOCAL_HANDLERS[name], arguments or {})
            log("local_tool", tool=name, result=_summary(result))
            return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]
        try:
            result = await gate.call(name, arguments or {})
        except Exception as exc:
            log("mcp_error", tool=name, s=time.monotonic() - t0,
                exception="{}: {}".format(type(exc).__name__, exc))
            raise
        log("mcp_reply", tool=name, s=time.monotonic() - t0, result=_summary(result))
        return [TextContent(type="text",
                            text=json.dumps(result, ensure_ascii=False, indent=2))]

    async def main():
        async with stdio_server() as (r, w):
            await app.run(r, w, app.create_initialization_options())

    try:
        asyncio.run(main())
    except BaseException as exc:
        log("agent_crashed", exception="{}: {}".format(type(exc).__name__, exc))
        raise
    finally:
        log("agent_stop", ops_sent=agent.ops_sent,
            busy_with=(gate.inflight or {}).get("tool"))
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
                  len(others), ", ".join(str(p) for p in others)))

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
                 lambda r: "opened={}, unit table {} entries".format(
                     r.get("opened"), r.get("units_loaded")))
        if not r.get("opened"):
            print("    ", r.get("error"))
            return 1

        step(3, "list_model", {},
             lambda r: "blocks {}, streams {}".format(r.get("blocks"), r.get("streams")))

        r = step(4, "get_specs",
                 {"target": "block", "name": "COL1", "only_set": True},
                 lambda r: "{} fields set".format(r.get("count")))
        for k in ("NSTAGE", "BASIS_RR", "PRES1", "BASIS_D"):
            if k in (r.get("specs") or {}):
                v = r["specs"][k]
                print("      {:<10} {:<10} {}".format(
                    k, v["value"], v["unit"] or ""))

        step(5, "set_specs",
             {"target": "block", "name": "COL1",
              "specs": {"NSTAGE": 25, "PRES1": {"value": 1.5, "unit": "bar"}}},
             lambda r: "{} written, {} rejected {}".format(
                 r.get("n_applied"), r.get("n_rejected"), r.get("rejected") or ""))

        step(6, "set_specs",
             {"target": "stream", "name": "FEED", "specs": {"TEMP": 60}},
             lambda r: "{} rejected: {}".format(
                 r.get("n_rejected"),
                 [v.get("detail") for v in (r.get("rejected") or {}).values()]))

        print("\n本次共送出 {} 個動作".format(agent.ops_sent))
        print("學生端從頭到尾不知道：'bar' 對應哪個索引、寫入後要複查、"
              "漏單位要擋下來 —— 那些都在雲端。")
        return 0
    finally:
        agent.shutdown()


LOCAL_TOOLS = (
    {"name": "client_status",
     "description": ("[Client] Check whether the agent.py/bridge.py on this computer are the course's current version. "
                     "Only asks the course server; changes nothing. Call it when a tool reply says client_update or "
                     "client_outdated, when a long call keeps timing out, or when the student asks about updates."),
     "inputSchema": {"type": "object", "properties": {}}},
    {"name": "client_update",
     "description": ("[Client] Update agent.py/bridge.py on this computer to the course's current version. Ask the student "
                     "first and pass confirm=true ONLY after they agreed. It runs git pull when this folder is a git copy "
                     "without local changes, otherwise downloads the changed files; either way the result is checked "
                     "against the course server and nothing is changed if it does not match. Afterwards the student "
                     "must close Claude completely and open it again."),
     "inputSchema": {"type": "object", "properties": {
         "confirm": {"type": "boolean", "description": "true only after the student said yes"}},
         "required": ["confirm"]}},
)
LOCAL_HANDLERS = {"client_status": lambda args: client_status(),
                  "client_update": lambda args: client_update(args.get("confirm"))}


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--selftest":
        raise SystemExit(selftest(sys.argv[2]))
    if len(sys.argv) > 1 and sys.argv[1] in ("--check-update", "--update"):
        # 自己在終端機執行就是同意：python agent.py --check-update 只查；--update 真的更新
        out = client_update(True) if sys.argv[1] == "--update" else client_status()
        print(json.dumps(out, ensure_ascii=False, indent=2))
        raise SystemExit(0 if out.get("updated") or out.get("up_to_date") or out.get("checked") else 1)
    serve_mcp()
