"""本地橋接層。

學生電腦上唯一會執行的東西。收到什麼做什麼，不做任何判斷 ——
不重試、不補單位、不驗證結果、不知道自己在做什麼。那些全是雲端的責任。

被逆向出來只會看到「怎麼讀寫 Aspen 的資料節點」，
那是查官方文件就能寫的東西，不是護城河。

必須用有 pywin32 的直譯器執行：D:\\Aspen_MCP\\.venv\\Scripts\\python.exe
規格見 bridge_spec.md，實測數據見 spike_conclusion.md。
"""
from __future__ import annotations

import glob
import io
import os
import subprocess
import threading
import time
from typing import Any

import win32com.client as win32


class BridgeError(Exception):
    def __init__(self, code: str, detail: str = ""):
        super().__init__("{}: {}".format(code, detail))
        self.code = code
        self.detail = detail


def _ok(data: Any = None) -> dict:
    return {"ok": True, "data": data, "error": None}


# history 檔裡代表「這次跑不起來」的字樣。抓這幾行回報給使用者，
# 比任何我們自己編的訊息都準確 —— 那是 Aspen 自己講的話。
_FATAL_MARKS = (
    "ERROR WHILE CHECKING INPUT SPECIFICATIONS",
    "SIMULATION PROGRAM CANNOT BE EXECUTED",
    "DUE TO PREVIOUS SEVERE ERROR",
    "SEVERE ERROR",
    "**  ERROR",
)


def _extract_errors(text: str, limit: int = 25) -> list:
    """從 history 檔挑出錯誤段落。抓到標記後連同後續幾行一起帶走 ——
    Aspen 的錯誤訊息是多行的，只抓標題那一行看不出原因。"""
    if not text:
        return []
    lines = text.splitlines()
    out, i = [], 0
    while i < len(lines) and len(out) < limit:
        if any(mark in lines[i].upper() for mark in _FATAL_MARKS):
            for line in lines[i:i + 8]:
                stripped = line.strip()
                if stripped and stripped not in out:
                    out.append(stripped)
            i += 8
            continue
        i += 1
    return out[:limit]


def _stamps(path: str) -> dict:
    """同名（不論副檔名）的檔案目前的修改時間。"""
    stem = os.path.splitext(path)[0]
    out = {}
    for p in glob.glob(stem + ".*"):
        try:
            out[os.path.basename(p)] = os.path.getmtime(p)
        except OSError:
            pass
    return out


def _written(before: dict, after: dict) -> list:
    """這次呼叫實際新增或改寫了哪些檔案。

    只看「同名檔案存不存在」是不夠的：先前存過一次 .bkp，之後匯出報表
    什麼都沒寫，卻因為那個 .bkp 還在而被當成匯出成功。
    比對前後的修改時間才問得出「這一次到底寫了什麼」。

    純粹陳述事實，不做判斷 —— 「沒寫出來代表要先跑模擬」是雲端的知識。
    """
    return sorted(name for name, ts in after.items()
                  if before.get(name) != ts)


# 有幾個錯誤碼在程式裡被丟出來的地方太多，逐一寫訊息一定會漏。實測結果
# 是呼叫端收到「NOT_CONNECTED: 」—— 一個冒號後面什麼都沒有，看不出要做
# 什麼。這裡給預設值，呼叫點自己有話講時不會被蓋掉。
_DEFAULT_DETAIL = {
    "NOT_CONNECTED": "目前沒有連著的 Aspen。先用 open_project 開檔，"
                     "或用 new_project 建一個空白專案。",
}


def _err(code: str, detail: str = "") -> dict:
    text = str(detail)[:500]
    if not text.strip():
        text = _DEFAULT_DETAIL.get(code, text)
    return {"ok": False, "data": None,
            "error": {"code": code, "detail": text}}


# ── 有沒有可以操作的桌面 ────────────────────────────────────────────
# UI 自動化（電解質精靈、經濟分析）要有一個真的在渲染的桌面。遠端連線
# 中斷、或畫面鎖住時，視窗還在、行程還活著，但 UIA 找不到任何控制項。
# 實測：這種狀態下精靈十次全敗，每次 5–51 秒，錯誤訊息只說「找不到某個
# 下拉選單」，完全沒有指向真正的原因。先問一句就能省掉這些。
def _interactive_desktop() -> tuple:
    """(能不能操作桌面, 說明)。查不出來就當作可以，不擋路。"""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        # OpenInputDesktop 拿的是「目前正在接收輸入的那個桌面」。
        # session 被中斷或螢幕鎖住時，這裡就拿不到。
        handle = user32.OpenInputDesktop(0, False, 0x0001)  # DESKTOP_READOBJECTS
        if handle:
            user32.CloseDesktop(handle)
            return True, "ok"
        remote = bool(user32.GetSystemMetrics(0x1000))      # SM_REMOTESESSION
        return False, ("沒有可以操作的桌面：目前這個工作階段沒有在接收輸入"
                       "（遠端連線已中斷、或畫面鎖住了）。"
                       "UI 自動化在這種狀態下一定失敗。"
                       "請把遠端桌面接回來、或在本機解鎖畫面後重試。"
                       + ("（偵測到這是遠端連線工作階段）" if remote else ""))
    except Exception:
        return True, "檢查不出來，當作可以用"


# 開檔最多試幾次、每次之間等多久。模組層級的常數，測試可以蓋掉。
_OPEN_ATTEMPTS = 3
_RETRY_SLEEP = 3.0

# 送出停止要求之後，最多等引擎幾秒才承認它沒停下來。
STOP_WAIT_S = 10.0


def _build_typelib(app) -> str | None:
    """把 early-bound 型別庫建起來（相當於手動跑 makepy）。

    WithEvents 要有型別庫才掛得上去。快取沒建、或被建壞的時候（這個專案
    遇過 0 byte 的損毀檔），會丟「This COM object can not automate the
    makepy process」。這不是橋接程式版本舊，重連 MCP 也不會變好 ——
    要真的去把型別庫生出來。成功回 None，失敗回原因。
    """
    try:
        from win32com.client import gencache
        gencache.EnsureDispatch(app)
        return None
    except Exception as exc:
        return repr(exc)[:200]


# ── 殘留的 Aspen 行程 ───────────────────────────────────────────────
# Aspen 掛掉、或這支程式被砍掉來不及關檔時，AspenPlus.exe 會留下來，
# 而且鎖著剛才那個 .bkp。下一次開檔就報「Unable to open file」，訊息
# 裡完全看不出原因 —— 使用者只能自己去工作管理員找。
def _aspen_pids() -> list:
    """目前機器上所有 AspenPlus.exe 的行程編號。查不到就回空清單。"""
    try:
        # 不用 text=True：tasklist 的輸出是主控台編碼（中文 Windows 是
        # cp950），學生若設了 PYTHONUTF8=1 就會在讀取執行緒裡炸
        # UnicodeDecodeError，然後這裡回空清單 —— 「沒有殘留行程」變成假話。
        # 行程編號是 ASCII，用 replace 解碼就夠。
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq AspenPlus.exe", "/NH", "/FO",
             "CSV"], capture_output=True, timeout=15)
    except Exception:
        return []
    pids = []
    text = (out.stdout or b"").decode("utf-8", errors="replace")
    for line in text.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[0].lower().startswith("aspenplus"):
            try:
                pids.append(int(parts[1]))
            except ValueError:
                pass
    return pids


def _locked_by(path: str) -> str | None:
    """檔案被別的行程開著就回一句說明，沒有就回 None。

    Windows 下 Aspen 開著的 .bkp 會被獨佔；最常見的兇手是殘留的
    AspenPlus.exe。試著以寫入模式開一下，開不了就是被鎖。
    """
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r+b"):
            return None
    except PermissionError:
        others = _aspen_pids()
        tail = (" 目前還有 {} 個 AspenPlus.exe 在跑（PID {}），很可能就是它們。"
                .format(len(others), "、".join(str(p) for p in others))
                if others else "")
        return "檔案被別的程式開著，無法寫入：{}。{}".format(path, tail)
    except OSError:
        return None


def _pid_alive(pid: int) -> bool:
    return pid in _aspen_pids()


def _kill_pid(pid: int) -> None:
    try:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                       capture_output=True, timeout=15)
    except Exception:
        pass


# Control Panel 訊息最多留這麼多行。一次長模擬可能吐上萬行，全留著會把
# 記憶體和回傳都撐爆；真正有用的是最後那一段和「TERMINAL ERROR」那幾行。
CP_MAX_LINES = 4000


class _ControlPanelEvents:
    """接 Aspen 的 OnControlPanelMessage 事件，把每一行收進 sink。

    WithEvents 會自己 new 這個類別（不能帶參數），所以 sink 事後才掛上去。
    收到的是 Aspen 在 Control Panel 視窗裡講的原話 —— 「NO COMPONENTS HAVE
    BEEN DEFINED」「SIMULATION WILL NOT BE EXECUTED」這種。這些話**不會**
    出現在 Run-Status 節點或 history 檔裡：輸入被拒絕時 Aspen 連 Run-Status
    都不建，事後去讀什麼都讀不到。只有事件當下接住才拿得到。
    """

    def __init__(self):
        self.sink = None

    def OnControlPanelMessage(self, clear, msg):
        sink = self.sink
        if sink is None:
            return
        try:
            if clear:
                sink.append("[CONTROL PANEL CLEARED]")
            if msg is not None:
                text = str(msg).rstrip()
                if text.strip():
                    sink.append(text)
            if len(sink) > CP_MAX_LINES:
                del sink[:len(sink) - CP_MAX_LINES]
        except Exception:
            pass


class Bridge:
    """20 個通用動作。不含業務邏輯。"""

    def __init__(self) -> None:
        self.app = None
        self._ui_cache: dict = {}      # 控制項 handle 表（快取，非邏輯）
        self._run_started: float | None = None
        self._run_done: threading.Event | None = None
        self._run_error: str | None = None
        self._path: str | None = None      # 目前開著的檔，用來找 history 檔
        self._pid: int | None = None       # 自己這一份 Aspen 的行程編號
        self._cp_messages: list = []        # 這一次執行的 Control Panel 原話
        self._cp_handler = None             # WithEvents 的事件物件，跑完就放掉
        self._cp_capture: str = "not_started"
        self.op_count = 0

    # ══ 連線與檔案 ══════════════════════════════════════════════════
    def connect(self, version: str | None = None) -> dict:
        """已經連著就直接沿用，不重新 Dispatch。

        open_project 的開檔順序每次都會送一次 connect（ELECTROLYTE 精靈
        的流程還會連續開好幾次檔，確保 UI 自動化拿到真實檔名）。
        Dispatch() 不是「接上既有那個」，是「生一個新的」—— 沒有這道
        存活檢查，每次 open_project 都會多開一個 AspenPlus.exe，
        舊的那個（可能剛被精靈改過、還沒存檔）就變成孤兒程序：
        使用者在畫面上看到的是舊的，這裡讀到的卻是重新從磁碟載入的新的。
        """
        if self.app is not None:
            try:
                self.app.Visible  # noqa: B018  # 純粹用來確認 COM 物件還活著
                return _ok({"progid": "reused"})
            except Exception:
                self.app = None    # 舊的死了，才真的需要開新的
        progid = "Apwn.Document.{}.0".format(version) if version else "Apwn.Document"
        try:
            self.app = win32.Dispatch(progid)
            # 記下行程編號。關檔時要確認這一份真的結束了，開檔失敗時也
            # 要能把「別人留下的」和「自己這一份」分開講。
            try:
                self._pid = int(self.app.ProcessId)
            except Exception:
                self._pid = None
            return _ok({"progid": progid, "pid": self._pid})
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def create_file(self) -> dict:
        """開一個空白模型（File → New）。

        InitNew2() 只在這一份 Aspen 還沒開過模型時有用。已經開著模型再叫
        它，實測丟 com_error 2001「Unexpected Error」—— 舊 MCP 把這個
        例外吞掉回 False，呼叫端沒檢查，於是「新專案」其實是原來那個
        模型繼續用，後面加的元件、反應全疊在舊模型上。而 app.Close() 只會
        把 COM 連線切斷、行程留著（實測），也救不回來。

        唯一可靠的做法是換一份新的 Aspen：關掉自己這一份（close() 會確認
        行程真的結束）、再 Dispatch 一份，然後 InitNew2。回傳裡講清楚有
        沒有換過行程、前一個模型的路徑是什麼 —— 那個模型是連存都沒存就
        關掉的，雲端要把這件事告訴使用者。
        """
        if self.app is None:
            return _err("NOT_CONNECTED")
        try:
            self.app.InitNew2()
            self._path = None
            return _ok({"recycled": False, "pid": self._pid})
        except Exception as first:
            previous_path, old_pid = self._path, self._pid
            closed = self.close()
            reconnected = self.connect()
            if not reconnected.get("ok"):
                return _err("COM_ERROR",
                            "InitNew2 失敗（{}），重開 Aspen 也失敗：{}".format(
                                repr(first), reconnected.get("error")))
            try:
                self.app.InitNew2()
            except Exception as second:
                return _err("COM_ERROR", "重開一份 Aspen 之後 InitNew2 仍失敗：{}".format(repr(second)))
            self._path = None
            return _ok({"recycled": True, "pid": self._pid, "previous_pid": old_pid,
                        "previous_path": previous_path,
                        "previous_had_to_force_close": (closed.get("data") or {}).get("had_to_force_close")})

    def open_file(self, path: str, mode: str = "archive") -> dict:
        """開檔。mode 決定用哪個 COM 方法，由雲端指定，橋接層不替它選。

        archive → InitFromArchive，用於 .bkp/.apw 壓縮檔
        file    → InitFromFile2，用於 .apwz 等

        兩者不等價：實測 InitFromFile2 開同一個 .bkp 後，部分字串欄位
        （例如物流的 LSSOURCE）讀回 None，而 InitFromArchive 讀回空字串。
        這種差異不會報錯，只會讓資料悄悄變樣，所以必須明確指定。
        """
        if self.app is None:
            return _err("NOT_CONNECTED", "需先 connect")
        loader = {"archive": "InitFromArchive", "file": "InitFromFile2"}.get(mode)
        if loader is None:
            return _err("BAD_ARG", "mode 只能是 archive 或 file")
        if not os.path.isfile(path):
            folder = os.path.dirname(os.path.abspath(path))
            hint = ("目錄也不存在" if not os.path.isdir(folder) else
                    "目錄在，但沒有這個檔名（大小寫與副檔名 .bkp/.apw/.apwz 都要對）")
            return _err("FILE_NOT_FOUND",
                        self._open_failure("找不到檔案：{}（{}）".format(path, hint)))
        # 開檔會間歇性失敗（「Unable to open file」2041），同一個檔案等幾秒
        # 再開就成功 —— 實測重試 100% 恢復。這是少數該由橋接層自己處理的
        # 重試：檔案存在性已經確認過，剩下的是 Aspen 自己還沒準備好。
        # 回報試了幾次，免得真的壞掉時被這層重試遮住。
        last = None
        for attempt in range(1, _OPEN_ATTEMPTS + 1):
            try:
                getattr(self.app, loader)(path)
                self._path = path
                data = {"path": path, "loader": loader}
                if attempt > 1:
                    data["attempts"] = attempt
                return _ok(data)
            except Exception as exc:
                last = exc
                if attempt < _OPEN_ATTEMPTS:
                    time.sleep(_RETRY_SLEEP)
        return _err("COM_ERROR", "{}（已重試 {} 次）".format(
            self._open_failure(repr(last)), _OPEN_ATTEMPTS))

    def _open_failure(self, detail: str) -> str:
        """開檔失敗時，把能查到的線索一起帶回去。

        Aspen 自己的 FailedToOpenDescription 有時是空的（實測「檔案不
        存在」就是空的），所以不能只靠它。真正常見的原因是別的
        AspenPlus.exe 還鎖著那個檔，那個用行程清單就看得出來。
        """
        parts = [detail]
        for attr in ("FailedToOpenDescription", "FailedToOpenKey"):
            try:
                said = getattr(self.app, attr)
            except Exception:
                continue
            if said:
                parts.append("Aspen 說：{}".format(said))
        others = [p for p in _aspen_pids() if p != self._pid]
        if others:
            parts.append(
                "另外有 {} 個 AspenPlus.exe 還在跑（PID {}）。這類殘留會鎖住 "
                ".bkp，是開檔失敗最常見的原因。確認畫面上沒有你自己開著、"
                "還沒存檔的 Aspen 之後，可以用 "
                "Stop-Process -Id {} -Force 清掉再重試。".format(
                    len(others), "、".join(str(p) for p in others),
                    ",".join(str(p) for p in others)))
        return "；".join(parts)

    def save(self, path: str | None = None, overwrite: bool | None = None,
             export_type: int | None = None) -> dict:
        """把模型寫到檔案。給 path 就另存新檔，再給 export_type 就是匯出。

        存檔和匯出走的是不同的 COM 方法（`SaveAs` 對 `Export`），但都是
        「把這個模型寫成一個檔案」，所以是同一個動作的兩種模式，不另開一個。
        `export_type` 是 Aspen 的格式編號，編號對應哪種格式屬於雲端的知識。

        `overwrite` 省略時只傳一個引數給 `SaveAs`，和原始程式一致。
        先前這裡固定傳 `SaveAs(path, True)` —— 多一個引數就是不同的呼叫，
        這種差異不會報錯，只會在某些情況下產生不一樣的檔案。
        要覆寫就明講。
        """
        if self.app is None:
            return _err("NOT_CONNECTED")
        # 目錄不存在時 Aspen 只回「Aspen.Unknown No message available」，
        # 檔被別人開著時回「Unable to open file」——兩個都看不出原因，
        # 但兩個都能在丟給 COM 之前查出來。
        if path:
            folder = os.path.dirname(os.path.abspath(path))
            if not os.path.isdir(folder):
                return _err("DIR_NOT_FOUND",
                            "目錄不存在：{}。先建立目錄，或改存到既有的資料夾。"
                            .format(folder))
            if os.path.exists(path) and not os.access(path, os.W_OK):
                return _err("FILE_READONLY", "檔案是唯讀的：{}".format(path))
            locked = _locked_by(path)
            if locked:
                return _err("FILE_LOCKED", locked)
        try:
            if export_type is not None:
                if not path:
                    return _err("BAD_ARG", "匯出必須指定 path")
                before = _stamps(path)
                self.app.Export(int(export_type), path)
                # 匯出成功與否要看檔案在不在，不能只看有沒有丟例外。
                # 實測：報表格式在模型還沒跑過時，Export 回報成功卻什麼都沒寫出來。
                # 而且 Aspen 會自己把副檔名換成該格式的，呼叫端要的檔名不一定算數。
                return _ok({"path": path, "export_type": int(export_type),
                            "written": _written(before, _stamps(path))})
            if path:
                if overwrite is None:
                    self.app.SaveAs(path)
                else:
                    self.app.SaveAs(path, bool(overwrite))
                self._path = path
            else:
                self.app.Save()
            return _ok({"path": path})
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def close(self) -> dict:
        """關掉這一份 Aspen，並確認行程真的結束了。

        實測 Close() 會讓 AspenPlus.exe 一起結束，所以正常路徑不會留下
        孤兒。會留下的是異常路徑：Aspen 自己掛掉、或這支程式被砍時來不及
        走到這裡。那些殘留會鎖住 .bkp，下一次 open_project 就報
        「Unable to open file」，而錯誤訊息完全看不出原因。

        所以這裡多做一件事：Close() 之後回頭確認自己那個行程編號真的
        不在了。還在就把它收掉 —— 收的是自己開的那一份，不會動到使用者
        自己開著的 Aspen 視窗。
        """
        pid = self._pid
        self._detach_control_panel()
        if self.app is not None:
            try:
                self.app.Close()
            except Exception:
                pass
            self.app = None
        self._ui_cache.clear()
        self._pid = None
        leftover = False
        if pid and _pid_alive(pid):
            leftover = True
            _kill_pid(pid)
        return _ok({"pid": pid, "had_to_force_close": leftover})

    # 應用層級的 COM 屬性。原本有一個專用的 show_gui，但那把「要不要開 GUI」
    # 這個決定寫死在本地。改成通用讀寫後，決定權回到雲端，動作數也沒增加。
    _APP_PROPS = ("Visible", "SuppressDialogs")

    def app_prop(self, name: str, value=None) -> dict:
        if self.app is None:
            return _err("NOT_CONNECTED")
        if name not in self._APP_PROPS:
            return _err("BAD_ARG", "只允許 {}".format(", ".join(self._APP_PROPS)))
        try:
            if value is None:
                return _ok({"name": name, "value": getattr(self.app, name)})
            setattr(self.app, name, value)
            return _ok({"name": name, "value": value})
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    # ══ 樹狀存取 ════════════════════════════════════════════════════
    def _node(self, path: str, child: str | None = None):
        """找一個節點。

        child 不為空時，先用路徑找到父節點，再從它的子節點裡**按名稱列舉**
        取出目標。

        為什麼需要這條路：Aspen 的 FindNode **解析不了最後一段含空白的路徑**。
        純物性參數的儲存格叫 "TC NBUTYLVA"、反應計量的叫 "H2O MIXED"，
        children 明明列得出來，FindNode 卻回 PATH_NOT_FOUND。
        列舉得到、定址不到 —— 舊程式在這些地方一律用 Elements.Item()，
        原因就在這裡。
        """
        if self.app is None:
            raise BridgeError("NOT_CONNECTED", "需先 connect")
        node = self.app.Tree.FindNode(path)
        if node is None:
            raise BridgeError("PATH_NOT_FOUND", path)
        if child is None:
            return node
        try:
            found = node.Elements(child)
        except Exception:
            found = None
        if found is not None:
            return found
        # Elements(名稱) 對含空白的名稱同樣無效，只剩逐一列舉比對這條路。
        # 慢，但這種節點一層通常只有幾十個，而且沒有別的辦法。
        try:
            for element in node.Elements:
                if element.Name == child:
                    return element
        except Exception:
            pass
        raise BridgeError("PATH_NOT_FOUND", "{} 底下沒有 {!r}".format(
            path, child))

    def get(self, path: str, child: str | None = None) -> dict:
        try:
            node = self._node(path, child)
            data = {"exists": True}
            for attr in ("Value", "Dimension", "UnitString", "Basis"):
                try:
                    data[attr.lower()] = getattr(node, attr)
                except Exception:
                    data[attr.lower()] = None
            # HAP_COMPSTATUS(12)：這個節點自己的完成度／執行結果狀態，
            # 是個位元遮罩（0x20=結果有錯誤、0x40=輸入不完整…）。
            #
            # 這一項以前不在這裡，但 check_model 一直在讀回傳值裡的
            # compstatus —— 讀到的永遠是 None，於是每個區塊、每條物流
            # 都被判成「未知」。那從來不是「Aspen 讀不到」，是這支
            # get 從來沒問過。實測 AttributeValue(12) 在區塊節點上
            # 一向答得出來（收斂的塔回 RESULTS_SUCCESS、沒收斂的回
            # RESULTS_ERRORS），資料一直都在，只是沒有人去拿。
            try:
                data["compstatus"] = node.AttributeValue(12)
            except Exception:
                data["compstatus"] = None
            return _ok(data)
        except BridgeError as exc:
            return _err(exc.code, exc.detail)
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def set(self, path: str, value: Any, unit: int | None = None,
            basis: str | None = None, allow_default_unit: bool = False,
            child: str | None = None) -> dict:
        """寫入一個節點。

        用 Aspen 原生的寫入方法，不自己拼裝：
            有 basis → SetValueUnitAndBasis(value, unit or 0, basis)
            有 unit  → SetValueAndUnit(value, unit)
            都沒有   → node.Value = value

        先前這裡是「先設 UnitOfMeasure 再設 Value」，而且設單位失敗時
        默默跳過。那和原始程式的行為不同，而且正是本專案記錄過的
        「回報成功但沒真的寫進去」那一類錯誤。現在失敗就回報失敗。

        單位守門（唯一一條寫在橋接層的規則）：欄位有單位維度卻沒指定單位時
        拒絕寫入。Aspen 預設英制，省略單位會靜默寫成 lbmol/hr、psia ——
        本專案驗證期間這個陷阱出現過三次，每次都產生「看似成功、內部自洽、
        甚至與文獻對得上」的錯誤結果。雲端要繞過它必須明講
        allow_default_unit=True，繞過是明示的，不是預設的。
        """
        try:
            node = self._node(path, child)
        except BridgeError as exc:
            return _err(exc.code, exc.detail)

        # value=None 是「清掉這個欄位」，不是「寫一個數值但沒給單位」——
        # 兩者以前共用同一條路，導致 SetValueAndUnit(None, unit) 丟
        # COM 例外（'The supplied argument is an invalid type'，因為
        # SetValueAndUnit 的簽章不接受空值），而完全不給 unit 的版本
        # 又會先被下面的單位守門擋下來，回報 UNIT_REQUIRED——兩條路都
        # 走不通，等於這個橋接層根本沒有「清空」這個動作。
        #
        # 直接 node.Value = None 也不行，一樣是那個 invalid type 例外
        # ——pywin32 把 Python None 送過去變成 COM 認不得的類型。
        # 用 win32com 直接對節點做過實測（`node.Clear()` 這個看起來像
        # 專門用的方法在這個版本回報「還沒實作」）：`node.Value = ''`
        # （空字串）才是真的會清空、讀回來變成 None 的寫法——就跟使用者
        # 在 GUI 裡把欄位的文字整個刪掉一樣，Aspen 認得「空字串」是
        # 「這裡沒有值」，但認不得 Python 的 None／VT_NULL。
        if value is None:
            try:
                node.Value = ""
                return _ok({"path": path, "how": "Value", "cleared": True})
            except Exception as exc:
                return _err("COM_ERROR", repr(exc))

        dim = None
        try:
            dim = node.Dimension
        except Exception:
            pass
        if dim not in (None, 0) and unit is None and not allow_default_unit:
            return _err("UNIT_REQUIRED",
                        "{} 有單位維度 {}，必須指定 unit".format(path, dim))

        try:
            if basis is not None:
                node.SetValueUnitAndBasis(value, unit or 0, basis)
                how = "SetValueUnitAndBasis"
            elif unit is not None:
                node.SetValueAndUnit(value, unit)
                how = "SetValueAndUnit"
            else:
                node.Value = value
                how = "Value"
            return _ok({"path": path, "how": how})
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def children(self, path: str) -> dict:
        try:
            node = self._node(path)
            names = [el.Name for el in node.Elements]
            return _ok({"names": names, "count": len(names)})
        except BridgeError as exc:
            return _err(exc.code, exc.detail)
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def row(self, path: str, action: str, dim: int = 0, index: int = 0,
            label: str | None = None, secondary: bool = False) -> dict:
        """表格列的增刪與標籤設定。

        Aspen 有一部分資料不是樹，是表格：元件清單、亨利元件、純物性參數、
        反應式的化學計量、塔內件段落、平衡計算的項目 —— 都是「一列一筆」，
        用 `InsertRow` / `RemoveRow` / `SetLabel` 操作，而且列要用序號指涉。

        這些用 get/set/add_element 表達不出來，所以是獨立的一個動作。
        原始程式在 6 個模組、35 處呼叫這組 API。

        action:
            insert  在 dim 維度的 index 位置插入一列
            remove  移除該列
            label   設定該列的標籤（新增元件、新增反應物時用）

        dim 幾乎都是 0；反應式的副物流標籤用 1。
        """
        if action not in ("insert", "remove", "label", "labels"):
            return _err("BAD_ARG",
                        "action 只能是 insert、remove、label 或 labels")
        if action == "label" and label is None:
            return _err("BAD_ARG", "label 動作必須給 label")
        try:
            elements = self._node(path).Elements
            if action == "labels":
                # 讀某個維度上的所有標籤。二維表格（純物性參數、反應式的
                # 化學計量）的列是參數、欄是元件，兩邊都要讀得到才操作得了。
                #
                # 這也是判斷「這張表是幾維」的**非破壞性**做法：dim=1 讀得到
                # 東西就是二維。舊程式是試著寫一個標籤上去，失敗就當一維 ——
                # 拿寫入當探測，探測失敗時有沒有改到資料沒人知道。
                out = []
                index = 0
                while index < 2000:
                    try:
                        label = elements.Label(int(dim), index)
                    except Exception:
                        break
                    if label is None or not str(label).strip():
                        break
                    out.append(str(label).strip())
                    index += 1
                return _ok({"path": path, "dim": int(dim), "labels": out,
                            "count": len(out)})
            if action == "insert":
                elements.InsertRow(int(dim), int(index))
            elif action == "remove":
                elements.RemoveRow(int(dim), int(index))
            else:
                elements.SetLabel(int(dim), int(index), bool(secondary), label)
            return _ok({"path": path, "action": action,
                        "dim": int(dim), "index": int(index)})
        except BridgeError as exc:
            return _err(exc.code, exc.detail)
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    # 讀取一個節點時要抓的 COM 屬性。編號來自 Aspen 的 HAP_* 常數。
    # 這是「Aspen 節點有哪些欄位」的事實，不是判斷，所以放在橋接層。
    _ATTRS = {
        "enterable": 7,      # HAP_ENTERABLE  是不是可輸入欄位
        "prompt": 19,        # HAP_PROMPT     欄位說明文字
        "pq": 2,             # HAP_UNITROW    單位類別
        "um": 3,             # HAP_UNITCOL    單位索引
        "basis": 13,         # HAP_BASIS
        "outvar": 18,        # HAP_OUTVAR     是不是輸出變數
        "compstatus": 12,    # HAP_COMPSTATUS
        "recordtype": 6,
    }

    def _record(self, node, detail: str) -> dict | None:
        """把一個節點壓成純資料。detail="value" 只取值，"full" 取全部屬性。"""
        try:
            value = node.Value
        except Exception:
            value = None

        if detail == "value":
            if value is None:
                return None
            entry = {"value": value}
            try:
                if node.Dimension not in (None, 0):
                    entry["unit"] = node.UnitString
                    entry["dimension"] = node.Dimension
            except Exception:
                pass
            return entry

        entry: dict = {"value": value}
        for key, code in self._ATTRS.items():
            try:
                raw = node.AttributeValue(code)
            except Exception:
                continue
            if raw is None:
                continue
            # 屬性 5 是選項清單，是物件不是純量，另外處理
            entry[key] = raw
        try:
            if node.Dimension not in (None, 0):
                entry["dimension"] = node.Dimension
                entry["unitstring"] = node.UnitString
        except Exception:
            pass
        try:
            opts = node.AttributeValue(5)
            if opts is not None and hasattr(opts, "Elements"):
                entry["options"] = [opts.Elements(i).Value
                                    for i in range(opts.Elements.Count)]
        except Exception:
            pass
        return entry

    def subtree(self, path: str, max_nodes: int = 20000,
                detail: str = "value") -> dict:
        """一次取回整個子樹。

        detail="value" 只回傳有值的節點（最省）。
        detail="full"  回傳每個節點的完整屬性紀錄，讓雲端不必為了問
                       「這欄位可不可輸入／說明是什麼／單位類別為何」
                       再往返一次。原本的 _traverse_elements 在本地逐節點問了
                       這些屬性；搬到雲端後那會變成上千次往返，所以必須一次帶回。

        實測：單一區塊 Input 底下有 1,266 個節點，全模型 5,056 個。
        逐項往返在校內網路要 54.6 秒，一次 subtree 只要 4.1 秒。
        批次不是最佳化，是可行性前提。
        """
        if detail not in ("value", "full"):
            return _err("BAD_ARG", "detail 只能是 value 或 full")
        try:
            root = self._node(path)
        except BridgeError as exc:
            return _err(exc.code, exc.detail)

        out: dict = {}
        stack = [(path, root, None)]
        visited = 0
        collisions = 0
        while stack and visited < max_nodes:
            cur, node, idx = stack.pop()
            visited += 1
            entry = self._record(node, detail)
            if entry is not None:
                if idx is not None:
                    # 兄弟節點間的順序。Aspen 的單位表、成分表都是「用序號指涉」，
                    # 只有名字是不夠的。
                    entry["i"] = idx
                key = cur[len(path):].lstrip("\\")
                if key in out:
                    # 同層可以有同名節點。單位表的 TEMPERATURE 底下就有兩個 K
                    # （索引 1 和 3）。原本用名字當鍵，後者直接蓋掉前者，
                    # 一次靜默吃掉 78 個單位，而且完全不報錯 ——
                    # 直到逐欄比對才發現 case14 的溫度單位查不到。
                    key = "{}#{}".format(key, idx)
                    collisions += 1
                out[key] = entry
            try:
                for n, el in enumerate(node.Elements):
                    stack.append((cur + "\\" + el.Name, el, n))
            except Exception:
                pass
        return _ok({"root": path, "nodes": out, "count": len(out),
                    "visited": visited, "truncated": visited >= max_nodes,
                    "collisions": collisions, "detail": detail})

    def add_element(self, parent_path: str, name: str) -> dict:
        """在指定節點底下新增一個元素。名稱原樣傳給 Aspen。

        Aspen 用 `"名稱!型別"` 這種寫法帶型別（例如 `"COL2!RADFRAC"`），
        那是 Aspen 的慣例，屬於雲端的知識，橋接層不替它拼字串。

        先前這裡是 `Elements.Add(name, type_name)` 兩個引數 —— 和原始程式的
        單引數呼叫不是同一件事。這一點特別要緊：本專案記錄過一次
        `add_block` 回報 COM 失敗、卻留下一個刪不掉的殘骸物件，
        擋住所有後續執行，最後整個檔案重建。新增失敗不一定等於什麼都沒發生。
        """
        try:
            parent = self._node(parent_path)
            parent.Elements.Add(name)
            return _ok({"parent": parent_path, "name": name})
        except BridgeError as exc:
            return _err(exc.code, exc.detail)
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def remove_element(self, parent_path: str, name: str) -> dict:
        try:
            parent = self._node(parent_path)
            parent.Elements.Remove(name)
            return _ok()
        except BridgeError as exc:
            return _err(exc.code, exc.detail)
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    # ══ 執行 ════════════════════════════════════════════════════════
    def run(self, action: str = "start") -> dict:
        """啟動或中止模擬引擎。

        `Engine.Run2()` 名字看起來像非同步，**實測是同步的**：case09 這一次
        呼叫就卡住 227.8 秒才返回。先前這裡靠它「立刻返回、再輪詢 IsRunning」，
        那個假設是錯的 —— 輪詢迴圈永遠只跑一圈，雲端的 timeout 形同虛設，
        跑不完的模擬沒有任何辦法中止，學生的機器就一直卡著。

        改成丟到工作執行緒。COM 物件屬於建立它的 apartment，跨執行緒直接用會得到
        `CoInitialize 尚未被呼叫`，所以要用 COM 自己的方式把介面交過去：
        主執行緒 `CoMarshalInterThreadInterfaceInStream` 封裝，
        工作執行緒 `CoGetInterfaceAndReleaseStream` 取出。

        這樣主執行緒就空著，可以輪詢、可以逾時、可以呼叫 Stop()。
        """
        if self.app is None:
            return _err("NOT_CONNECTED")
        import pythoncom

        try:
            if action == "stop":
                self.app.Engine.Stop()
                # Stop() 只是「請求」停止 —— 要等工作執行緒裡的 Run2()
                # 真的返回，_run_done 才會被設起來。不等的話，呼叫端以為
                # 停好了，下一次 start 卻一直撞 ALREADY_RUNNING，而且看不出
                # 為什麼（實測經濟分析逾時後就是卡在這裡，最後只能殺行程）。
                stopped = True
                if self._run_done is not None:
                    stopped = self._run_done.wait(STOP_WAIT_S)
                return _ok({"status": "stopped" if stopped else "still_running",
                            "stopped": stopped,
                            "waited_s": 0 if stopped else STOP_WAIT_S,
                            "detail": None if stopped else
                            "已經送出停止要求，但引擎在 {:.0f} 秒內沒有停下來。"
                            "這一份 Aspen 不能再啟動新的執行了 —— 先用 "
                            "close_project 關掉再重開，或直接換一份。"
                            .format(STOP_WAIT_S)})
            if action != "start":
                return _err("BAD_ARG", "action 只能是 start 或 stop")
            if self._run_done is not None and not self._run_done.is_set():
                running_for = (time.monotonic() - (self._run_started or 0)
                               if self._run_started else None)
                return _err("ALREADY_RUNNING",
                            "這一份 Aspen 還有一次執行沒有結束{}。"
                            "同時跑兩次會讓兩邊都拿不到正確結果。"
                            "先用 run(action='stop') 停掉；若停不下來，"
                            "那次執行已經卡死在引擎裡，close_project 再重開"
                            "是唯一的出路。".format(
                                "（已經跑了 {:.0f} 秒）".format(running_for)
                                if running_for else ""))

            stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
                pythoncom.IID_IDispatch, self.app)
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

        self._run_error = None
        self._run_done = threading.Event()
        self._run_started = time.monotonic()

        # 監聽要掛在主執行緒 —— 連接點屬於建立 COM 物件的 apartment，
        # 事件也是在主執行緒 PumpWaitingMessages() 時送達。工作執行緒
        # 只負責喊 Run2()，不碰事件。
        self._cp_messages = []
        self._detach_control_panel()
        try:
            self._cp_handler = win32.WithEvents(self.app, _ControlPanelEvents)
            self._cp_handler.sink = self._cp_messages
            self._cp_capture = "ok"
        except Exception as exc:
            # WithEvents 需要 early-bound 型別庫（makepy）。快取沒建起來時
            # 就是掛在這裡。先自己把型別庫生出來再試一次 —— 這是能當場修好
            # 的事，沒道理讓學生整場拿不到 Control Panel 訊息。
            build_err = _build_typelib(self.app)
            self._cp_handler = None
            self._cp_capture = "unavailable: " + repr(exc)[:200]
            if build_err is None:
                try:
                    self._cp_handler = win32.WithEvents(self.app,
                                                        _ControlPanelEvents)
                    self._cp_handler.sink = self._cp_messages
                    self._cp_capture = "ok"
                except Exception as again:
                    self._cp_handler = None
                    self._cp_capture = ("unavailable: 型別庫(makepy)已重建，"
                                        "但監聽仍掛不上：" + repr(again)[:160])
            else:
                # 這一步失敗不該拖垮執行本身 —— 引擎照跑，只是拿不到
                # Control Panel 原話，要把「拿不到」講出來，不能默默
                # 當成沒有訊息。
                self._cp_capture = ("unavailable: 型別庫(makepy)建不起來："
                                    + build_err)

        def _worker():
            pythoncom.CoInitialize()
            app = None
            try:
                app = win32.Dispatch(
                    pythoncom.CoGetInterfaceAndReleaseStream(
                        stream, pythoncom.IID_IDispatch))
                app.Engine.Run2()
            except Exception as exc:
                self._run_error = repr(exc)
            finally:
                # 一定要先放掉代理物件，再 CoUninitialize。
                # 反過來的話，代理會在 apartment 已經關掉之後才被回收，
                # 結果是主執行緒之後呼叫 Close() 直接死鎖 —— 不報錯、不逾時，
                # 就是永遠不回來。這個順序錯誤讓測試在跑完之後掛了 11 分鐘。
                app = None
                self._run_done.set()
                pythoncom.CoUninitialize()

        threading.Thread(target=_worker, daemon=True).start()
        return _ok({"status": "started"})

    def run_status(self, poll_s: float = 0.0) -> dict:
        """回報引擎狀態。poll_s > 0 時最多在本地等這麼久再回報。

        等待放在本地是為了少往返。一次模擬動輒數分鐘，若雲端每 5 秒問一次，
        那是幾百次無謂的網路來回。等多久仍由雲端決定，橋接層只是照做。

        等待期間要呼叫 `PumpWaitingMessages()`：Aspen 是跨行程的 COM 伺服器，
        不抽訊息佇列的話，主執行緒收不到回呼，也叫不動 Stop()。
        """
        if self.app is None:
            return _err("NOT_CONNECTED")
        if self._run_done is None:
            return _ok({"status": "idle", "completed": True})

        import pythoncom
        deadline = time.monotonic() + max(0.0, float(poll_s))
        while True:
            pythoncom.PumpWaitingMessages()
            if self._run_done.is_set():
                elapsed = round(time.monotonic() - self._run_started, 1)
                # 引擎結束後尾端還有幾行在佇列裡，再抽幾次才收得齊。
                for _ in range(10):
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.02)
                self._detach_control_panel()
                extra = {"control_panel": list(self._cp_messages),
                         "control_panel_capture": self._cp_capture}
                if self._run_error:
                    return _ok({"status": "failed", "completed": True,
                                "detail": self._run_error, "elapsed_s": elapsed,
                                **extra})
                return _ok({"status": "completed", "completed": True,
                            "elapsed_s": elapsed, **extra})
            if time.monotonic() >= deadline:
                return _ok({"status": "running", "completed": False,
                            "elapsed_s": round(
                                time.monotonic() - self._run_started, 1)})
            self._run_done.wait(0.25)

    def _detach_control_panel(self) -> None:
        """放掉事件物件。留著會持有 COM 參考，關檔時可能卡住。"""
        handler, self._cp_handler = self._cp_handler, None
        if handler is None:
            return
        try:
            handler.sink = None
            close = getattr(handler, "close", None)
            if callable(close):
                close()
        except Exception:
            pass

    def reinit(self) -> dict:
        if self.app is None:
            return _err("NOT_CONNECTED")
        try:
            self.app.Reinit()
            return _ok()
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def get_log(self, history: bool = True, tail: int = 6000) -> dict:
        """回報執行狀態，以及 Aspen 自己寫的 history 檔。

        `run_status_exists` 是關鍵的一項：輸入轉譯失敗時，Aspen **根本不會**
        建立 Run-Status 節點。先前只讀節點值，讀到空字串就當作「沒有訊息」，
        於是一個從未執行的模型會被回報成「執行完成」。

        history 檔裡有 Aspen 自己的錯誤敘述，例如
        「NO PRESSURE OR WORK SPECIFICATIONS ARE GIVEN」——
        那是最準確的線索，卻一直被丟掉。
        """
        if self.app is None:
            return _err("NOT_CONNECTED")
        out: dict = {"run_status_exists": False, "text": "",
                     "history": "", "errors": [],
                     "control_panel": list(self._cp_messages),
                     "control_panel_capture": self._cp_capture}
        try:
            node = self.app.Tree.FindNode(
                r"\Data\Results Summary\Run-Status\Output\PER_ERROR")
            out["run_status_exists"] = node is not None
            if node is not None:
                out["text"] = str(node.Value)[:20000]
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

        if history:
            found = self._history_text(tail)
            out["history"] = found
            out["errors"] = _extract_errors(found)
        return _ok(out)

    def _history_text(self, tail: int) -> str:
        """讀最新的 .his 檔。

        Aspen 把它寫在哪不固定：開舊檔時放在模型檔旁邊，新建的專案還沒有
        檔案路徑，就落在它自己的工作目錄或系統暫存區。所以幾個候選位置
        都找一遍，取**最近寫入**的那一個。

        只認 90 秒內寫的檔案 —— 更舊的是上一次執行留下的，
        拿它當這次的錯誤會指向完全不相干的問題。
        """
        candidates = []
        if self._path:
            candidates.append(os.path.dirname(self._path))
        candidates.append(os.getcwd())
        temp = os.environ.get("TEMP") or os.environ.get("TMP")
        if temp:
            candidates.append(temp)

        newest, newest_at = None, 0.0
        cutoff = time.time() - 90
        seen = set()
        for folder in candidates:
            key = os.path.normcase(folder or "")
            if not key or key in seen:
                continue
            seen.add(key)
            try:
                found = glob.glob(os.path.join(folder, "*.his"))
            except OSError:
                continue
            for path in found:
                try:
                    stamp = os.path.getmtime(path)
                except OSError:
                    continue
                if stamp >= cutoff and stamp > newest_at:
                    newest, newest_at = path, stamp

        if newest is None:
            return ""
        try:
            with io.open(newest, encoding="utf-8", errors="replace") as handle:
                return handle.read()[-tail:]
        except OSError:
            return ""

    # ══ GUI 自動化 ══════════════════════════════════════════════════
    # elec_wizard、add_utility(preset)、set_costing_template、
    # run_economic_analysis、export_economic_excel 這五項無法走 COM。
    # 實測：搜尋 356 ms、動作 2.3 ms —— 搜尋是動作的 155 倍。
    # 雲端必須積極快取 handle，找一次重複用。

    @staticmethod
    def _auto():
        import uiautomation as auto
        return auto

    def _cache(self, ctrl) -> str:
        cid = "c{}".format(len(self._ui_cache))
        self._ui_cache[cid] = ctrl
        return cid

    @staticmethod
    def _rect(ctrl) -> dict | None:
        """控制項在螢幕上的位置。

        看起來多餘，其實不可少：Aspen 的導覽列上有**兩個都叫 Properties** 的
        控制項，只能靠上下位置分辨（在中線以上的是樹節點，以下的是按鈕）。
        沒有座標，雲端就沒辦法做這個判斷。
        """
        try:
            r = ctrl.BoundingRectangle
            return {"left": r.left, "top": r.top,
                    "right": r.right, "bottom": r.bottom}
        except Exception:
            return None

    def _describe(self, ctrl, level: int = 0) -> dict:
        out = {"id": self._cache(ctrl), "level": level}
        for key, get in (("type", lambda: ctrl.ControlTypeName),
                         ("name", lambda: (ctrl.Name or "")[:120]),
                         ("auto_id", lambda: (ctrl.AutomationId or "")[:60]),
                         ("enabled", lambda: bool(ctrl.IsEnabled))):
            try:
                out[key] = get()
            except Exception:
                pass
        rect = self._rect(ctrl)
        if rect:
            out["rect"] = rect
        return out

    def _root(self, root_id: str | None, root_name: str | None,
              timeout: float = 2.0):
        """決定從哪裡開始找。已快取的控制項優先，其次具名視窗，最後整個桌面。"""
        if root_id:
            ctrl = self._ui_cache.get(root_id)
            if ctrl is None:
                raise BridgeError("UI_NOT_FOUND", root_id)
            return ctrl
        auto = self._auto()
        if root_name:
            win = auto.WindowControl(searchDepth=3, RegexName=root_name)
            if not win.Exists(timeout, 0.2):
                raise BridgeError("UI_NOT_FOUND", root_name)
            return win
        return auto.GetRootControl()

    def ui_tree(self, root_name: str | None = None, root_id: str | None = None,
                depth: int = 3, limit: int = 3000) -> dict:
        """列出某個控制項底下的所有子控制項。

        root_id 讓雲端可以「從剛剛找到的那個視窗往下看」，而不是每次都從桌面
        重新掃。精靈開起來之後，所有搜尋都應該限定在精靈視窗內 ——
        全域搜尋不只慢，還會找到主視窗上同名的東西。
        """
        ok, why = _interactive_desktop()
        if not ok:
            return _err("NO_INTERACTIVE_DESKTOP", why)
        try:
            self._auto()
        except ImportError as exc:
            return _err("NO_UIAUTOMATION", repr(exc))
        try:
            node = self._root(root_id, root_name)
        except BridgeError as exc:
            return _err(exc.code, exc.detail)
        except Exception as exc:
            return _err("UI_ERROR", repr(exc))

        controls: list = []

        def walk(ctrl, level):
            if level > depth or len(controls) >= limit:
                return
            try:
                children = ctrl.GetChildren()
            except Exception:
                return
            for child in children:
                try:
                    controls.append(self._describe(child, level))
                except Exception:
                    continue
                walk(child, level + 1)

        walk(node, 1)
        return _ok({"controls": controls, "count": len(controls),
                    "truncated": len(controls) >= limit})

    def ui_find(self, name: str | None = None, control_type: str | None = None,
                auto_id: str | None = None, sub_name: str | None = None,
                class_name: str | None = None, root_id: str | None = None,
                root_name: str | None = None, depth: int = 8,
                timeout: float = 3.0) -> dict:
        """找一個控制項。

        三種條件可以混用：名稱（正規表示式）、控制項型別、AutomationId。
        AutomationId 是最可靠的 —— 它不會因為介面語言或版本而變，
        Aspen 精靈的下拉選單就只能靠它認（cboDBSelect）。

        root_id／root_name 限定搜尋範圍。精靈開起來後一定要限定在精靈視窗內，
        否則會找到主視窗上同名的控制項。
        """
        ok, why = _interactive_desktop()
        if not ok:
            return _err("NO_INTERACTIVE_DESKTOP", why)
        try:
            auto = self._auto()
        except ImportError as exc:
            return _err("NO_UIAUTOMATION", repr(exc))
        try:
            scope = self._root(root_id, root_name) if (root_id or root_name)                 else None
        except BridgeError as exc:
            return _err(exc.code, exc.detail)

        kw: dict = {"searchDepth": depth}
        if name:
            kw["RegexName"] = name
        if auto_id:
            kw["AutomationId"] = auto_id
        if class_name:
            # Windows 的標準對話框類別是 #32770。檔案瀏覽視窗沒有固定標題
            # （會隨語言變），只能靠類別認。
            kw["ClassName"] = class_name
        if sub_name:
            # 子字串比對。用正規表示式也可以，但選項文字裡常有 + 這種
            # 會被當成語法的字元（"Hydronium ion H3O+"），SubName 沒這問題。
            kw["SubName"] = sub_name
        cls_name = (control_type or "") + "Control"
        try:
            if scope is not None:
                factory = getattr(scope, cls_name, None) or scope.Control
                ctrl = factory(**kw)
            else:
                cls = getattr(auto, cls_name, auto.Control)
                ctrl = cls(**kw)
            if not ctrl.Exists(timeout, 0.2):
                return _ok({"found": False})
            data = self._describe(ctrl)
            data["found"] = True
            return _ok(data)
        except Exception as exc:
            return _err("UI_ERROR", repr(exc))

    def ui_wait(self, name: str | None = None, auto_id: str | None = None,
                control_type: str | None = None, sub_name: str | None = None,
                root_id: str | None = None, timeout: float = 10.0) -> dict:
        """等一個控制項出現。和 ui_find 同樣的條件，只是等比較久。

        精靈換頁後要等下一頁的標題出現才能繼續。等待放在本地 ——
        若改成雲端每秒問一次，一次換頁就是好幾趟往返。
        """
        return self.ui_find(name=name, auto_id=auto_id, sub_name=sub_name,
                            control_type=control_type, root_id=root_id,
                            timeout=timeout)

    # 控制項上可以做的原始動作。每一個都只是一次 UI Automation 呼叫，
    # 這裡不做任何「先試 A 不行再試 B」的判斷 —— 那個順序是知識，在雲端。
    # 雲端要退回路徑時，用 batch 的 first_success 模式一次把整串送下來。
    def ui_act(self, control_id: str, action: str,
               text: str | None = None) -> dict:
        ok, why = _interactive_desktop()
        if not ok:
            return _err("NO_INTERACTIVE_DESKTOP", why)
        ctrl = self._ui_cache.get(control_id)
        if ctrl is None:
            return _err("UI_NOT_FOUND", control_id)
        auto = self._auto()
        try:
            if action == "read":
                return _ok(self._ui_read(ctrl))
            if action == "exists":
                return _ok({"exists": bool(ctrl.Exists(
                    float(text) if text else 1.0, 0.2))})
            if action == "parent":
                parent = ctrl.GetParentControl()
                if parent is None:
                    return _err("UI_NOT_FOUND", "沒有父控制項")
                return _ok(self._describe(parent))
            if action == "invoke":
                ctrl.GetInvokePattern().Invoke()
            elif action == "click":
                ctrl.Click()
            elif action == "setfocus":
                ctrl.SetFocus()
            elif action == "sendkeys":
                # 送到這個控制項，不是送到目前焦點所在。核取方塊要用
                # "{Space}" 切換時，兩者結果不同 —— 全域送鍵會打到別的地方。
                ctrl.SendKeys(text or "", waitTime=0.8)
            elif action == "sendkeys_global":
                auto.SendKeys(text or "")
            elif action == "scroll_into_view":
                ctrl.GetScrollItemPattern().ScrollIntoView()
            elif action == "double_click":
                ctrl.DoubleClick()
            elif action == "coord_click":
                # 移動滑鼠到控制項中心點一下，再把游標移回原位。
                # 有些控制項三種 pattern 都不吃，只認真的滑鼠事件。
                # 還游標是禮貌：學生的滑鼠不該無故被搬走。
                self._coord_click(ctrl)
            elif action == "post_click":
                # 送 BM_CLICK 訊息給按鈕，完全不碰滑鼠。
                # Win32 對話框（重新初始化、確認視窗）的按鈕常常
                # 對 InvokePattern 沒反應，但吃這個。
                import win32gui
                handle = ctrl.NativeWindowHandle
                if not handle:
                    return _err("UI_ERROR", "這個控制項沒有視窗代號")
                win32gui.PostMessage(handle, 0x00F5, 0, 0)   # BM_CLICK
            elif action == "foreground":
                import win32gui
                handle = ctrl.NativeWindowHandle
                if not handle:
                    return _err("UI_ERROR", "這個控制項沒有視窗代號")
                win32gui.SetForegroundWindow(handle)
            elif action == "close_window":
                # 送 WM_CLOSE 給視窗本身，不找按鈕。用在「這個視窗有沒有
                # OK 都無所謂，反正要把它關掉」的場合 —— 例如匯出流程
                # 留下的確認視窗，找不到 OK 按鈕時不能就這樣放著，
                # 那會一直擋住同一條工作執行緒後面的所有動作。
                import win32con
                import win32gui
                handle = ctrl.NativeWindowHandle
                if not handle:
                    return _err("UI_ERROR", "這個控制項沒有視窗代號")
                win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
            elif action == "settext":
                ctrl.GetValuePattern().SetValue(text or "")
            elif action == "select":
                ctrl.GetSelectionItemPattern().Select()
            elif action == "expand":
                ctrl.GetExpandCollapsePattern().Expand()
            elif action == "collapse":
                ctrl.GetExpandCollapsePattern().Collapse()
            elif action == "toggle":
                ctrl.GetTogglePattern().Toggle()
            elif action == "legacy_default":
                # 舊式 MSAA 介面。Aspen 的精靈有不少控制項沒有實作
                # 現代的 UI Automation pattern，只認這個。
                ctrl.GetLegacyIAccessiblePattern().DoDefaultAction()
            elif action == "legacy_select":
                ctrl.GetLegacyIAccessiblePattern().Select(
                    auto.AccessibleSelection.TakeFocus
                    | auto.AccessibleSelection.TakeSelection)
            else:
                return _err("UNKNOWN_ACTION", action)
            return _ok({"control_id": control_id, "action": action})
        except Exception as exc:
            return _err("UI_ERROR", repr(exc))

    @staticmethod
    def _coord_click(ctrl) -> None:
        import win32api
        import win32con
        rect = ctrl.BoundingRectangle
        cx = int(rect.left + (rect.right - rect.left) / 2)
        cy = int(rect.top + (rect.bottom - rect.top) / 2)
        old = win32api.GetCursorPos()
        win32api.SetCursorPos((cx, cy))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, cx, cy, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, cx, cy, 0, 0)
        time.sleep(0.05)
        win32api.SetCursorPos(old)

    def _ui_read(self, ctrl) -> dict:
        """讀一個控制項目前的狀態。

        沒有這個，雲端就沒辦法判斷「這個核取方塊勾了沒」「下拉選單現在選什麼」，
        也就寫不出精靈那種一步看一步的流程。每個屬性各自 try —— 不同控制項
        支援的 pattern 不一樣，讀不到就是沒有，不是錯誤。
        """
        out: dict = {}
        for key, get in (
            ("name", lambda: ctrl.Name),
            ("type", lambda: ctrl.ControlTypeName),
            ("auto_id", lambda: ctrl.AutomationId),
            ("enabled", lambda: bool(ctrl.IsEnabled)),
            ("offscreen", lambda: bool(ctrl.IsOffscreen)),
            ("rect", lambda: Bridge._rect(ctrl)),
            ("value", lambda: ctrl.GetValuePattern().Value),
            ("selected", lambda: bool(
                ctrl.GetSelectionItemPattern().IsSelected)),
            ("toggle", lambda: int(ctrl.GetTogglePattern().ToggleState)),
            ("expand", lambda: int(
                ctrl.GetExpandCollapsePattern().ExpandCollapseState)),
            ("legacy_value", lambda: ctrl.GetLegacyIAccessiblePattern().Value),
            ("legacy_state", lambda: int(
                ctrl.GetLegacyIAccessiblePattern().State)),
            ("hwnd", lambda: int(ctrl.NativeWindowHandle) or None),
            ("class_name", lambda: ctrl.ClassName),
            # 下拉選單目前選的項目。Aspen 精靈的幾個下拉選單既沒有
            # ValuePattern 也沒有 legacy value，只有這個讀得到 ——
            # 少了它，雲端就無法確認「到底選成功了沒」。
            ("selection", lambda: [i.Name for i in
                                   ctrl.GetSelectionPattern().GetSelection()]),
        ):
            try:
                out[key] = get()
            except Exception:
                pass
        return out

    def read_table(self, path: str, sheet: str | None = None,
                   max_rows: int = 400, max_cols: int = 40,
                   open_timeout_s: float = 20.0) -> dict:
        """把試算表讀成純資料回傳。

        經濟分析的結果是匯出成 .xlsx 的，檔案在學生電腦上，雲端看不到。
        所以要有一個動作把內容取回去。

        這裡只做「把儲存格變成資料」這件事 —— 哪一欄是總資本支出、
        哪個數字要拿來比較，那是雲端的判斷，不在這裡。

        **開檔這一步實測過真的會卡住。**「Send to Excel/ASW」會把剛匯出
        的檔案在 Excel 裡開起來（沒有勾選項可以關掉這個行為），export
        回報完成的時候只確認了 GUI 對話框關了，不保證 OS 層那個檔案已經
        真的寫完、鎖也放掉了。剛匯出完立刻讀，實測撞過一次：
        `openpyxl.load_workbook` 卡住不動，過了 30 分鐘（外部逾時）才
        中止；手動重試一次，同一個檔案幾乎立刻就讀成功。橋接層是單執行
        緒，這裡卡住等於卡住後面每一個呼叫，不能沒有上限。

        開檔放進背景執行緒、限時等待；逾時或讀到壞掉的檔案就短暫等待
        後重試（檔案通常在幾秒內就穩定），重試用完才把逾時／錯誤回報
        出去 —— 讓卡住的執行緒繼續留在背景（daemon，不會擋 process
        結束），橋接層本身不再被它卡住。
        """
        try:
            import openpyxl
        except ImportError as exc:
            return _err("NO_OPENPYXL", repr(exc))
        if not os.path.isfile(path):
            return _err("FILE_NOT_FOUND", path)

        book = None
        last_err: Exception | None = None
        for attempt in range(3):
            slot: dict = {}

            def _load():
                try:
                    slot["book"] = openpyxl.load_workbook(
                        path, data_only=True, read_only=True)
                except Exception as exc:
                    slot["error"] = exc

            worker = threading.Thread(target=_load, daemon=True)
            worker.start()
            worker.join(open_timeout_s)
            if worker.is_alive():
                last_err = TimeoutError(
                    "開啟超過 {:.0f} 秒沒回應".format(open_timeout_s))
            elif "error" in slot:
                last_err = slot["error"]
            else:
                book = slot.get("book")
                break
            if attempt < 2:
                time.sleep(2.0)
        if book is None:
            return _err("READ_ERROR" if not isinstance(last_err, TimeoutError)
                        else "READ_TIMEOUT",
                        "{}（檔案可能還沒真的寫完或被鎖住，稍等幾秒再試一次：{!r}）"
                        .format(repr(last_err), path))
        try:
            names = list(book.sheetnames)
            targets = [sheet] if sheet else names
            out: dict = {}
            for name in targets:
                if name not in names:
                    continue
                rows = []
                for row in book[name].iter_rows(max_row=max_rows,
                                                max_col=max_cols,
                                                values_only=True):
                    if any(c is not None for c in row):
                        rows.append([c for c in row])
                out[name] = rows
            return _ok({"path": path, "sheets": names, "data": out})
        finally:
            try:
                book.close()
            except Exception:
                pass

    def sleep(self, seconds: float) -> dict:
        """在本地等一段時間。

        看起來像是不該存在的動作，但沒有它更糟：Aspen 的精靈在點擊之間需要
        真實的延遲（原程式有 24 處 time.sleep），若改成雲端等待，每一次延遲
        都要多一趟網路往返，81 次互動會變成兩百多趟。

        等多久由雲端決定，橋接層只是照做。上限 60 秒，避免編排寫錯就卡死。
        """
        secs = max(0.0, min(float(seconds), 60.0))
        time.sleep(secs)
        return _ok({"slept": secs})

    # ══ 派送 ════════════════════════════════════════════════════════
    OPS = (
        "connect", "create_file", "open_file", "save", "close", "app_prop",
        "get", "set", "children", "subtree", "add_element", "remove_element",
        "row",
        "run", "run_status", "reinit", "get_log",
        "ui_tree", "ui_find", "ui_wait", "ui_act", "sleep", "read_table",
    )

    def execute(self, op: str, args: dict | None = None) -> dict:
        if op == "batch":
            return self.batch(**(args or {}))
        if op not in self.OPS:
            return _err("UNKNOWN_OP", op)
        self.op_count += 1
        try:
            return getattr(self, op)(**(args or {}))
        except TypeError as exc:
            return _err("BAD_ARGS", repr(exc))
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def batch(self, ops: list, stop_on_error: bool = True,
              mode: str | None = None) -> dict:
        """一次送多個動作。mode 決定遇到成敗時要不要繼續。

            all            全部做完，不管成敗
            stop_on_error  失敗就停（預設）
            first_success  **成功就停** —— 給退回路徑用

        first_success 是為了 GUI 而加的。Aspen 精靈裡有些控制項只認舊式的
        MSAA 介面，有些只認現代 pattern，程式要「先試 Invoke，不行試
        DoDefaultAction，再不行用點的」。若每一次嘗試都要一趟往返，
        81 次互動就會變成兩百多趟。

        順序仍然由雲端決定，橋接層只是照著做到第一個成功為止 ——
        那是控制流程，不是判斷該用哪個方法。
        """
        if mode is None:
            mode = "stop_on_error" if stop_on_error else "all"
        if mode not in ("all", "stop_on_error", "first_success"):
            return _err("BAD_ARG", "mode 只能是 all、stop_on_error 或 first_success")

        results = []
        for item in ops:
            res = self.execute(item.get("op"), item.get("args"))
            results.append(res)
            if mode == "stop_on_error" and not res["ok"]:
                break
            if mode == "first_success" and res["ok"]:
                break
        succeeded = next((i for i, r in enumerate(results) if r["ok"]), None)
        return _ok({"results": results, "executed": len(results),
                    "mode": mode, "first_ok": succeeded})
