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


def _err(code: str, detail: str = "") -> dict:
    return {"ok": False, "data": None,
            "error": {"code": code, "detail": str(detail)[:500]}}


class Bridge:
    """20 個通用動作。不含業務邏輯。"""

    def __init__(self) -> None:
        self.app = None
        self._ui_cache: dict = {}      # 控制項 handle 表（快取，非邏輯）
        self._run_started: float | None = None
        self._run_done: threading.Event | None = None
        self._run_error: str | None = None
        self._path: str | None = None      # 目前開著的檔，用來找 history 檔
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
            return _ok({"progid": progid})
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

    def create_file(self) -> dict:
        if self.app is None:
            return _err("NOT_CONNECTED")
        try:
            self.app.InitNew2()
            return _ok()
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

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
        try:
            getattr(self.app, loader)(path)
            self._path = path
            return _ok({"path": path, "loader": loader})
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

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
        if self.app is not None:
            try:
                self.app.Close()
            except Exception:
                pass
            self.app = None
        self._ui_cache.clear()
        return _ok()

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
                return _ok({"status": "stopping"})
            if action != "start":
                return _err("BAD_ARG", "action 只能是 start 或 stop")
            if self._run_done is not None and not self._run_done.is_set():
                return _err("ALREADY_RUNNING")

            stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
                pythoncom.IID_IDispatch, self.app)
        except Exception as exc:
            return _err("COM_ERROR", repr(exc))

        self._run_error = None
        self._run_done = threading.Event()
        self._run_started = time.monotonic()

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
                if self._run_error:
                    return _ok({"status": "failed", "completed": True,
                                "detail": self._run_error, "elapsed_s": elapsed})
                return _ok({"status": "completed", "completed": True,
                            "elapsed_s": elapsed})
            if time.monotonic() >= deadline:
                return _ok({"status": "running", "completed": False,
                            "elapsed_s": round(
                                time.monotonic() - self._run_started, 1)})
            self._run_done.wait(0.25)

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
                     "history": "", "errors": []}
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
