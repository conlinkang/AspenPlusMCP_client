# Aspen Plus MCP 學生端

讓 Claude Desktop 能操作你電腦上的 Aspen Plus，透過課程的雲端伺服器取得工具邏輯。

這個資料夾就是完整的本地端：

| 檔案 | 做什麼 |
|---|---|
| `agent.py` | 跟 Claude Desktop 對話的 MCP 伺服器，把工具呼叫轉發給雲端 |
| `bridge.py` | 讀寫你電腦上 Aspen Plus 的資料節點 |
| `setup_student.py` | 安裝程式：檢查環境、測連線、寫入 Claude Desktop 設定 |
| `requirements.txt` | 上面三個檔案要用的 Python 套件清單 |

不含任何工具邏輯（怎麼判斷收斂、怎麼設定反應、怎麼跑經濟分析……）——
那些都在雲端伺服器上執行，這裡只負責「轉發」跟「操作 Aspen」。

> **目前只服務國立中正大學（CCU）化工系課程**（伺服器架在中正校內）。
> 請用**中正校內網路**（教室、系上機房、校內 WiFi）；校外／宿舍網路
> 不保證連得進去，連不上請改用校內網路或學校 VPN。也**還沒有開放給
> 其他學校的師生申請**
> —— 如果你不是中正的學生，這份說明目前對你沒有用，請聯絡你自己課程
> 的授課教師確認有沒有對應的雲端服務。

---

## 能做什麼

裝好之後，用平常說話的方式請 Claude 幫你操作 Aspen Plus。雲端目前提供約 70 支工具，常用的分這幾類：

| 類別 | 做什麼 |
|---|---|
| 開檔、存檔、匯出 | 開新模型或舊模型、存檔，匯出 `.bkp`、`.inp`、報告、摘要等 |
| 建流程圖 | 加元件（可用名稱、化學式、CAS 號碼搜尋）、選物性方法、加單元與物流、接線、設規格 |
| 反應與蒸餾塔 | 反應集合與反應式、RadFrac 的反應段、板段壓力與效率、內件、設計規格 |
| 跑與看結果 | 執行模擬、讀收斂與警告、讀物流與單元結果、重設 |
| 分析 | 靈敏度分析、流程圖設計規格、Calculator、平衡集合 |
| 經濟 | 公用工程、物流價格、情境、成本模板、經濟評估與匯出 Excel |
| 說明與回報 | `skills`（每種單元、每個步驟的操作指南）、`report_problem`（回報問題）、`my_usage`（看用量） |

不用背工具名稱：直接說你要做什麼，Claude 會先查雲端的操作指南再動手。
操作的是**你電腦上**開著的那個 Aspen，不會動到別人的。

---

## 前置需求

- 這台電腦已經安裝 **Aspen Plus** 並且授權可以正常開啟
- 一個裝了 `requirements.txt` 裡那些套件的 Python（`pywin32`、`mcp`、
  `uiautomation`、`openpyxl`）——課程機房的 Aspen 環境通常已經有，
  例如 `D:\Aspen_MCP\.venv\Scripts\python.exe`；如果沒有，自己找個
  Python 裝：`python -m pip install -r requirements.txt`
  （`requirements.txt` 要跟其他三個檔案放同一個資料夾）
- 已安裝 **Claude Desktop**

---

## 步驟一：申請帳號

> 請在**中正大學校內網路**（教室、系上機房、校內 WiFi）操作這一步。
> **目前只接受中正的申請**，其他學校的 email 雖然格式上填得進去，但
> 送出後也連不到這台伺服器 —— 這是暫時性的限制，不是刻意排除，未來
> 若開放對外會另行公告。

1. 在校內網路瀏覽器開啟：**http://140.123.126.87:8787/signup**
2. 填寫：
   - **學術 email**（用來收驗證信；目前限中正大學信箱）
   - **姓名**
   - **身分**：教師 / 學生
3. 送出後會收到一封驗證信，**24 小時內**點裡面的連結完成驗證
4. 驗證完成後，申請會送交課程管理者審核
5. 審核通過後，系統會寄一封「**申請已核准**」的信到同一個信箱，裡面有
   領取 token 的連結（**24 小時內**有效）。點開連結、按「**顯示我的
   token**」，就會看到你的帳號代號和 **token**（**只顯示這一次**，請立刻
   複製保存 —— 沒存到只能請管理者重新產生一組）
6. 沒收到那封信（先看垃圾信匣）或連結過期了：到
   **http://140.123.126.87:8787/status** 輸入 email，系統會再寄一封
   領取信。5 分鐘內只能寄一次，新的連結會讓舊的失效。查詢頁本身**不會**
   顯示 token —— 一定要收得到你的信箱才領得到，別人知道你的 email 也
   領不走

每個帳號預設每月可呼叫工具 **1000 次**，每月自動重置。開 Claude 時載入
工具清單不算次數。額度用完時，呼叫工具會看到「這個帳號本期的工具呼叫
額度已經用完（下次重置：……）」—— 這時重試沒有用，等重置時間過後再用，
或聯絡課程助教調整額度。

想知道自己用了多少、還剩幾次、什麼時候重置，在 Claude 裡請它呼叫
`my_usage` 就好 —— 這個查詢本身不算次數，額度用完了也查得到。

---

## 步驟二：下載這四個檔案

兩種方式擇一：**手動下載**，或是**請你的 AI agent 幫你做**。

### 方式 A：手動下載

把 `agent.py`、`bridge.py`、`setup_student.py`、`requirements.txt` 四個
檔案放進同一個資料夾（例如 `D:\AspenPlusMCP_client\`）。四個檔案要放在
一起，位置隨意，但不能拆開。

### 方式 B：貼網址給 AI agent 安裝

如果你在用 Claude Code、Claude Desktop 或其他有讀寫檔案與執行指令能力
的 AI agent，可以把下面整段（含你已經拿到的 token）貼給它，讓它自動
完成下載、環境檢查、寫入設定：

```
請幫我安裝 Aspen Plus MCP 學生端：

1. 從 https://github.com/conlinkang/AspenPlusMCP_client 這個倉庫下載
   agent.py、bridge.py、setup_student.py、requirements.txt 四個檔案
   （用 git clone，或直接抓這四個 raw 網址都可以：
   https://raw.githubusercontent.com/conlinkang/AspenPlusMCP_client/master/agent.py
   https://raw.githubusercontent.com/conlinkang/AspenPlusMCP_client/master/bridge.py
   https://raw.githubusercontent.com/conlinkang/AspenPlusMCP_client/master/setup_student.py
   https://raw.githubusercontent.com/conlinkang/AspenPlusMCP_client/master/requirements.txt
   ），四個檔案放同一個資料夾。

2. 執行：
   python -m pip install -r requirements.txt
   python setup_student.py --url http://140.123.126.87:8787 --token 你的token

3. 如果它回報找不到裝齊 requirements.txt 那些套件的 Python，先問我要用
   哪一個直譯器（例如 D:\Aspen_MCP\.venv\Scripts\python.exe），再用
   --python 那個路徑重新執行，不要自己亂猜或亂裝套件。

4. 執行完告訴我結果，並提醒我要完全關閉再重新打開 Claude Desktop。
```

把 `你的token` 換成步驟一領到的那組 token。**這組文字裡只有 token 是
機密，其他都是公開資訊**——貼給 agent 之前確認沒有把 token 貼漏或貼
給不信任的地方。

> 這個倉庫是公開的，任何 AI agent 都抓得到，不需要你的 GitHub 帳號
> 授權；裡面確定不含任何工具邏輯或機密，只有轉發程式碼跟讀寫 Aspen
> 節點的程式碼。

---

## 步驟三：執行安裝程式

如果你用方式 B 讓 AI agent 代勞，這一步已經做完，可以直接跳到步驟四。

打開命令提示字元，切到剛剛放檔案的資料夾，執行：

```bash
python -m pip install -r requirements.txt
python setup_student.py --url http://140.123.126.87:8787 --token 你的token
```

（如果你的 Aspen 環境不是系統預設的 `python`，可以用 `--python` 指定，
例如 `--python D:\Aspen_MCP\.venv\Scripts\python.exe`；記得 `pip install`
也要對著同一個直譯器跑）

這支程式會依序做四件事，任何一步失敗都會停下來並說明原因：

1. 找一個裝齊 `requirements.txt` 那些套件的 Python
2. 測試能不能連到雲端、token 對不對
3. 測試能不能叫得動這台電腦上的 Aspen
4. 把設定寫進 Claude Desktop 的設定檔
   （`%APPDATA%\Claude\claude_desktop_config.json`，只新增/更新
   `aspen` 這一項，其他既有的 MCP 設定不會被動到；原檔案會備份成
   `.json.bak`）

只想檢查環境、還不想寫入設定的話，加 `--check`：

```bash
python setup_student.py --check
```

---

## 步驟四：重新啟動 Claude Desktop

**完全關閉**再重新打開（不是縮到最小化）—— Claude Desktop 只在啟動時
讀取設定檔，不重開不會生效。

打開後在對話裡應該能看到 `aspen` 這個 MCP 工具已經連上。

---

## 之後怎麼更新

`bridge.py` 和 `agent.py` 會隨著課程修正而更新（例如 2026-09-11 起，
`run_simulation` 會把 Aspen Control Panel 的訊息一起帶回來；2026-10-30 起，匯出 `.inp`
會自動拿掉 Aspen 寫了卻讀不回來的純物質性質表，這些都需要新版的 `bridge.py`）。

**最簡單的做法：請 Claude 幫你更新。** 對 Claude 說「檢查 Aspen 工具有沒有新版」，
它會呼叫 `client_status` 比對你電腦上的檔案和課程目前的版本；有新版時它會先問你，
你同意後才執行 `client_update`：

- 這個資料夾是用 `git clone` 下載的、而且你沒改過檔案：它會 `git pull`；
- 否則：它只重新下載有變的檔案；
- 兩種做法更新完都會再和課程伺服器核對一次，不符就不改。

也可以自己做：重新下載這幾個檔案蓋掉原來的（或在資料夾裡 `git pull`）。

不論哪種做法，更新完都要 **完全關閉再重開 Claude Desktop**（用 Codex 等其他 MCP 用戶端的話，
一樣要重開，或在設定裡重新啟動 aspen 這個 MCP）—— `agent.py` 是啟動時載入的，不重開不會換到
新版。

雲端那一半由管理者更新，學生不用做任何事。想知道雲端目前是哪一版，在 Claude 裡呼叫
`my_usage`：回覆的 `server_version` 就是（這個查詢不扣額度）。

工具回覆裡若出現 `control_panel_capture` 不是 `ok`、`note` 說「橋接程式
是舊版」、或有 `client_update` / `client_outdated` 的字樣，就是還沒更新到新版。

---

## 遇到問題怎麼回報

請 Claude 幫你回報就好（它會呼叫 `report_problem`），不用自己寫信。它會先問你幾件事：
想做的是什麼製程、預期看到什麼、做了哪些步驟，這樣管理者才重現得出來。回報會自動帶上：

- 最近 30 次工具呼叫（雲端自己記的）；
- 你電腦的環境資訊：`agent.py`／`bridge.py` 的版本、Windows 版本、Aspen 行程狀態、工作日誌最後幾行；
  Aspen 當掉的話，還有 Windows 事件記錄裡最近 24 小時的 Aspen 崩潰；
- **模型檔：只有你明確說「可以」才會送**，而且只送目前開著的那份模型（或回報流程自己匯出的副本）。
  模型是你的作業內容，只有管理者登入後台才下載得到，保留 90 天後連同回報一起刪除。
  token 和名為 password 的內容送出前會被遮掉。

回報有每日上限（管理者在後台設定）；超過時照回覆說的等隔天。管理者會在後台看到你的回報並處理；
處理完的修正會隨著雲端更新上線，不需要你再做什麼。

---

## 疑難排解

| 安裝程式回報 | 原因 / 怎麼處理 |
|---|---|
| 找不到裝齊 requirements.txt 那些套件的 Python | 用 `--python` 指定正確的直譯器，或在某個環境跑 `pip install -r requirements.txt`（要跟三個 .py 檔同一個資料夾） |
| 雲端拒絕這組 token | token 複製時漏了字（重新複製一次），或帳號已被停權、token 已被重新產生 —— 聯絡課程管理者。查詢頁不會再顯示 token |
| （申請時）沒收到核准信、領取連結過期或已經用過 | 先看垃圾信匣；到 `/status` 輸入 email 再寄一次領取信。token 已經領過卻沒存到的話，只能請管理者重新產生 |
| 連不到雲端 | 確認你人在**中正大學**校內網路（目前只開放中正校內，其他學校還沒開放），且網址沒打錯 |
| 叫不動 Aspen | 這台電腦沒裝 Aspen Plus，或授權沒生效／過期 |
| （方式 B）AI agent 說抓不到 GitHub 網址 | 確認這台電腦能連外網；倉庫是公開的，不需要登入或授權 |
| Claude Desktop 裡看不到 aspen 工具 | 確認有**完全關閉**再重開，不是只是切到背景 |
| 工具回覆說 `reopen_required` 或要你重開檔案 | 在已跑過的模型上又加了元件，Aspen 之後會卡住；用 `open_project` 重新開同一個檔，再繼續 |
| 工具一直 `still_running` 或逾時 | 同樣的參數再呼叫一次；仍不行就把當天的 `logs\` 日誌附給管理者，或請 Claude `report_problem` |

用量或帳號問題（配額用完、忘記 token、需要調整身分）請聯絡課程管理者。

### 工作日誌

`agent.py` 會把每一次工具呼叫記在同一個資料夾的 `logs\agent-年-月-日.log`
（一天一個檔，保留 14 天）：
- 呼叫了哪支工具、花了幾秒；
- 慢的或失敗的 Aspen 動作，以及 Aspen 的原話；
- 連不上雲端、重試的經過。

日誌只存在你的電腦上，**不記 token**。回報「逾時」「卡住」這類問題時，請把
當天的日誌檔一起附上，並說明大約幾點發生。不想記錄的話，可以把環境變數
`ASPEN_LOG` 設成 `0`。

工具回覆裡出現 `still_running` 是正常的：那次工作還在 Aspen 裡跑，為了不讓
Claude 判逾時，會先回話。用同樣的參數再呼叫一次，就會拿到結果。
