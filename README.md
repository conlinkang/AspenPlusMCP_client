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
那些都在雲端伺服器上執行，這裡只負責「转发」跟「操作 Aspen」。

> **目前只服務國立中正大學（CCU）化工系課程，只能在中正校內網路
> 使用**（伺服器架在中正校內，暫不對外開放）。伺服器網址是校內私有
> IP，校外／宿舍網路連不進去，也**還沒有開放給其他學校的師生申請**
> —— 如果你不是中正的學生，這份說明目前對你沒有用，請聯絡你自己課程
> 的授課教師確認有沒有對應的雲端服務。

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

> 這一步只有**在中正大學校內網路**（教室、系上機房、校內 WiFi）才能
> 打得開。**目前只接受中正的申請**，其他學校的 email 雖然格式上填
> 得進去，但送出後也連不到這台伺服器 —— 這是暫時性的限制，不是刻意
> 排除，未來若開放對外會另行公告。

1. 在校內網路瀏覽器開啟：**http://192.168.50.113:8787/signup**
2. 填寫：
   - **學術 email**（用來收驗證信；目前限中正大學信箱）
   - **姓名**
   - **身分**：教師 / 學生
3. 送出後會收到一封驗證信，**24 小時內**點裡面的連結完成驗證
4. 驗證完成後，申請會送交課程管理者審核
5. 審核通過後（**不會另外寄信通知**），自己回到
   **http://192.168.50.113:8787/status**，輸入 email 查詢，
   會看到一組 **token**（**只顯示這一次**，請立刻複製保存 —— 沒存到
   只能請管理者重新產生一組）

每個帳號預設每月可呼叫工具 **1000 次**，每月自動重置。

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
   python setup_student.py --url http://192.168.50.113:8787 --token 你的token

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
python setup_student.py --url http://192.168.50.113:8787 --token 你的token
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
`run_simulation` 會把 Aspen Control Panel 的訊息一起帶回來，這需要新版的
`bridge.py`）。管理者通知有新版時：

1. 重新下載這三個檔案，蓋掉原來的（或在資料夾裡 `git pull`）。
2. **完全關閉再重開 Claude Desktop** —— `agent.py` 是 Claude 啟動時載入的，
   不重開不會換到新版。

工具回覆裡若出現 `control_panel_capture` 不是 `ok`、或 `note` 說「橋接程式
是舊版」，就是還沒更新到新版。

---

## 疑難排解

| 安裝程式回報 | 原因 / 怎麼處理 |
|---|---|
| 找不到裝齊 requirements.txt 那些套件的 Python | 用 `--python` 指定正確的直譯器，或在某個環境跑 `pip install -r requirements.txt`（要跟三個 .py 檔同一個資料夾） |
| 雲端拒絕這組 token | token 打錯字，或帳號已被停權 —— 回 `/status` 用 email 重新查一次，或聯絡課程管理者 |
| 連不到雲端 | 確認你人在**中正大學**校內網路（目前只開放中正校內，其他學校還沒開放），且網址沒打錯 |
| 叫不動 Aspen | 這台電腦沒裝 Aspen Plus，或授權沒生效／過期 |
| （方式 B）AI agent 說抓不到 GitHub 網址 | 確認這台電腦能連外網；倉庫是公開的，不需要登入或授權 |
| Claude Desktop 裡看不到 aspen 工具 | 確認有**完全關閉**再重開，不是只是切到背景 |

用量或帳號問題（配額用完、忘記 token、需要調整身分）請聯絡課程管理者。
