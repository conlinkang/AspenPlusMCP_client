"""學生端安裝程式。

學生執行這一支，它會：
  1. 找一個裝了 pywin32 與 mcp 的 Python
  2. 測試能不能連到雲端、token 對不對
  3. 測試能不能叫得動這台電腦上的 Aspen
  4. 把設定寫進 Claude Desktop 的設定檔

四件事任一失敗就停下來，並說清楚要怎麼修 —— 半套的安裝比沒安裝更難查。

用法：
    python setup_student.py --url https://aspen.example.net --token 你的token
    python setup_student.py --check          只檢查，不寫任何設定
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
NEEDED = ("win32com", "mcp")

# Claude Desktop 的設定檔位置。Windows 只有這一個。
CONFIG = (Path(os.environ.get("APPDATA", ""))
          / "Claude" / "claude_desktop_config.json")


def say(ok: bool, text: str, detail: str = "") -> bool:
    print("{} {}".format("[OK]  " if ok else "[失敗]", text))
    if detail:
        for line in str(detail).splitlines():
            print("       " + line)
    return ok


# ── 1. 找可用的 Python ──────────────────────────────────────────────
def probe(python: Path) -> list:
    """回傳這個 Python 缺哪些套件。空清單代表都有。"""
    code = ("import importlib.util as u, json, sys;"
            "print(json.dumps([m for m in {!r} if u.find_spec(m) is None]))"
            .format(list(NEEDED)))
    try:
        out = subprocess.run([str(python), "-c", code], capture_output=True,
                             text=True, timeout=60)
        return json.loads(out.stdout.strip() or "[]")
    except Exception:
        return list(NEEDED)


def candidates(explicit: str | None) -> list:
    if explicit:
        return [Path(explicit)]
    found = []
    # Aspen 自己的環境通常已經有 pywin32，優先試
    for guess in (r"D:\Aspen_MCP\.venv\Scripts\python.exe",
                  r"C:\Aspen_MCP\.venv\Scripts\python.exe"):
        if Path(guess).is_file():
            found.append(Path(guess))
    found.append(Path(sys.executable))
    which = shutil.which("python")
    if which:
        found.append(Path(which))
    seen, unique = set(), []
    for path in found:
        key = str(path).lower()
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def pick_python(explicit: str | None):
    print("尋找可用的 Python…")
    problems = []
    for python in candidates(explicit):
        missing = probe(python)
        if not missing:
            say(True, "Python：{}".format(python))
            return python, None
        problems.append("{}：缺 {}".format(python, "、".join(missing)))
    say(False, "找不到同時裝有 pywin32 與 mcp 的 Python",
        "\n".join(problems)
        + "\n\n請在其中一個環境安裝：\n"
          "    <該 python> -m pip install pywin32 mcp")
    return None, problems


# ── 2. 測雲端 ───────────────────────────────────────────────────────
def check_cloud(url: str, token: str) -> bool:
    print("測試雲端連線…")
    req = urllib.request.Request(
        url.rstrip("/") + "/tools",
        data=json.dumps({}).encode("utf-8"),
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer " + token})
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        return say(True, "雲端連線正常",
                   "可用工具 {} 個".format(len(payload.get("tools") or [])))
    except urllib.error.HTTPError as exc:
        if exc.code == 401:
            return say(False, "雲端拒絕這組 token",
                       "token 可能打錯、或帳號已停權。請找課程助教確認。")
        return say(False, "雲端回應 HTTP {}".format(exc.code))
    except Exception as exc:
        return say(False, "連不到雲端",
                   "{}\n\n請確認網址正確、而且這台電腦可以連外。".format(exc))


# ── 3. 測 Aspen ─────────────────────────────────────────────────────
def check_aspen(python: Path) -> bool:
    print("測試這台電腦上的 Aspen…")
    code = ("import win32com.client as w;"
            "a=w.Dispatch('Apwn.Document');print('OK')")
    try:
        out = subprocess.run([str(python), "-c", code], capture_output=True,
                             text=True, timeout=120)
    except Exception as exc:
        return say(False, "測試 Aspen 時出錯", repr(exc))
    if "OK" in (out.stdout or ""):
        return say(True, "Aspen 可以叫得動")
    return say(False, "叫不動 Aspen",
               (out.stderr or "").strip()[-400:]
               + "\n\n請確認這台電腦已安裝 Aspen Plus 並且啟用過授權。")


# ── 4. 寫設定 ───────────────────────────────────────────────────────
def write_config(python: Path, url: str, token: str, dry: bool) -> bool:
    entry = {
        "command": str(python),
        "args": [str(HERE / "agent.py")],
        "env": {"ASPEN_CLOUD": url.rstrip("/"), "ASPEN_TOKEN": token},
    }

    existing = {}
    if CONFIG.is_file():
        try:
            existing = json.loads(CONFIG.read_text(encoding="utf-8"))
        except ValueError:
            return say(False, "現有的設定檔不是合法的 JSON",
                       "請先備份並修正 {}".format(CONFIG))

    # 只動 aspen 這一項，其他 MCP 伺服器原封不動 ——
    # 學生可能已經裝了別的工具，整份覆蓋會把那些弄不見。
    servers = existing.setdefault("mcpServers", {})
    replacing = "aspen" in servers
    servers["aspen"] = entry

    if dry:
        print()
        print("（--check 模式，沒有寫入任何檔案）預計寫入：")
        print(json.dumps({"mcpServers": {"aspen": entry}},
                         ensure_ascii=False, indent=2))
        return True

    try:
        CONFIG.parent.mkdir(parents=True, exist_ok=True)
        if CONFIG.is_file():
            shutil.copy2(CONFIG, CONFIG.with_suffix(".json.bak"))
        CONFIG.write_text(json.dumps(existing, ensure_ascii=False, indent=2),
                          encoding="utf-8")
    except OSError as exc:
        return say(False, "寫不進設定檔", "{}\n{}".format(CONFIG, exc))

    return say(True, "{}設定完成".format("更新" if replacing else "寫入"),
               "{}\n其他 MCP 設定未更動；原檔已備份為 .json.bak".format(CONFIG))


def main() -> int:
    parser = argparse.ArgumentParser(description="Aspen MCP 學生端安裝")
    parser.add_argument("--url", help="雲端網址")
    parser.add_argument("--token", help="你的 token")
    parser.add_argument("--python", help="指定要用的 python.exe")
    parser.add_argument("--check", action="store_true",
                        help="只檢查，不寫入設定")
    args = parser.parse_args()

    if not args.check and (not args.url or not args.token):
        parser.error("要寫入設定就必須提供 --url 與 --token")

    print("Aspen MCP 學生端安裝")
    print("=" * 46)

    python, _ = pick_python(args.python)
    if python is None:
        return 1

    if args.url and args.token:
        if not check_cloud(args.url, args.token):
            return 1

    if not check_aspen(python):
        return 1

    if args.url and args.token:
        if not write_config(python, args.url, args.token, args.check):
            return 1

    print()
    print("=" * 46)
    if args.check:
        print("檢查完畢，沒有寫入任何東西。")
    else:
        print("安裝完成。請**完全關閉 Claude Desktop 再重新開啟** ——")
        print("它只在啟動時讀設定檔，不重開不會生效。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
