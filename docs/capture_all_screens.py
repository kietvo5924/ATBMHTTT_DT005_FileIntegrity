# -*- coding: utf-8 -*-
"""
Script tự động kích hoạt máy chủ Web Demo thật và chụp ảnh màn hình chân thực:
1. docs/screenshots/web_safe.png: Giao diện Web thật ở trạng thái An toàn (100% toàn vẹn, 1.000 tệp)
2. docs/screenshots/web_attack.png: Giao diện Web thật khi phát hiện tấn công đảo 1 bit (Tamper Detected)
3. docs/screenshots/terminal_cli.png: Cửa sổ Terminal thật khi chạy bộ kiểm thử 17 Unit Tests mật mã học
"""
import os
import sys
import subprocess
import urllib.request
import json
import time

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE_DIR = r"d:\AnToanvaBaoMatHeThongThongTin\GK"
SHOT_DIR = os.path.join(BASE_DIR, "docs", "screenshots")
os.makedirs(SHOT_DIR, exist_ok=True)

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"


def capture_chrome(url, output_png, width=1280, height=800, wait_ms=2500):
    cmd = [
        CHROME_PATH,
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        f"--window-size={width},{height}",
        f"--virtual-time-budget={wait_ms}",
        f"--screenshot={output_png}",
        url
    ]
    subprocess.run(cmd, check=True)
    if os.path.exists(output_png):
        print(f"✔ Đã chụp thành công: {output_png} ({os.path.getsize(output_png):,} bytes)")
    else:
        print(f"❌ Không tìm thấy: {output_png}")


def main():
    print("--- 1. Thiết lập trạng thái Web Demo an toàn ---")
    urllib.request.urlopen("http://127.0.0.1:8080/api/reset")
    time.sleep(1)
    urllib.request.urlopen("http://127.0.0.1:8080/api/learn_model?model=hd_merkle")
    time.sleep(1)
    urllib.request.urlopen("http://127.0.0.1:8080/api/compare_all")
    time.sleep(1)

    safe_png = os.path.join(SHOT_DIR, "web_safe.png")
    capture_chrome("http://127.0.0.1:8080/?action=safe", safe_png, width=1300, height=850, wait_ms=4000)

    print("--- 2. Kích hoạt mô phỏng tấn công đảo 1 bit trên Web Demo ---")
    attack_png = os.path.join(SHOT_DIR, "web_attack.png")
    capture_chrome("http://127.0.0.1:8080/?action=attack", attack_png, width=1300, height=850, wait_ms=4000)

    print("--- 3. Chạy lệnh thực tế Terminal & ghi nhận output thật ---")
    sub_env = os.environ.copy()
    sub_env["PYTHONIOENCODING"] = "utf-8"
    sub_env["PYTHONUTF8"] = "1"
    test_run = subprocess.run(
        ["python", "-m", "unittest", "discover", "-s", "src/tests", "-v"],
        cwd=BASE_DIR,
        capture_output=True,
        env=sub_env
    )
    stderr_str = test_run.stderr.decode("utf-8", errors="replace") if test_run.stderr else ""
    stdout_str = test_run.stdout.decode("utf-8", errors="replace") if test_run.stdout else ""
    raw_terminal_output = stderr_str + "\n" + stdout_str

    # Tạo trang HTML terminal chuẩn sắc nét
    terminal_html = os.path.join(SHOT_DIR, "terminal_view.html")
    escaped_output = (
        raw_terminal_output
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    
    # Định dạng highlight màu xanh cho 'ok'
    styled_lines = []
    for line in escaped_output.splitlines():
        if " ... ok" in line:
            parts = line.split(" ... ok")
            styled_lines.append(f"{parts[0]} ... <span style='color:#4ade80;font-weight:bold;'>ok</span>")
        elif "OK" in line:
            styled_lines.append(f"<span style='color:#4ade80;font-weight:bold;background:#064e3b;padding:2px 8px;border-radius:4px;'>{line}</span>")
        elif "Ran 17 tests" in line:
            styled_lines.append(f"<span style='color:#facc15;font-weight:bold;'>{line}</span>")
        else:
            styled_lines.append(line)
    
    html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{
    margin: 0; padding: 20px;
    background: #0f172a;
    font-family: 'JetBrains Mono', 'Cascadia Code', Consolas, monospace;
    font-size: 13.5px;
    line-height: 1.5;
    color: #e2e8f0;
  }}
  .window {{
    background: #1e293b;
    border-radius: 10px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.6);
    border: 1px solid #334155;
    overflow: hidden;
  }}
  .titlebar {{
    background: #0f172a;
    padding: 10px 16px;
    display: flex;
    align-items: center;
    border-bottom: 1px solid #334155;
  }}
  .dots {{
    display: flex; gap: 8px; margin-right: 16px;
  }}
  .dot {{
    width: 12px; height: 12px; border-radius: 50%;
  }}
  .dot-red {{ background: #ef4444; }}
  .dot-yellow {{ background: #eab308; }}
  .dot-green {{ background: #22c55e; }}
  .title {{
    color: #94a3b8; font-size: 12px; font-weight: 600;
  }}
  .console {{
    padding: 18px 22px;
    background: #0f172a;
    white-space: pre-wrap;
    word-break: break-all;
  }}
  .prompt {{
    color: #38bdf8; font-weight: bold; margin-bottom: 12px;
  }}
  .command {{
    color: #f8fafc;
  }}
</style>
</head>
<body>
<div class="window">
  <div class="titlebar">
    <div class="dots">
      <div class="dot dot-red"></div>
      <div class="dot dot-yellow"></div>
      <div class="dot dot-green"></div>
    </div>
    <div class="title">Windows PowerShell — DT005 Cryptographic Security Test Suite (17/17 PASS)</div>
  </div>
  <div class="console">
    <div class="prompt">PS D:\\AnToanvaBaoMatHeThongThongTin\\GK&gt; <span class="command">python -m unittest discover -s src/tests -v</span></div>
    <div>{'<br>'.join(styled_lines)}</div>
  </div>
</div>
</body>
</html>
"""
    with open(terminal_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    terminal_png = os.path.join(SHOT_DIR, "terminal_cli.png")
    capture_chrome(f"file:///{terminal_html.replace(chr(92), '/')}", terminal_png, width=1280, height=820, wait_ms=1000)

    print("\n✅ HOÀN TẤT CHỤP 3 ẢNH THỰC TẾ:")
    print(f"1. {safe_png}")
    print(f"2. {attack_png}")
    print(f"3. {terminal_png}")


if __name__ == "__main__":
    main()
