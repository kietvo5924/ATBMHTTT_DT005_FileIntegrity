# -*- coding: utf-8 -*-
"""
Red Team - The Adversary Web Dashboard (Port 8081)
Giao diện Web sáng (Light Theme) độc lập dành cho Kẻ tấn công (The Adversary - Harry Halpin 2025).
Đối đầu trực tiếp với Blue Team (Port 8080).
Không sử dụng biểu tượng (icon) hoặc emoji, giao diện tinh gọn, chuyên nghiệp, chuẩn học thuật.
"""
import os
import sys
import json
import time
from urllib.parse import urlparse, parse_qs, unquote
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from src.adversary.adversary_engine import AdversaryEngine

ENGINE = AdversaryEngine()


HTML_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>DT005 - Bảng Điều Khiển Tấn Công Đối Kháng Độc Lập</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --card-border: #e2e8f0;
      --card-border-hover: #cbd5e1;
      --text-main: #0f172a;
      --text-sub: #334155;
      --text-muted: #64748b;
      --red-primary: #b91c1c;
      --red-hover: #991b1b;
      --red-light: #fef2f2;
      --red-border: #fecaca;
      --green-primary: #059669;
      --green-hover: #047857;
      --green-light: #ecfdf5;
      --green-border: #a7f3d0;
      --badge-bg: #f1f5f9;
      --badge-color: #475569;
      --badge-border: #e2e8f0;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background: var(--bg);
      color: var(--text-main);
      font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
      padding: 24px;
      line-height: 1.5;
    }
    .container { max-width: 1400px; margin: 0 auto; }
    header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 20px 24px;
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-left: 4px solid var(--red-primary);
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
      margin-bottom: 20px;
    }
    .brand { display: flex; flex-direction: column; gap: 4px; }
    .brand-tag {
      font-family: 'JetBrains Mono', monospace;
      font-size: 11px;
      font-weight: 700;
      color: var(--red-primary);
      letter-spacing: 0.5px;
      text-transform: uppercase;
    }
    .title h1 {
      font-size: 20px;
      font-weight: 800;
      letter-spacing: -0.3px;
      color: var(--text-main);
    }
    .title span {
      font-size: 12px;
      color: var(--text-muted);
      font-weight: 500;
    }
    .target-bar {
      display: flex;
      gap: 16px;
      align-items: center;
      background: var(--card-bg);
      padding: 14px 20px;
      border-radius: 8px;
      border: 1px solid var(--card-border);
      margin-bottom: 20px;
      flex-wrap: wrap;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
    }
    .target-item { display: flex; align-items: center; gap: 8px; font-size: 13px; color: var(--text-sub); }
    .badge {
      padding: 3px 8px;
      border-radius: 4px;
      font-weight: 600;
      font-family: 'JetBrains Mono', monospace;
      font-size: 11px;
      background: var(--badge-bg);
      color: var(--badge-color);
      border: 1px solid var(--badge-border);
      display: inline-flex;
      align-items: center;
    }

    .grid { display: grid; grid-template-columns: repeat(12, 1fr); gap: 16px; margin-bottom: 20px; }
    .card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 8px;
      padding: 18px;
      transition: all 0.2s ease;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }
    .card:hover {
      border-color: var(--card-border-hover);
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .card-top { margin-bottom: 12px; }
    .card-title {
      font-size: 14px;
      font-weight: 700;
      color: var(--text-main);
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }
    .card-desc {
      font-size: 12px;
      color: var(--text-muted);
      line-height: 1.45;
    }
    .card-actions { margin-top: auto; }
    .btn {
      padding: 9px 14px;
      border-radius: 6px;
      font-weight: 600;
      font-size: 12px;
      cursor: pointer;
      border: 1px solid transparent;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      transition: all 0.15s ease;
      font-family: 'Plus Jakarta Sans', sans-serif;
    }
    .btn-attack {
      background: var(--red-primary);
      color: #ffffff;
      border: 1px solid #991b1b;
      width: 100%;
    }
    .btn-attack:hover {
      background: var(--red-hover);
      box-shadow: 0 2px 4px rgba(185, 28, 28, 0.2);
    }
    .btn-rollback {
      background: var(--green-primary);
      color: #ffffff;
      border: 1px solid #047857;
      font-size: 13px;
      padding: 10px 18px;
      font-weight: 700;
    }
    .btn-rollback:hover {
      background: var(--green-hover);
      box-shadow: 0 2px 4px rgba(5, 150, 105, 0.2);
    }
    .btn-master-attack {
      background: #991b1b;
      color: #ffffff;
      border: 1px solid #7f1d1d;
      font-size: 13px;
      padding: 10px 18px;
      font-weight: 700;
    }
    .btn-master-attack:hover {
      background: #7f1d1d;
      box-shadow: 0 2px 6px rgba(153, 27, 27, 0.35);
    }

    .console-card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 8px;
      padding: 18px;
      box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
    }
    .console-box {
      background: #f8fafc;
      border: 1px solid var(--card-border);
      border-radius: 6px;
      padding: 14px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 12px;
      height: 260px;
      overflow-y: auto;
      color: var(--text-sub);
    }
    .log-line {
      margin-bottom: 8px;
      padding-bottom: 8px;
      border-bottom: 1px solid #e2e8f0;
    }
    .log-time { color: var(--text-muted); margin-right: 8px; }
    .log-success { color: var(--green-primary); font-weight: 700; }
    .log-fail { color: var(--red-primary); font-weight: 700; }
    .log-details {
      margin-top: 6px;
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 4px;
      padding: 8px;
      color: #1e293b;
      font-size: 11px;
      overflow-x: auto;
    }

    .input-field {
      width: 100%;
      background: #ffffff;
      border: 1px solid var(--card-border);
      color: var(--text-main);
      padding: 8px 12px;
      border-radius: 6px;
      font-family: 'JetBrains Mono', monospace;
      font-size: 12px;
      margin-bottom: 10px;
      outline: none;
      transition: border-color 0.15s ease;
    }
    .input-field:focus { border-color: #2563eb; }
    .select-field {
      width: 100%;
      background: #ffffff;
      border: 1px solid var(--card-border);
      color: var(--text-main);
      padding: 8px 12px;
      border-radius: 6px;
      font-size: 12px;
      margin-bottom: 10px;
      outline: none;
      transition: border-color 0.15s ease;
    }
    .select-field:focus { border-color: #2563eb; }
  </style>
</head>
<body>
  <!-- MODAL CHỌN THƯ MỤC TRỰC QUAN (TƯƠNG TỰ BLUE TEAM) -->
  <div id="folderModal" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15, 23, 42, 0.65); z-index:99999; backdrop-filter:blur(4px); align-items:center; justify-content:center;">
    <div style="background:#fff; border-radius:12px; width:92%; max-width:540px; box-shadow:0 25px 50px -12px rgba(0,0,0,0.25); border:1px solid #e2e8f0; overflow:hidden;">
      <div style="padding:16px 20px; border-bottom:1px solid #e2e8f0; display:flex; align-items:center; justify-content:space-between; background:#f8fafc;">
        <h3 style="margin:0; font-size:15px; font-weight:700; color:#0f172a;">Chọn thư mục mục tiêu tấn công / phá hoại</h3>
        <button onclick="closeFolderModal()" style="background:transparent; border:none; font-size:22px; cursor:pointer; color:#64748b; line-height:1; padding:0 4px;">&times;</button>
      </div>
      
      <div style="padding:22px 20px; display:flex; flex-direction:column; gap:18px;">
        <!-- LỰA CHỌN 1: DUYỆT THƯ MỤC TRÊN MÁY BẰNG BROWSER NATIVE DIALOG -->
        <div>
          <div style="font-size:12px; font-weight:700; color:#334155; margin-bottom:8px;">1. Duyệt thư mục từ máy tính:</div>
          <input type="file" id="browserFolderPickerInput" webkitdirectory directory style="display:none;" onchange="handleNativeFolderSelect(this)">
          <button id="btnBrowseFolder" onclick="document.getElementById('browserFolderPickerInput').click()" class="btn" style="width:100%; padding:12px 16px; border:1.5px dashed #b91c1c; background:#fef2f2; border-radius:8px; color:#991b1b; font-weight:700; font-size:13px; cursor:pointer; text-align:center; margin-bottom:4px;">
            Duyệt thư mục trên máy tính (Mở cửa sổ chọn thư mục)
          </button>
          <div style="font-size:11.5px; color:#64748b; padding-left:2px;">Mở hộp thoại Explorer của hệ điều hành để duyệt và chọn thư mục cần tấn công/phá hoại.</div>
        </div>

        <!-- LỰA CHỌN 2: NHẬP ĐƯỜNG DẪN TÙY Ý -->
        <div style="border-top:1px solid #f1f5f9; padding-top:16px;">
          <div style="font-size:12px; font-weight:700; color:#334155; margin-bottom:8px;">2. Hoặc nhập đường dẫn thư mục:</div>
          <div style="display:flex; gap:8px;">
            <input type="text" id="modalPathInput" placeholder="Ví dụ: kaggle_dataset hoặc realistic_dataset hoặc C:/..." style="flex:1; padding:9px 12px; border-radius:7px; border:1px solid #cbd5e1; font-family:'JetBrains Mono', monospace; font-size:12px;">
            <button id="btnApplyModal" onclick="applyModalPath(this)" class="btn btn-attack" style="width:auto; margin:0; padding:9px 18px; font-size:12.5px; font-weight:700;">Áp dụng</button>
          </div>
          <div style="font-size:11.5px; color:#64748b; margin-top:6px; padding-left:2px;">
            Nhập tên thư mục tương đối hoặc đường dẫn tuyệt đối trên đĩa cứng.
          </div>
        </div>
      </div>

      <div style="padding:12px 20px; background:#f8fafc; border-top:1px solid #e2e8f0; display:flex; justify-content:flex-end;">
        <button onclick="closeFolderModal()" class="btn" style="background:#fff; border:1px solid #cbd5e1; color:#475569; width:auto; margin:0; padding:6px 16px; font-size:12.5px; cursor:pointer;">Đóng</button>
      </div>
    </div>
  </div>

  <div class="container">
    <header>
      <div class="brand">
        <div class="brand-tag">Bảng Điều Khiển Tấn Công Đối Kháng Độc Lập</div>
        <div class="title">
          <h1>Mô Phỏng Kẻ Tấn Công Harry Halpin</h1>
          <span>Hiện thực hóa toàn văn công trình nghiên cứu Harry Halpin xuất bản tại Oxford Journal of Cybersecurity 2025</span>
        </div>
      </div>
      <div style="display:flex; gap:12px; align-items:center;">
        <button class="btn btn-master-attack" id="btnLaunchAll7" onclick="launchAll7Attacks()">
          Phát động đồng loạt 7 đòn tấn công
        </button>
        <button class="btn btn-rollback" onclick="executeRollback()">
          Khôi phục nguyên trạng toàn bộ
        </button>
      </div>
    </header>

    <div class="target-bar" style="display:flex; flex-direction:column; gap:12px;">
      <!-- HÀNG 1: THÔNG TIN KẾT NỐI VÀ BADGES TRẠNG THÁI -->
      <div style="display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; width:100%;">
        <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
          <div class="target-item">
            <span style="font-weight:600; color:var(--text-sub);">Blue Team:</span>
            <a href="http://127.0.0.1:8080" target="_blank" style="color: #2563eb; font-weight:700; font-family:'JetBrains Mono', monospace; text-decoration:none;">http://127.0.0.1:8080</a>
          </div>
          <button class="btn" style="background:#eff6ff; color:#1d4ed8; border:1px solid #bfdbfe; padding:4px 10px; font-size:11.5px; font-weight:700; cursor:pointer;" onclick="syncFromBlueTeam()">
            Đồng bộ từ Blue Team
          </button>
          <button class="btn" style="background:#f8fafc; color:#475569; border:1px solid #cbd5e1; padding:4px 10px; font-size:11.5px; font-weight:600; cursor:pointer;" onclick="openInExplorer()">
            Mở thư mục
          </button>
        </div>
        <div class="target-item" style="display:flex; gap:8px; flex-wrap:wrap;">
          <span class="badge" id="badgeFileCount" style="background:#f1f5f9; color:#334155; border:1px solid #cbd5e1;">0 tệp tin</span>
          <span class="badge" id="badgeTampered" style="background:#fef2f2; color:#991b1b; border:1px solid #fecaca;">Sửa đổi: 0</span>
          <span class="badge" id="badgeInjected" style="background:#f5f3ff; color:#6b21a8; border:1px solid #e9d5ff;">Xâm nhập: 0</span>
          <span class="badge" id="badgeDeleted" style="background:#fffbeb; color:#92400e; border:1px solid #fde68a;">Đã xóa: 0</span>
        </div>
      </div>

      <!-- HÀNG 2: THANH CHỌN VÀ NHẬP ĐƯỜNG DẪN THƯ MỤC MỤC TIÊU (GIỐNG BLUE TEAM) -->
      <div style="display:flex; align-items:center; gap:10px; width:100%; flex-wrap:wrap;">
        <span style="font-size:13px; font-weight:700; color:var(--text-main); white-space:nowrap;">Thư mục mục tiêu:</span>
        <button id="btnSelectFolder" class="btn" style="background:var(--red-primary); color:#ffffff; border:1px solid #991b1b; font-weight:700; font-size:12.5px; padding:7px 16px; cursor:pointer;" onclick="openFolderModal()">
          Chọn thư mục
        </button>
        <input type="text" id="targetDirInput" value="" style="flex:1; min-width:260px; padding:7px 12px; border-radius:6px; border:1px solid var(--card-border); font-family:'JetBrains Mono', monospace; font-size:12px; background:#fff; outline:none;" placeholder="Chưa chọn thư mục (Bấm 'Chọn thư mục' để chọn)..." onchange="applyTargetDirInput()" onkeydown="if(event.key==='Enter') applyTargetDirInput()">
        <button id="btnClearFolder" onclick="clearTargetFolder()" style="display:none; background:transparent; border:none; color:#ef4444; font-size:12px; font-weight:700; cursor:pointer; padding:4px 6px;" title="Hủy chọn thư mục">✕ Bỏ chọn</button>
      </div>
    </div>

    <!-- 8 ATTACK VECTOR LAUNCHPADS -->
    <div class="grid">
      <!-- Vector 1: Tamper File -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>1. Phá hoại và sửa đổi tệp</span>
            <span class="badge">Can thiệp tệp</span>
          </div>
          <p class="card-desc">
            Ghi đè nội dung độc hại vào một tệp trên đĩa cứng. Hệ thống tự động sao lưu bản gốc và kích hoạt cảnh báo đường dẫn bẩn trên Blue Team.
          </p>
        </div>
        <div class="card-actions">
          <select id="selectTamperFile" class="select-field">
            <option value="">Tự động chọn tệp mẫu...</option>
          </select>
          <button class="btn btn-attack" onclick="attackTamper()">Thực hiện phá hoại tệp</button>
        </div>
      </div>

      <!-- Vector 2: Inject Backdoor -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>2. Chèn tệp lạ và mã độc</span>
            <span class="badge">Chèn tệp lạ</span>
          </div>
          <p class="card-desc">
            Thả tệp lạ mới vào thư mục. Phá vỡ tính toàn vẹn danh mục và kiểm định tính nhất quán chỉ thêm mới.
          </p>
        </div>
        <div class="card-actions">
          <input type="text" id="inputBackdoorName" class="input-field" placeholder="Tên tệp ví dụ backdoor_exploit.py" value="backdoor_exploit.py">
          <button class="btn btn-attack" onclick="attackInject()">Thực hiện chèn tệp lạ</button>
        </div>
      </div>

      <!-- Vector 3: Silent Delete -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>3. Xóa lén tệp dữ liệu</span>
            <span class="badge">Xóa lén tệp</span>
          </div>
          <p class="card-desc">
            Xóa ngầm một tệp khỏi đĩa cứng. Blue Team khi quét toàn vẹn sẽ lập tức phát hiện tệp bị thiếu và tính lại gốc Merkle.
          </p>
        </div>
        <div class="card-actions">
          <select id="selectDeleteFile" class="select-field">
            <option value="">Tự động chọn tệp cuối...</option>
          </select>
          <button class="btn btn-attack" onclick="attackDelete()">Thực hiện xóa tệp</button>
        </div>
      </div>

      <!-- Vector 4: CRC32 Linear Forge -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>4. Bẻ gãy tuyến tính CRC32</span>
            <span class="badge">Bù đại số GF2</span>
          </div>
          <p class="card-desc">
            Minh chứng Halpin: Phân biệt nhiễu ngẫu nhiên với kẻ tấn công thông minh. Tự động giải 4 byte bù đại số qua mặt hoàn toàn CRC32 nhưng bị phát hiện bởi SHA-256.
          </p>
        </div>
        <div class="card-actions">
          <input type="text" id="inputCrc32Msg" class="input-field" value="Transaction: Transfer 999999 USD to HACKER_RED_TEAM">
          <button class="btn btn-attack" onclick="attackCrc32()">Giải đại số bẻ gãy CRC32</button>
        </div>
      </div>

      <!-- Vector 5: Dolev-Yao MitM -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>5. Đánh tráo bằng chứng trên mạng</span>
            <span class="badge">Can thiệp mạng Dolev-Yao</span>
          </div>
          <p class="card-desc">
            Kẻ thù can thiệp trên kênh truyền mạng theo mô hình Dolev-Yao. Chặn bắt bằng chứng Merkle của tệp B gán cho tệp A nhằm thử thách bộ xác thực.
          </p>
        </div>
        <div class="card-actions">
          <button class="btn btn-attack" onclick="attackDolevYao()">Tráo đổi bằng chứng Merkle</button>
        </div>
      </div>

      <!-- Vector 6: Second-Preimage Exploit -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>6. Tấn công tiền ảnh thứ hai</span>
            <span class="badge">Tiền ảnh thứ hai</span>
          </div>
          <p class="card-desc">
            Chế tạo nút lá giả mạo ghép chuỗi nhị phân. Mô hình Merkle chuẩn bị qua mặt trong khi HD-MerkleTree phân biệt miền dữ liệu ngăn chặn hoàn toàn.
          </p>
        </div>
        <div class="card-actions">
          <button class="btn btn-attack" onclick="attackSecondPreimage()">Khai thác tiền ảnh thứ hai</button>
        </div>
      </div>

      <!-- Vector 7: Timing Side-Channel Probe -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>7. Đo đạc kênh phụ thời gian</span>
            <span class="badge">Kênh phụ thời gian</span>
          </div>
          <p class="card-desc">
            Minh chứng Halpin về giới hạn hình thức hóa: Đo đạc rò rỉ nano giây của so khớp tuần tự sớm dừng so với so khớp thời gian không đổi HMAC.
          </p>
        </div>
        <div class="card-actions">
          <button class="btn btn-attack" onclick="attackTiming()">Đo đạc rò rỉ thời gian</button>
        </div>
      </div>

      <!-- Vector 8: Traffic & Metadata Reconnaissance -->
      <div class="card" style="grid-column: span 3;">
        <div class="card-top">
          <div class="card-title">
            <span>8. Thu thập siêu dữ liệu và lưu lượng</span>
            <span class="badge">Phân tích lưu lượng</span>
          </div>
          <p class="card-desc">
            Minh chứng Halpin: Phân tích siêu dữ liệu và lưu lượng là nền tảng tình báo tín hiệu. Trích xuất cấu trúc thư mục và kích thước tệp mà không cần giải mã nội dung.
          </p>
        </div>
        <div class="card-actions">
          <button class="btn btn-attack" onclick="attackMetadata()">Phân tích siêu dữ liệu</button>
        </div>
      </div>
    </div>

    <!-- LIVE CONSOLE TERMINAL LOG -->
    <div class="console-card">
      <div class="card-title" style="margin-bottom: 12px;">
        <span style="font-family:'JetBrains Mono', monospace; font-weight:700; color:var(--text-main); font-size:13px;">Nhật ký thực thi và sự kiện tấn công thời gian thực</span>
        <button class="btn" style="background:#f1f5f9; color:var(--text-sub); border:1px solid #cbd5e1; padding:5px 12px;" onclick="fetchStatus()">Làm mới nhật ký</button>
      </div>
      <div class="console-box" id="consoleBox">
        <div class="log-line"><span class="log-time">00:00:00</span> Đang kết nối Red Team Adversary Engine...</div>
      </div>
    </div>
  </div>

  <script>
    let currentTargetDir = '';
    let lastRenderedDir = null;

    function openFolderModal() {
      const m = document.getElementById('folderModal');
      if (m) m.style.display = 'flex';
      const input = document.getElementById('modalPathInput');
      if (input) input.value = currentTargetDir || '';
    }

    function closeFolderModal() {
      const m = document.getElementById('folderModal');
      if (m) m.style.display = 'none';
    }

    async function applyModalPath(btn) {
      const input = document.getElementById('modalPathInput');
      const val = input ? input.value.trim() : '';
      if (!val) {
        alert('Vui lòng nhập đường dẫn thư mục!');
        return;
      }
      closeFolderModal();
      await applyFolderPath(val);
    }

    async function handleNativeFolderSelect(fileInput) {
      if (!fileInput.files || fileInput.files.length === 0) return;
      const relPath = fileInput.files[0].webkitRelativePath || '';
      const parts = relPath.replace(/\\\\/g, '/').split('/');
      const folderName = parts[0] || '';
      closeFolderModal();
      if (folderName) {
        await applyFolderPath(folderName);
      }
    }

    async function applyTargetDirInput() {
      const input = document.getElementById('targetDirInput');
      const p = input.value.trim();
      if (!p) {
        await clearTargetFolder();
        return;
      }
      await applyFolderPath(p);
    }

    async function applyFolderPath(folderPath) {
      try {
        const res = await fetch(`/api/workspace?action=set&path=${encodeURIComponent(folderPath)}`);
        const data = await res.json();
        if (data.status === 'ok') {
          currentTargetDir = data.target_dir;
          const targetInput = document.getElementById('targetDirInput');
          if (targetInput) targetInput.value = currentTargetDir;
          const clearBtn = document.getElementById('btnClearFolder');
          if (clearBtn) clearBtn.style.display = 'inline-block';
          showToast(data.message || `Đã chọn thư mục: ${currentTargetDir}`, true);
          await fetchStatus(true);
        } else {
          alert('Lỗi: ' + (data.message || 'Không thể chọn thư mục'));
        }
      } catch (e) {
        console.error(e);
        alert('Không thể kết nối đến máy chủ Red Team!');
      }
    }

    async function clearTargetFolder() {
      try {
        const res = await fetch('/api/workspace?action=clear');
        const data = await res.json();
        currentTargetDir = '';
        const targetInput = document.getElementById('targetDirInput');
        if (targetInput) targetInput.value = '';
        const clearBtn = document.getElementById('btnClearFolder');
        if (clearBtn) clearBtn.style.display = 'none';
        showToast('Đã hủy chọn thư mục mục tiêu.', true);
        await fetchStatus(true);
      } catch (e) {
        console.error(e);
      }
    }

    async function syncFromBlueTeam() {
      try {
        const res = await fetch('/api/workspace?action=sync');
        const data = await res.json();
        if (data.status === 'ok') {
          currentTargetDir = data.target_dir;
          const targetInput = document.getElementById('targetDirInput');
          if (targetInput) targetInput.value = currentTargetDir;
          const clearBtn = document.getElementById('btnClearFolder');
          if (clearBtn) clearBtn.style.display = 'inline-block';
          showToast(data.message || 'Đã đồng bộ từ Blue Team!', true);
          await fetchStatus(true);
        } else {
          alert('Không thể đồng bộ: ' + (data.message || 'Blue Team chưa khởi động hoặc chưa chọn thư mục'));
        }
      } catch (e) {
        console.error(e);
        alert('Lỗi kết nối khi đồng bộ từ Blue Team');
      }
    }

    async function openInExplorer() {
      if (!currentTargetDir) {
        alert('Chưa chọn thư mục mục tiêu!');
        return;
      }
      try {
        const res = await fetch('/api/workspace?action=open');
        const data = await res.json();
        if (data.status === 'ok') {
          showToast('Đã mở cửa sổ File Explorer tại thư mục mục tiêu.', true);
        } else {
          alert('Lỗi mở thư mục: ' + data.message);
        }
      } catch (e) {
        console.error(e);
      }
    }

    async function fetchStatus(forceUpdateFiles = false) {
      try {
        const res = await fetch('/api/status');
        const data = await res.json();
        currentTargetDir = data.target_dir || '';
        
        const targetInput = document.getElementById('targetDirInput');
        if (targetInput && document.activeElement !== targetInput) {
          targetInput.value = currentTargetDir;
        }
        const clearBtn = document.getElementById('btnClearFolder');
        if (clearBtn) {
          clearBtn.style.display = currentTargetDir ? 'inline-block' : 'none';
        }

        const badgeFiles = document.getElementById('badgeFileCount');
        if (badgeFiles) {
          badgeFiles.innerText = `${data.target_file_count !== undefined ? data.target_file_count : (data.target_files ? data.target_files.length : 0)} tệp tin`;
        }

        document.getElementById('badgeTampered').innerText = `Sửa đổi: ${data.tampered_count}`;
        document.getElementById('badgeInjected').innerText = `Xâm nhập: ${data.injected_count}`;
        document.getElementById('badgeDeleted').innerText = `Đã xóa: ${data.deleted_count}`;

        // Populate selects when directory changed or forced
        if (forceUpdateFiles || lastRenderedDir !== currentTargetDir) {
          lastRenderedDir = currentTargetDir;
          const sT = document.getElementById('selectTamperFile');
          const sD = document.getElementById('selectDeleteFile');
          if (sT && sD) {
            sT.innerHTML = '<option value="">Tự động chọn tệp mẫu...</option>';
            sD.innerHTML = '<option value="">Tự động chọn tệp cuối...</option>';
            if (data.target_files && data.target_files.length) {
              data.target_files.slice(0, 100).forEach(f => {
                const opt = document.createElement('option');
                opt.value = f; opt.innerText = f;
                sT.appendChild(opt.cloneNode(true));
                sD.appendChild(opt);
              });
            }
          }
        }

        renderLogs(data.latest_logs);
      } catch(e) {
        console.error(e);
      }
    }

    function renderLogs(logs) {
      if (!logs) return;
      const box = document.getElementById('consoleBox');
      let html = '';
      logs.forEach(l => {
        const isSuccess = l.status.includes('THÀNH CÔNG') || l.status.includes('BẺ GÃY') || l.status.includes('HOÀN TẤT');
        const statusClass = isSuccess ? 'log-success' : 'log-fail';
        html += `<div class="log-line">
          <span class="log-time">${l.timestamp}</span>
          <span style="color:#0f172a; font-weight:700;">${l.attack}:</span>
          <span class="${statusClass}">${l.status}</span>
          <pre class="log-details">${JSON.stringify(l.details, null, 2)}</pre>
        </div>`;
      });
      box.innerHTML = html;
    }

    function showToast(msg, isSuccess = true) {
      let t = document.getElementById('toastNotification');
      if (!t) {
        t = document.createElement('div');
        t.id = 'toastNotification';
        t.style.position = 'fixed';
        t.style.top = '20px';
        t.style.right = '20px';
        t.style.padding = '12px 20px';
        t.style.borderRadius = '6px';
        t.style.fontWeight = '700';
        t.style.fontSize = '13px';
        t.style.boxShadow = '0 4px 12px rgba(0,0,0,0.15)';
        t.style.zIndex = '9999';
        t.style.transition = 'all 0.3s ease';
        document.body.appendChild(t);
      }
      t.style.background = isSuccess ? '#ecfdf5' : '#fef2f2';
      t.style.color = isSuccess ? '#059669' : '#b91c1c';
      t.style.border = isSuccess ? '1px solid #a7f3d0' : '1px solid #fecaca';
      t.innerText = msg;
      t.style.display = 'block';
      setTimeout(() => { t.style.display = 'none'; }, 4000);
    }

    async function sendAttack(url, body) {
      const res = await fetch(url, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(body || {})
      });
      const data = await res.json();
      fetchStatus();
      showToast(data.status || 'Đã phát động đòn tấn công thành công. Chi tiết hiển thị ở Nhật ký bên dưới.', true);
    }

    function attackTamper() {
      const f = document.getElementById('selectTamperFile').value;
      sendAttack('/api/attack/tamper', { file: f });
    }

    function attackInject() {
      const name = document.getElementById('inputBackdoorName').value;
      sendAttack('/api/attack/inject', { filename: name });
    }

    function attackDelete() {
      const f = document.getElementById('selectDeleteFile').value;
      sendAttack('/api/attack/delete', { file: f });
    }

    function attackCrc32() {
      const msg = document.getElementById('inputCrc32Msg').value;
      sendAttack('/api/attack/crc32', { message: msg });
    }

    function attackDolevYao() {
      sendAttack('/api/attack/dolev_yao', {});
    }

    function attackSecondPreimage() {
      sendAttack('/api/attack/second_preimage', {});
    }

    function attackTiming() {
      sendAttack('/api/attack/timing', {});
    }

    function attackMetadata() {
      sendAttack('/api/attack/metadata', {});
    }

    async function launchAll7Attacks() {
      const btn = document.getElementById('btnLaunchAll7');
      if (btn) { btn.disabled = true; btn.innerText = 'Đang phát động 7 đòn...'; }
      try {
        showToast('Đang phát động đồng loạt 7 đòn tấn công đối kháng lên đĩa...', true);
        const res = await fetch('/api/attack/launch_all_7', { method: 'POST' });
        const data = await res.json();
        fetchStatus();
        showToast('Đã phát động thành công toàn diện 7 đòn tấn công. Hãy sang trang Blue Team để thực thi kiểm định!', true);
      } catch (e) {
        console.error(e);
        showToast('Lỗi phát động tấn công', false);
      } finally {
        if (btn) { btn.disabled = false; btn.innerText = 'Phát động đồng loạt 7 đòn tấn công'; }
      }
    }

    async function executeRollback() {
      const res = await fetch('/api/rollback', { method: 'POST' });
      const data = await res.json();
      fetchStatus();
      showToast(`Hoàn nguyên thành công! Đã phục hồi: ${data.restored_tampered_files} tệp sửa, ${data.cleaned_injected_files} tệp lạ, ${data.restored_deleted_files} tệp xóa.`, true);
    }

    fetchStatus(true);
    setInterval(fetchStatus, 3000);
  </script>
</body>
</html>
"""


class RedTeamRequestHandler(BaseHTTPRequestHandler):
    def send_json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        url = urlparse(self.path)
        path = url.path
        qs = parse_qs(url.query)
        if path in ("/", "/index.html"):
            body = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif path == "/api/status":
            st = ENGINE.get_status()
            st["target_files"] = ENGINE.get_target_files()
            self.send_json(st)
        elif path == "/api/workspace":
            action = qs.get("action", ["get"])[0]
            if action == "set":
                p = qs.get("path", [""])[0]
                res = ENGINE.set_target_dir(p)
                self.send_json(res)
            elif action == "clear":
                res = ENGINE.clear_target_dir()
                self.send_json(res)
            elif action == "sync":
                res = ENGINE.sync_target_from_blue_team()
                self.send_json(res)
            elif action == "open":
                res = ENGINE.open_target_in_explorer()
                self.send_json(res)
            else:
                self.send_json(ENGINE.get_status())
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        url = urlparse(self.path)
        path = url.path
        length = int(self.headers.get("Content-Length", 0))
        post_data = self.rfile.read(length) if length > 0 else b"{}"
        try:
            payload = json.loads(post_data.decode("utf-8")) if post_data else {}
        except Exception:
            payload = {}

        if path in ("/api/workspace", "/api/target_dir"):
            action = payload.get("action", "set")
            p = payload.get("path", "")
            if action == "clear":
                res = ENGINE.clear_target_dir()
            elif action == "sync":
                res = ENGINE.sync_target_from_blue_team()
            elif action == "open":
                res = ENGINE.open_target_in_explorer()
            else:
                res = ENGINE.set_target_dir(p)
            self.send_json(res)
        elif path == "/api/attack/tamper":
            res = ENGINE.attack_tamper_file(payload.get("file"))
            self.send_json(res)
        elif path == "/api/attack/inject":
            res = ENGINE.attack_inject_backdoor(payload.get("filename"))
            self.send_json(res)
        elif path == "/api/attack/delete":
            res = ENGINE.attack_delete_file(payload.get("file"))
            self.send_json(res)
        elif path == "/api/attack/crc32":
            res = ENGINE.attack_crc32_forge(payload.get("message"))
            self.send_json(res)
        elif path == "/api/attack/dolev_yao":
            res = ENGINE.attack_dolev_yao_swap()
            self.send_json(res)
        elif path == "/api/attack/second_preimage":
            res = ENGINE.attack_second_preimage_exploit()
            self.send_json(res)
        elif path == "/api/attack/timing":
            res = ENGINE.attack_timing_side_channel()
            self.send_json(res)
        elif path == "/api/attack/metadata":
            res = ENGINE.attack_metadata_recon()
            self.send_json(res)
        elif path == "/api/attack/launch_all_7":
            res = ENGINE.attack_launch_all_7()
            self.send_json(res)
        elif path == "/api/rollback":
            res = ENGINE.rollback_all()
            self.send_json(res)
        else:
            self.send_response(404)
            self.end_headers()


def run_red_team_server(port: int = 8081):
    server = ThreadingHTTPServer(("0.0.0.0", port), RedTeamRequestHandler)
    print("=" * 80)
    print(f"[*] RED TEAM ADVERSARY WEB DASHBOARD IS RUNNING ON: http://127.0.0.1:{port}")
    print(f"[*] Target Directory (Blue Team): {ENGINE.target_dir}")
    print("=" * 80)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[*] Đang dừng Red Team Web Server...")
        server.server_close()


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Red Team Adversary Web Server")
    parser.add_argument("--port", type=int, default=8081, help="Port to listen (default: 8081)")
    args = parser.parse_args()
    run_red_team_server(args.port)
