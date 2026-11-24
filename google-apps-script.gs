/**
 * Nhận đăng ký từ website Ngọc Hương Academy và ghi vào Google Sheet.
 * Cách cài: mở Google Sheet → Tiện ích mở rộng → Apps Script → dán toàn bộ file này → Lưu
 *          → Triển khai → Tùy chọn triển khai mới → Ứng dụng web
 *          (Thực thi dưới dạng: Tôi · Người có quyền truy cập: Bất kỳ ai) → Triển khai.
 */

const SHEET_NAME = 'DangKy';
const NOTIFY_EMAIL = true; // true = gửi email báo cho chủ Sheet mỗi khi có đăng ký mới

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Bẫy chống spam: ô ẩn "website" chỉ bot mới điền
  if (p.website) return ok_();

  const clean = (v, max) => {
    let s = String(v || '').trim().slice(0, max || 500);
    if (/^[=+\-@]/.test(s)) s = "'" + s; // chống chèn công thức
    return s;
  };

  const row = [
    new Date(),
    clean(p.name, 120),
    "'" + clean(p.phone, 20).replace(/^'/, ''), // giữ số 0 đầu
    clean(p.email, 120),
    clean(p.course, 120),
    clean(p.note, 1000),
    clean(p.source, 60),
    'Mới'
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).appendRow(row);
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      'Đăng ký mới: ' + row[1] + ' – ' + row[4],
      'Họ tên: ' + row[1] + '\nSĐT: ' + row[2].replace(/^'/, '') + '\nEmail: ' + row[3] +
      '\nKhoá học: ' + row[4] + '\nLời nhắn: ' + row[5] + '\nNguồn: ' + row[6] +
      '\n\nXem danh sách: ' + SpreadsheetApp.getActive().getUrl()
    );
  }
  return ok_();
}

function doGet() {
  return ContentService.createTextOutput('Ngọc Hương Academy form endpoint đang hoạt động.');
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
