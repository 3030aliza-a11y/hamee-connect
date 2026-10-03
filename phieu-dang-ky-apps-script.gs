/**
 * HAMEE – PHIẾU ĐĂNG KÝ GIA NHẬP → GOOGLE SHEET (miễn phí)
 * Mỗi phiếu khách gửi trên webApp sẽ thành 1 dòng trong tab "PhieuDangKy" + gửi email báo Ban thư ký.
 *
 * CÀI ĐẶT (1 lần, khoảng 5 phút):
 * 1. Mở thư mục Drive HAMEE → Mới → Google Trang tính → đặt tên "HAMEE – Phiếu đăng ký hội viên".
 * 2. Trong Sheet: Tiện ích mở rộng → Apps Script → xoá code mẫu → dán toàn bộ file này → Lưu.
 * 3. Chọn hàm "taoBang" → Chạy → cấp quyền. Rồi chọn hàm "capQuyen" → Chạy → cấp quyền (Gmail, Drive).
 * 4. Triển khai → Triển khai mới → Loại: Ứng dụng web
 *      Thực thi với tư cách: Tôi  ·  Người có quyền truy cập: Bất kỳ ai → Triển khai.
 * 5. Copy link Web App (…/exec) gửi cho người phụ trách webApp để gắn vào form.
 *
 * CÁCH 2 (nếu Sheet không có menu Tiện ích mở rộng → Apps Script):
 *   a. Mở Google Sheet trên máy tính, copy ID trong link: docs.google.com/spreadsheets/d/<<ID>>/edit
 *   b. Vào https://script.google.com → Dự án mới → dán file này → điền ID vào SHEET_ID bên dưới → Lưu.
 *   c. Làm tiếp bước 3–5 ở trên (chạy taoBang, Triển khai ứng dụng web).
 *
 * XỬ LÝ: cột "Trạng thái" mặc định "Chờ duyệt". Ban thư ký đổi thành "Đã kết nạp" / "Từ chối" / "Đang xác minh"
 * (có danh sách chọn sẵn). Cột "Ghi chú xử lý" để ghi lại trao đổi.
 */
const SHEET_ID = "1T5LVQxx6tpOQL4rBlouBGCnQkrR8ERWyWk2pSdKZCKA"; // Sheet "HAMEE – Phiếu đăng ký hội viên"
const ss_ = () => SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
const EMAIL_HAMEE = "hoicokhidien@gmail.com";
const TAB = "PhieuDangKy";
const COLS = [
  ["thoi_gian", "Thời gian gửi"], ["trang_thai", "Trạng thái"], ["ghi_chu_xu_ly", "Ghi chú xử lý"],
  ["name", "Tên doanh nghiệp / Tổ chức"], ["addr", "Địa chỉ"], ["license", "Giấy phép ĐKKD số"], ["founded", "Ngày thành lập"],
  ["cphone", "Điện thoại DN"], ["cemail", "Email DN"], ["web", "Website"], ["business", "Ngành nghề SXKD"], ["sectors", "Nhóm ngành HAMEE"],
  ["capital", "Vốn đăng ký KD"], ["rev", "Doanh số năm"], ["staff", "Số lao động"], ["field", "Lĩnh vực hoạt động"], ["owner", "Hình thức sở hữu"],
  ["rep", "Người đại diện"], ["gender", "Giới tính"], ["dob", "Ngày sinh"], ["hometown", "Quê quán"],
  ["cccd", "Số căn cước"], ["cccdDate", "Ngày cấp"], ["cccdPlace", "Nơi cấp"], ["title", "Chức danh"], ["phone", "ĐTDĐ"], ["email", "Email người đại diện"],
  ["asName", "Trợ lý / thư ký"], ["asTitle", "Chức danh trợ lý"], ["asPhone", "ĐT trợ lý"],
  ["refName", "Người giới thiệu"], ["refOrg", "Đơn vị người giới thiệu"], ["file", "Tệp GPKD (tên tệp)"], ["fileUrl", "Link tệp GPKD"]
];

function taoBang() {
  const ss = ss_();
  const sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB, 0);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS.map(c => c[1])]).setFontWeight("bold").setBackground("#1565d8").setFontColor("#ffffff").setWrap(true);
  sh.setFrozenRows(1); sh.setFrozenColumns(4);
  sh.setColumnWidths(1, COLS.length, 150); sh.setColumnWidth(4, 260);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(["Chờ duyệt", "Đang xác minh", "Đã kết nạp", "Từ chối"], true).build();
  sh.getRange(2, 2, 1000, 1).setDataValidation(rule);
  const s1 = ss.getSheetByName("Sheet1") || ss.getSheetByName("Trang tính1"); if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action !== "dang_ky_hoi_vien") return out({ ok: false });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = ss_();
    if (!ss.getSheetByName(TAB)) taoBang();
    const sh = ss.getSheetByName(TAB);
    // Lưu tệp GPKD (nếu có) vào thư mục chứa Sheet
    if (d.fileData && d.file) {
      try {
        const parent = DriveApp.getFileById(ss.getId()).getParents();
        const folder = parent.hasNext() ? parent.next() : DriveApp.getRootFolder();
        const sub = folder.getFoldersByName("GPKD hội viên"); const f = sub.hasNext() ? sub.next() : folder.createFolder("GPKD hội viên");
        const blob = Utilities.newBlob(Utilities.base64Decode(d.fileData), d.fileType || "application/octet-stream", (d.name || "DN") + " - " + d.file);
        d.fileUrl = f.createFile(blob).getUrl();
      } catch (err) { d.fileUrl = "Lỗi lưu tệp: " + err; }
    }
    d.thoi_gian = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");
    d.trang_thai = "Chờ duyệt";
    const row = COLS.map(([k]) => { const v = Array.isArray(d[k]) ? d[k].join("; ") : (d[k] ?? ""); return /^0\d+/.test(String(v)) ? "'" + v : v; });
    sh.appendRow(row);
    MailApp.sendEmail({
      to: EMAIL_HAMEE,
      subject: `[HAMEE] Phiếu đăng ký hội viên mới: ${d.name}`,
      body: `Có phiếu đăng ký gia nhập mới (${d.thoi_gian})\n\nDoanh nghiệp: ${d.name}\nNgười đại diện: ${d.rep} – ${d.title}\nĐTDĐ: ${d.phone}\nEmail: ${d.email}\nNhóm ngành: ${(d.sectors || []).join("; ")}\n\nXem đầy đủ: ${ss.getUrl()}`
    });
  } finally { lock.releaseLock(); }
  return out({ ok: true });
}
// ---- OTP đăng nhập hội viên (qua email, miễn phí) ----
function doGet(e) {
  const p = (e && e.parameter) || {};
  const email = String(p.email || "").trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  if (p.action === "otp_send") {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return out({ ok: false, error: "email" });
    const n = +(cache.get("cnt_" + email) || 0);
    if (n >= 5) return out({ ok: false, error: "qua_nhieu" });          // tối đa 5 mã / giờ / email
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache.put("otp_" + email, code, 600);                                 // hiệu lực 10 phút
    cache.put("cnt_" + email, String(n + 1), 3600);
    MailApp.sendEmail({
      to: email,
      subject: `Mã đăng nhập HAMEE: ${code}`,
      body: `Mã đăng nhập tài khoản hội viên HAMEE của anh/chị là: ${code}\n\nMã có hiệu lực trong 10 phút. Không chia sẻ mã này cho người khác.\nNếu anh/chị không yêu cầu đăng nhập, vui lòng bỏ qua email này.\n\nHiệp hội Doanh nghiệp Cơ khí – Điện TP.HCM (HAMEE) · 028 3973 4081`
    });
    return out({ ok: true });
  }
  if (p.action === "otp_verify") {
    const code = cache.get("otp_" + email);
    const tries = +(cache.get("try_" + email) || 0);
    if (tries >= 5) return out({ ok: false, error: "qua_nhieu" });
    if (code && code === String(p.code || "").trim()) { cache.remove("otp_" + email); cache.remove("try_" + email); return out({ ok: true }); }
    cache.put("try_" + email, String(tries + 1), 600);
    return out({ ok: false, error: "sai_ma" });
  }
  return out({ ok: true, service: "HAMEE phieu dang ky" });
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Chạy 1 lần trong trình soạn thảo để cấp đủ quyền (Sheet, Gmail, Drive) trước khi Triển khai
function capQuyen() {
  MailApp.getRemainingDailyQuota();
  DriveApp.getRootFolder();
  ss_().getName();
}
