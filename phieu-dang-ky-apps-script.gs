/**
 * HAMEE – PHIẾU ĐĂNG KÝ GIA NHẬP → GOOGLE SHEET (miễn phí)
 * Mỗi phiếu khách gửi trên webApp sẽ thành 1 dòng trong tab "PhieuDangKy" + gửi email báo Ban thư ký.
 *
 * CÀI ĐẶT (1 lần, khoảng 5 phút):
 * 1. Mở thư mục Drive HAMEE → Mới → Google Trang tính → đặt tên "HAMEE – Phiếu đăng ký hội viên".
 * 2. Trong Sheet: Tiện ích mở rộng → Apps Script → xoá code mẫu → dán toàn bộ file này → Lưu.
 * 3. Chọn hàm "taoBang" ở thanh trên → bấm Chạy → cấp quyền. Tab PhieuDangKy được tạo sẵn tiêu đề cột.
 * 4. Triển khai → Triển khai mới → Loại: Ứng dụng web
 *      Thực thi với tư cách: Tôi  ·  Người có quyền truy cập: Bất kỳ ai → Triển khai.
 * 5. Copy link Web App (…/exec) gửi cho người phụ trách webApp để gắn vào form.
 *
 * XỬ LÝ: cột "Trạng thái" mặc định "Chờ duyệt". Ban thư ký đổi thành "Đã kết nạp" / "Từ chối" / "Đang xác minh"
 * (có danh sách chọn sẵn). Cột "Ghi chú xử lý" để ghi lại trao đổi.
 */
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
  const ss = SpreadsheetApp.getActive();
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
    const ss = SpreadsheetApp.getActive();
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
function doGet() { return out({ ok: true, service: "HAMEE phieu dang ky" }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
