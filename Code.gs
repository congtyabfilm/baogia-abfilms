// ==========================================
// AB FILMS - TỰ ĐỘNG TẠO BÁO GIÁ PHIM BẢO VỆ
// ==========================================
var GEMINI_API_KEY = "AQ.Ab8RN6J2NP3SlHIYtmTF2VwGbNmmnbFIEIGthP1BO_vELB9vag";

function onOpen() {
  DocumentApp.getUi()
    .createMenu("🤖 AB Films AI")
    .addItem("Tạo Báo Giá Nhanh", "showSidebar")
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("AB Films - Trợ Lý Báo Giá")
    .setWidth(400);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Gọi AI Gemini với cơ chế tự động thử lại và đổi model dự phòng khi máy chủ Google quá tải (503)
 */
function callGeminiSmart(parts) {
  var models = ["gemini-3.5-flash-lite", "gemini-3.5-flash", "gemini-flash-latest"];
  var lastError = "";

  for (var m = 0; m < models.length; m++) {
    var modelName = models[m];
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + GEMINI_API_KEY;

    for (var attempt = 1; attempt <= 2; attempt++) {
      try {
        var response = UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({
            contents: [{ parts: parts }],
            generationConfig: {
              responseMimeType: "application/json"
            }
          }),
          muteHttpExceptions: true
        });

        var resCode = response.getResponseCode();
        var resText = response.getContentText();

        if (resCode === 200) {
          var resJson = JSON.parse(resText);
          var raw = resJson.candidates[0].content.parts[0].text;
          return JSON.parse(raw);
        } else if (resCode === 503 || resCode === 429) {
          Utilities.sleep(1500);
          lastError = "Lỗi " + resCode + " (" + modelName + "): " + resText;
        } else {
          lastError = "Lỗi " + resCode + " (" + modelName + "): " + resText;
          break;
        }
      } catch (err) {
        lastError = err.toString();
        Utilities.sleep(1000);
      }
    }
  }

  throw new Error("Không thể kết nối AI (vui lòng thử lại sau vài giây): " + lastError);
}

/**
 * Bóc tách nội dung bằng AI và tự động nhân bản thành file mới
 */
function processAndDuplicateDoc(payload) {
  try {
    var promptSystem = "Bạn là trợ lý AI chuyên bóc tách thông tin làm báo giá cho Công ty AB Films.\n" +
      "Hãy phân tích tin nhắn hoặc hình ảnh đầu vào và trả về JSON chuẩn xác:\n" +
      "- customerName: Tên khách hàng (ví dụ: 'Anh Nam', 'Chị Quỳnh', 'Chị Hạnh', nếu không có ghi 'Khách hàng')\n" +
      "- address: Địa chỉ công trình (ví dụ: 'Hà Đông', 'Bắc Ninh', '62 Vũ Trọng Khánh')\n" +
      "- hasDiscount: boolean (true nếu có nhắc đến chiết khấu / giảm giá / bớt %, false nếu không)\n" +
      "- discountPercent: phần trăm chiết khấu (kiểu số nguyên, ví dụ 5 nghĩa là 5%, 10 nghĩa là 10%)\n" +
      "- items: mảng các mã phim:\n" +
      "  + filmCode: Tên mã film (ví dụ: 'PPF', 'Pnc 50', 'Top 70', 'Silikante Pro', 'Silikante Luxury', 'Silihome'...)\n" +
      "  + area: diện tích m2 (kiểu số, ví dụ 5 hoặc 10 hoặc 25. LƯU Ý: Nếu yêu cầu nêu diện tích chung thì gán cho tất cả các mã)\n" +
      "  + originalUnitPrice: đơn giá gốc VNĐ/m2 (số nguyên đầy đủ: 900k -> 900000, 520k -> 520000, 700k -> 700000, '1tr4' -> 1400000, '2tr1' -> 2100000)\n" +
      "Chỉ trả về JSON thuần túy, không kèm giải thích.";

    var parts = [{ text: promptSystem }];

    if (payload.image) {
      parts.push({
        inlineData: {
          mimeType: payload.mimeType || "image/png",
          data: payload.image
        }
      });
      if (payload.text) {
        parts.push({ text: "Ghi chú thêm: " + payload.text });
      }
    } else {
      parts.push({ text: "Nội dung cần làm báo giá: \n" + payload.text });
    }

    // 1. Gọi AI Gemini
    var data = callGeminiSmart(parts);

    // 2. Ngày tháng hiện tại
    var now = new Date();
    data.date = {
      day: now.getDate(),
      month: now.getMonth() + 1,
      year: now.getFullYear()
    };

    // 3. TỰ ĐỘNG NHÂN BẢN VÀ ĐẶT TÊN THEO ĐỊNH DẠNG: "Tên khách - BG Phim bảo vệ nội thất"
    var templateDoc = DocumentApp.getActiveDocument();
    var templateFile = DriveApp.getFileById(templateDoc.getId());
    
    var custNameClean = (data.customerName || "Khách hàng").trim();
    var newDocTitle = custNameClean + " - BG Phim bảo vệ nội thất";

    var parents = templateFile.getParents();
    var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    var newFile = templateFile.makeCopy(newDocTitle, folder);
    var newDoc = DocumentApp.openById(newFile.getId());

    // 4. Áp dụng dữ liệu vào file mới (Font 14, KHÔNG IN ĐẬM nội dung điền)
    applyDataToDocument(newDoc, data);

    return {
      success: true,
      newDocUrl: newDoc.getUrl(),
      newDocTitle: newDocTitle,
      extracted: data
    };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Căn chỉnh từng ô trong bảng: Font 14 Times New Roman
 */
function setCellContent(cell, text, isBold, colWidth) {
  if (colWidth) {
    cell.setWidth(colWidth);
  }
  cell.setPaddingTop(4);
  cell.setPaddingBottom(4);
  cell.setPaddingLeft(2);
  cell.setPaddingRight(2);
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

  var str = (text !== undefined && text !== null) ? String(text) : "";
  var lines = str.split("\n");
  
  var p0 = cell.getChild(0).asParagraph();
  p0.setText(lines[0] || " ");
  p0.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p0.setFontFamily("Times New Roman");
  p0.setFontSize(14);
  p0.setBold(isBold === true);
  p0.setLineSpacing(1.0);
  p0.setSpacingBefore(0);
  p0.setSpacingAfter(0);

  for (var i = 1; i < lines.length; i++) {
    var pExtra = cell.appendParagraph(lines[i] || " ");
    pExtra.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pExtra.setFontFamily("Times New Roman");
    pExtra.setFontSize(14);
    pExtra.setBold(isBold === true);
    pExtra.setLineSpacing(1.0);
    pExtra.setSpacingBefore(0);
    pExtra.setSpacingAfter(0);
  }
}

function applyDataToDocument(doc, data) {
  var body = doc.getBody();

  // 1. Dòng Tên khách & Địa chỉ (Font 14 thường, KHÔNG IN ĐẬM)
  var introFound = body.findText("Lời đầu tiên xin thay mặt Công ty AB Films gửi tới Quý khách hàng:");
  if (introFound) {
    var p = introFound.getElement().getParent().asParagraph();
    var prefix = "Lời đầu tiên xin thay mặt Công ty AB Films gửi tới Quý khách hàng: ";
    var customerInfo = (data.customerName || "Quý khách") + (data.address ? " - Địa chỉ: " + data.address : "");
    var fullText = prefix + customerInfo;

    p.setText(fullText);
    p.setFontFamily("Times New Roman").setFontSize(14);
    p.editAsText().setBold(false); // Chữ thường, không in đậm!
  }

  // 2. Dòng ngày tháng cuối trang (Font 14 thường, KHÔNG IN ĐẬM)
  var dateStr = "Hà Nội, ngày " + data.date.day + " tháng " + data.date.month + " năm " + data.date.year;
  var dateFound = body.findText("Hà Nội, ngày.*tháng.*năm.*");
  if (dateFound) {
    var dateP = dateFound.getElement().getParent().asParagraph();
    dateP.setText(dateStr);
    dateP.setFontFamily("Times New Roman").setFontSize(14).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    dateP.editAsText().setBold(false);
  }

  // 3. Tìm và cập nhật Bảng Báo Giá
  var tables = body.getTables();
  var quoteTable = null;
  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];
    if (t.getNumRows() > 0 && t.getRow(0).getText().indexOf("Mã film") !== -1) {
      quoteTable = t;
      break;
    }
  }
  if (!quoteTable && tables.length > 0) {
    quoteTable = tables[tables.length - 1];
  }

  if (quoteTable) {
    var parent = quoteTable.getParent();
    var tableIndex = parent.getChildIndex(quoteTable);

    var hasDiscount = data.hasDiscount === true && (Number(data.discountPercent) > 0);
    var headers = [];
    var colWidths = [];

    if (hasDiscount) {
      var discStr = (data.discountPercent || "0") + "%";
      headers = [
        "Mã film dán",
        "Diện tích\n(M2)",
        "Đơn giá\n(VNĐ/M2)",
        "Đơn giá sau CK\n" + discStr + " (VNĐ/M2)",
        "Chi phí hoàn thiện\nchưa VAT (VNĐ)",
        "Chi phí hoàn thiện\nbao gồm VAT (VNĐ)"
      ];
      colWidths = [85, 48, 72, 95, 105, 105];
    } else {
      headers = [
        "Mã film dán",
        "Diện tích\n(M2)",
        "Đơn giá\n(VNĐ/M2)",
        "Chi phí hoàn thiện\nchưa VAT (VNĐ)",
        "Chi phí hoàn thiện\nbao gồm VAT (VNĐ)"
      ];
      colWidths = [110, 60, 90, 125, 125];
    }

    var newTable = parent.insertTable(tableIndex);

    // Tiêu đề cột (Font 14 Bold)
    var headerRow = newTable.appendTableRow();
    for (var h = 0; h < headers.length; h++) {
      var cell = headerRow.appendTableCell();
      setCellContent(cell, headers[h], true, colWidths[h]);
    }

    // Các dòng mã phim (Font 14 thường, KHÔNG IN ĐẬM)
    var items = data.items || [];
    for (var r = 0; r < items.length; r++) {
      var item = items[r];
      var area = Number(item.area) || 0;
      var origPrice = Number(item.originalUnitPrice) || 0;
      var discPercent = hasDiscount ? (Number(data.discountPercent) || 0) : 0;
      
      var effectivePrice = hasDiscount ? Math.round(origPrice * (1 - discPercent / 100)) : origPrice;
      var totalNoVAT = Math.round(area * effectivePrice);
      var totalWithVAT = Math.round(totalNoVAT * 1.08); // +8% VAT

      var row = newTable.appendTableRow();
      var rowValues = [];
      if (hasDiscount) {
        rowValues = [
          item.filmCode || "",
          area.toString(),
          formatMoney(origPrice),
          formatMoney(effectivePrice),
          formatMoney(totalNoVAT),
          formatMoney(totalWithVAT)
        ];
      } else {
        rowValues = [
          item.filmCode || "",
          area.toString(),
          formatMoney(origPrice),
          formatMoney(totalNoVAT),
          formatMoney(totalWithVAT)
        ];
      }

      for (var c = 0; c < rowValues.length; c++) {
        var cCell = row.appendTableCell();
        // false: KHÔNG IN ĐẬM NỘI DUNG ĐIỀN
        setCellContent(cCell, rowValues[c], false, colWidths[c]);
      }
    }

    newTable.setBorderColor("#000000");
    newTable.setBorderWidth(1);
    quoteTable.removeFromParent();
  }

  doc.saveAndClose();
}

function formatMoney(num) {
  if (isNaN(num)) return "0";
  return Math.round(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}
