function generateHealthReport(e) {

  // ====== 📊 GET DATA ======
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const row = sheet.getLastRow();
  const d = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const labCell = d[32];

  // ====== 🧾 BASIC ======
  const name = d[1] || "";
  const hn = d[2] || "";
  const date = d[3] 
    ? Utilities.formatDate(new Date(d[3]), "Asia/Bangkok", "dd/MM/yyyy") 
    : "";
  const company = d[4] || "";

  const dob = d[5] 
    ? Utilities.formatDate(new Date(d[5]), "Asia/Bangkok", "dd/MM/yyyy") 
    : "";
  const age = d[6] || "";
  const sex = d[7] || "";

  const physician_th = d[8] || "";
  const physician_en = d[9] || "";

  // ====== 📄 TEMPLATE ======
  const templateId = "YOUR_ID_HERE";
  const folderId = "YOUR_ID_HERE";

  const copy = DriveApp.getFileById(templateId).makeCopy(`HealthReport_${hn}`);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // ====== 🩺 VITAL ======
  const temp = d[10] || "";
  const pulse = d[11] || "";
  const bp = String(d[12] || "");
  const rr = d[13] || "";

  let bp_result = "-";
  if (bp.includes("/")) {
    const [sys, dia] = bp.split("/").map(Number);
    if (!isNaN(sys) && !isNaN(dia)) {
      bp_result = (sys >= 140 || dia >= 90) ? "สูง" : "ปกติ";
    }
  }

  // ====== 📏 BODY ======
  const weight = Number(d[14]) || 0;
  const height = Number(d[15]) || 0;

  let bmi = "";
  if (weight && height) {
    bmi = (weight / ((height / 100) ** 2)).toFixed(2);
  }

  // ====== 👁 HELPER ======
  function mapStatus(status) {
    return {
      normal: status === "ปกติ" ? "✔" : "",
      abnormal: status === "ผิดปกติ" ? "✔" : "",
      na: status === "ไม่ตรวจ" ? "✔" : ""
    };
  }

  function replaceVision(name, status) {
    const map = mapStatus(status);
    body.replaceText(`{{${name}_normal}}`, map.normal);
    body.replaceText(`{{${name}_abnormal}}`, map.abnormal);
    body.replaceText(`{{${name}_na}}`, map.na);
  }

  function replacePE(name, status) {
    const map = mapStatus(status);

    body.replaceText(`{{${name}_normal}}`, map.normal);
    body.replaceText(`{{${name}_abnormal}}`, map.abnormal);
    body.replaceText(`{{${name}_na}}`, map.na);

    let resultText = "-";
    if (status === "ปกติ") resultText = "ปกติ";
    else if (status === "ไม่ตรวจ") resultText = "ไม่ตรวจ";
    else if (status === "ผิดปกติ") resultText = "พบความผิดปกติ";

    body.replaceText(`{{${name}_result}}`, resultText);
  }

  function getFileIdsFromCell(cellValue) {
    if (!cellValue) return [];
    return cellValue.split(',').map(url => {
      const match = url.match(/[-\w]{25,}/);
      return match ? match[0] : null;
    }).filter(id => id);
  }

  // ====== 👁 APPLY ======
  replaceVision("right_eye", d[16]);
  replaceVision("left_eye", d[17]);

  const color = mapStatus(d[18]);
  body.replaceText('{{color_normal}}', color.normal);
  body.replaceText('{{color_abnormal}}', color.abnormal);
  body.replaceText('{{color_na}}', color.na);

  replacePE("eye", d[19]);
  replacePE("ears", d[20]);
  replacePE("throat", d[21]);
  replacePE("nose", d[22]);
  replacePE("lymph", d[23]);
  replacePE("thyroid", d[24]);
  replacePE("heart", d[25]);
  replacePE("lung", d[26]);
  replacePE("abdomen", d[27]);
  replacePE("extremities", d[28]);

  // ====== 🧪 TEST ======
  body.replaceText('{{xray_result}}', d[29] || "-");
  body.replaceText('{{ekg_result}}', d[30] || "-");
  body.replaceText('{{other}}', d[31] || "-");

  // ====== 🔁 TEXT ======
  body.replaceText('{{name}}', name);
  body.replaceText('{{hn}}', hn);
  body.replaceText('{{date}}', date);
  body.replaceText('{{company}}', company);
  body.replaceText('{{dob}}', dob);
  body.replaceText('{{age}}', age);
  body.replaceText('{{sex}}', sex);

  body.replaceText('{{physician_th}}', physician_th);
  body.replaceText('{{physician_en}}', physician_en);

  body.replaceText('{{temp}}', temp);
  body.replaceText('{{pulse}}', pulse);
  body.replaceText('{{bp}}', bp);
  body.replaceText('{{bp_result}}', bp_result);
  body.replaceText('{{rr}}', rr);

  body.replaceText('{{weight}}', weight.toString());
  body.replaceText('{{height}}', height.toString());
  body.replaceText('{{bmi}}', bmi);

// ====== 📎 จัดการ {{lab_section}} ให้ล่องหนแบบปลอดภัย 100% ======
  const cleanupLab = body.findText('{{lab_section}}'); 
  if (cleanupLab) {
    const textElement = cleanupLab.getElement();
    
    // 1. ลบข้อความออกให้กลายเป็นความว่างเปล่า
    textElement.setText(''); 
    
    const p = textElement.getParent();
    // 2. ปรับระยะเว้นบรรทัด (Spacing) ให้เป็น 0 เพื่อให้บรรทัดนี้แฟบลงจนมองไม่เห็น
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);
  }

  // ====== 📎 SHOW PREVIEW (MAXIMIZED FULL PAGE A4 - ONE IMAGE PER PAGE) ======
  if (labCell && labCell.toString().trim() !== "") {
    
    // ขึ้นหน้าใหม่ก่อนเริ่มแปะรูปภาพแรก
    body.appendPageBreak(); 

    const fileIds = getFileIdsFromCell(labCell);

    // 📐 ดึงขนาดหน้ากระดาษและขอบ (Margin)
    const pageWidth = body.getPageWidth();
    const pageHeight = body.getPageHeight();
    const marginL = body.getMarginLeft();
    const marginR = body.getMarginRight();
    const marginT = body.getMarginTop();
    const marginB = body.getMarginBottom();

    // ✨ แนวทางใหม่: เพื่อขยายภาพให้เต็มแผ่นที่สุดถึงขอบกระดาษซ้าย-ขวา
    // เราจะไม่ใช้ maxWidth ที่หักลบ Margin ซ้าย-ขวาออก แต่จะใช้ความกว้างหน้ากระดาษจริงเลย
    const maxWidth = pageWidth; 
    
    // เรายังต้องลบ Margin บน-ล่าง เผื่อไว้ให้เนื้อหาไม่ชนขอบบนสุดของกระดาษ Docs
    const maxHeight = pageHeight - marginT - marginB; 

    // ใช้ index เพื่อเช็คว่าถึงรูปสุดท้ายหรือยัง
    fileIds.forEach((id, index) => {
      try {
        const file = DriveApp.getFileById(id);
        const mime = file.getMimeType();
        let img;

        if (mime.startsWith("image/")) {
          img = body.appendImage(file.getBlob());
        } else if (mime === "application/pdf") {
          const thumb = file.getThumbnail();
          if (thumb) {
            img = body.appendImage(thumb);
          }
        }

// ====== 🎯 AUTO SCALE (Maximizes size to paper edge) ======
        if (img) {
          let origWidth = img.getWidth();
          let origHeight = img.getHeight();

          let ratioWidth = maxWidth / origWidth;
          let ratioHeight = maxHeight / origHeight;
          let targetRatio = Math.min(ratioWidth, ratioHeight);

          img.setWidth(origWidth * targetRatio);
          img.setHeight(origHeight * targetRatio);

          const imgP = img.getParent();
          imgP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          imgP.setSpacingBefore(0);
          imgP.setSpacingAfter(0);
          
          // ✨ ใส่ค่าติดลบ เพื่อดึงภาพให้ทะลุเส้น Margin ออกไปชนขอบกระดาษจริงๆ!
          imgP.setIndentStart(-marginL);
          imgP.setIndentEnd(-marginR);
          imgP.setIndentFirstLine(-marginL);
        }

        // ====== 📄 บังคับขึ้นหน้าใหม่หลังแปะรูปแต่ละรูป (ยกเว้นรูปสุดท้าย) ======
        if (index < fileIds.length - 1) {
          body.appendPageBreak();
        }

      } catch (e) {
        Logger.log("❌ Error processing file ID " + id + ": " + e);
      }
    });
  }

  // ====== 💾 SAVE ======
  doc.saveAndClose();

  const finalPdf = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
  DriveApp.getFolderById(folderId).createFile(finalPdf);

  // ลบไฟล์ Docs ต้นฉบับทิ้งหลังจากได้ PDF แล้ว
  DriveApp.getFileById(copy.getId()).setTrashed(true);
}
