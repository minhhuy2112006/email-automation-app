//////////////////////
// Helper functions //
//////////////////////

function getSettings() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Setting Overal");

  const values = sheet.getRange(2, 1, 1, 11).getValues()[0]; // 11 cột tổng

  const useInline = (values[4] || "").toString().trim().toLowerCase() === "yes";
  const imageFolderId = useInline ? (values[8] || "") : "";

  const useAttachment = (values[9] || "").toString().trim().toLowerCase() === "yes";
  const attachmentFolderId = useAttachment ? (values[10] || "") : "";

  return {
    subject: values[0],
    send_time: values[1],
    template_male: values[2],
    template_female: values[3],
    use_inline_image: useInline,
    batch_size: Number(values[5]) || 10,
    min_delay_seconds: Number(values[6]) || 20,
    max_delay_seconds: Number(values[7]) || 45,
    image_folder_id: imageFolderId,
    use_attachment: useAttachment,
    attachment_folder_id: attachmentFolderId
  };
}



function loadHtmlFromDrive(fileId) {
  if (!fileId) throw new Error("Template ID is missing!");
  try {
    return DriveApp.getFileById(fileId).getBlob().getDataAsString("UTF-8");
  } catch (err) {
    throw new Error("Cannot access template file. Check ID and permissions.");
  }
}

// FIX GOOGLE DOCS BULLET (robust, no double-bullet)
function fixGoogleDocsBullet(html) {
  if (!html) return html;

  // 1) Remove <style> blocks (Docs style is unsafe for Gmail)
  html = html.replace(/<style[\s\S]*?<\/style>/gi, "");

  // 2) Normalize ul/ol tags and add inline styles so Gmail shows bullets/numbers
  html = html.replace(/<ul[^>]*>/gi, '<ul style="list-style-type: disc; margin-left:16px; padding-left:18px;">');
  html = html.replace(/<ol[^>]*>/gi, '<ol style="list-style-type: decimal; margin-left:16px; padding-left:18px;">');

  // 3) Remove attributes from li tags
  html = html.replace(/<li[^>]*>/gi, "<li>");

  // 4) Clean leading "bullet artifacts" inside <li>
  //    - remove leading spans that contain bullet chars (•, ·, ●, &bull;, &#8226;)
  //    - remove leading bullet characters or hyphens and non-breaking spaces
  //    - keep the rest of content intact (including inline tags)
  html = html.replace(/<li>([\s\S]*?)<\/li>/gi, (match, inner) => {
    let content = inner || "";
    // trim leading whitespace/newlines
    content = content.replace(/^[\s\u00A0]+/u, "");

    // remove leading <span ...>•</span> or similar (one or more)
    content = content.replace(/^(?:<span[^>]*>[\s\u00A0]*(?:&bull;|&#8226;|•|·|●|○|\u2022|\-)[\s\u00A0]*<\/span>)+/i, "");
    // remove any leading bullet characters/entities now
    content = content.replace(/^(?:&bull;|&#8226;|•|·|●|○|\u2022|\-|\u00B7|\u2023)[\s\u00A0]*/i, "");
    // also remove any leftover leading &nbsp; or normal spaces
    content = content.replace(/^[\s\u00A0]+/u, "");

    // If content starts with a <p>...</p>, unwrap single <p> so lists render cleaner
    if (/^<p[^>]*>[\s\S]*<\/p>$/.test(content.trim())) {
      content = content.replace(/^<p[^>]*>/i, "").replace(/<\/p>$/i, "");
    }

    // Return cleaned li (no manual bullet added; ul style provides bullets)
    return `<li>${content}</li>`;
  });

  return html;
}

function getAttachmentFile(folderId) {
  if (!folderId) return null;
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    if (files.hasNext()) return files.next().getBlob();
    return null;
  } catch (err) {
    throw new Error("Cannot access attachment folder. Check ID and permissions.");
  }
}



function parseSendTime(str) {
  if (!str) return null;
  if (str instanceof Date) return str;

  const formatted = str.toString().trim().replace(" ", "T");
  return new Date(formatted);
}

function setStatusColor(sheet, row, col, status) {
  if (status === "Successful") {
    sheet.getRange(row, col).setBackground("#b7e1cd");
  } else if (status === "Failed") {
    sheet.getRange(row, col).setBackground("#f4c7c3");
  } else {
    sheet.getRange(row, col).setBackground(null);
  }
}

function validateColumnIndex(idx, header) {
  for (const key in idx) {
    if (idx[key] < 0) {
      SpreadsheetApp.getUi().alert(`Thiếu cột: ${key}`);
      throw new Error(`Không tìm thấy cột "${key}" trong header: [${header.join(", ")}]`);
    }
  }
}

// Hàm lấy ảnh theo Name
function getRecipientImageByName(name, folderId) {
  if (!folderId) throw new Error("Image folder ID is missing!");
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(name + ".png");
    if (files.hasNext()) return files.next().getBlob();
    return null;
  } catch (err) {
    throw new Error("Cannot access image folder. Check ID and permissions.");
  }
}

//////////////////////
// Main Logic       //
//////////////////////

function sendScheduledBatch() {
  const ss = SpreadsheetApp.getActive();
  const recipientsSheet = ss.getSheetByName("Recipients");
  const logSheet = ss.getSheetByName("log");
  const archiveSheet = ss.getSheetByName("Archive");
  const settings = getSettings();

  const sendTime = parseSendTime(settings.send_time);
  if (!sendTime) return;
  const now = new Date();
  if (now < sendTime) return;

  let htmlMale = fixGoogleDocsBullet(loadHtmlFromDrive(settings.template_male));
  let htmlFemale = fixGoogleDocsBullet(loadHtmlFromDrive(settings.template_female));

  const data = recipientsSheet.getDataRange().getValues();
  const header = data[0].map(h => h.toString().trim().toLowerCase());
  const rows = data.slice(1);

  const idx = {
    email: header.indexOf("email"),
    name: header.indexOf("name"),
    gender: header.indexOf("gender"),
    status: header.indexOf("status"),
    academic_year: header.indexOf("academic_year")
  };
  validateColumnIndex(idx, header);

  // Reset temporary "Ready" statuses to empty so they can be sent
  for (let i = 0; i < rows.length; i++) {
    if ((rows[i][idx.status] || "").toString().trim().toLowerCase() === "ready") {
      recipientsSheet.getRange(i + 2, idx.status + 1).setValue(""); // reset Ready -> empty
      rows[i][idx.status] = ""; // cập nhật mảng local luôn
    }
  }
  
  const pending = rows
    .map((r, i) => ({ r, row: i + 2 }))
    .filter(e => !e.r[idx.status] || e.r[idx.status].toString().trim().toLowerCase() === "pending");
  if (pending.length === 0) return;

  const batch = pending.slice(0, settings.batch_size);

  // Lấy attachment chung nếu có
  let sharedAttachment = null;
  if (settings.use_attachment && settings.attachment_folder_id) {
    sharedAttachment = getAttachmentFile(settings.attachment_folder_id);
    if (!sharedAttachment) Logger.log("⚠️ Không tìm thấy file đính kèm trong folder " + settings.attachment_folder_id);
  }

  for (let i = 0; i < batch.length; i++) {
    const entry = batch[i];
    const row = entry.r;
    const rowIndex = entry.row;

    let status = "Successful";
    let errorMessage = "";

    const email = row[idx.email];
    const name = row[idx.name];
    const gender = (row[idx.gender] || "").toString().toLowerCase();
    const academicYear = row[idx.academic_year] || "";

    const template = (gender.startsWith("nữ") || gender === "nu") ? htmlFemale : htmlMale;
    let html = template.replace(/\{0\}/g, name);
    let options = { htmlBody: html };

    // Inline image
    if (settings.use_inline_image && settings.image_folder_id) {
      const imgBlob = getRecipientImageByName(name, settings.image_folder_id);
      if (!imgBlob) {
        status = "Failed";
        errorMessage = "Image not found for: " + name;
      } else {
        options.inlineImages = { "image": imgBlob };
        html += '<br><img src="cid:image" style="max-width:300px;">';
        options.htmlBody = html;
      }
    }

    // Gắn attachment chung
    if (sharedAttachment && status === "Successful") {
      options.attachments = [sharedAttachment];
    }

    // Thử gửi email
    if (status === "Successful") {
      try {
        GmailApp.sendEmail(email, settings.subject, "", options);
      } catch (err) {
        if (err.toString().toLowerCase().includes("service invoked too many times")) {
          // Quota exceeded → tạo draft
          GmailApp.createDraft(email, settings.subject, "", options);
          status = "Draft";
          errorMessage = "Quota exceeded, saved as draft";
        } else {
          status = "Failed";
          errorMessage = err.toString();
        }
      }
    }

    // Cập nhật status + màu
    recipientsSheet.getRange(rowIndex, idx.status + 1).setValue(status);
    if (status === "Draft") {
      recipientsSheet.getRange(rowIndex, idx.status + 1).setBackground("#fff2cc"); // vàng
    } else {
      setStatusColor(recipientsSheet, rowIndex, idx.status + 1, status);
    }

    // Log
    logSheet.appendRow([new Date(), name, email, status, errorMessage]);

    // Archive (có Academic Year)
    archiveSheet.appendRow([
      new Date(),
      name,
      academicYear,
      row[idx.gender],
      email,
      status,
      errorMessage
    ]);

    // Delay
    if (i < batch.length - 1) {
      let min = settings.min_delay_seconds;
      let max = settings.max_delay_seconds;
      if (max > 15) max = 15;
      if (min > max) min = max;
      const wait = Math.floor(Math.random() * (max - min + 1)) + min;
      Utilities.sleep(wait * 1000);
    }
  }

  // Xóa các dòng đã gửi hoặc draft khỏi Recipients
  const last = recipientsSheet.getLastRow();
  if (last > 1) {
    const allRows = recipientsSheet.getRange(2, 1, last - 1, header.length).getValues();
    const keepRows = allRows.filter(r => {
      const s = (r[idx.status] || "").toString().trim().toLowerCase();
      return s === "" || s.toLowerCase() === "pending" || s.toLowerCase() === "failed";
    });
    recipientsSheet.deleteRows(2, last - 1);
    if (keepRows.length > 0) {
      recipientsSheet.getRange(2, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
    }
  }
}

function previewScheduledBatch() {
  const ss = SpreadsheetApp.getActive();
  const recipientsSheet = ss.getSheetByName("Recipients");
  const settings = getSettings();

  let htmlMale = fixGoogleDocsBullet(loadHtmlFromDrive(settings.template_male));
  let htmlFemale = fixGoogleDocsBullet(loadHtmlFromDrive(settings.template_female));

  const data = recipientsSheet.getDataRange().getValues();
  const header = data[0].map(h => h.toString().trim().toLowerCase());
  const rows = data.slice(1);

  const idx = {
    email: header.indexOf("email"),
    name: header.indexOf("name"),
    gender: header.indexOf("gender"),
    status: header.indexOf("status")
  };
  validateColumnIndex(idx, header);

  // --- Lấy file attachment chung nếu có ---
  let sharedAttachment = null;
  let sharedAttachmentName = "No attachment";
  if (settings.use_attachment && settings.attachment_folder_id) {
    const folder = DriveApp.getFolderById(settings.attachment_folder_id);
    const files = folder.getFiles();
    if (files.hasNext()) {
      sharedAttachment = files.next();
      sharedAttachmentName = sharedAttachment.getName();
    }
  }

  let previewHtml = '<h2>Email Preview</h2><hr>';

  rows.forEach((row, i) => {
    const email = row[idx.email];
    const name = row[idx.name];
    const gender = (row[idx.gender] || "").toString().toLowerCase();

    const template = (gender.startsWith("nữ") || gender === "nu") ? htmlFemale : htmlMale;
    let html = template.replace(/\{0\}/g, name);

    let status = "Ready"; // default
    let imgHtml = '';

    // Inline image
    if (settings.use_inline_image && settings.image_folder_id) {
      const imgBlob = getRecipientImageByName(name, settings.image_folder_id);
      if (imgBlob) {
        const base64 = Utilities.base64Encode(imgBlob.getBytes());
        imgHtml = `<br><img src="data:${imgBlob.getContentType()};base64,${base64}" style="max-width:300px;">`;
      } else {
        status = "Missing Image";
      }
    }

    // Cập nhật tạm status vào sheet
    recipientsSheet.getRange(i + 2, idx.status + 1).setValue(status);

    // --- Thêm thông tin attachment ---
    let attachmentHtml = settings.use_attachment && sharedAttachmentName
      ? `<b>Attachment:</b> ${sharedAttachmentName}<br>`
      : '';

    previewHtml += `<div style="margin-bottom:40px; padding:10px; border:1px solid #ccc;">
      <b>To:</b> ${email}<br>
      <b>Name:</b> ${name}<br>
      <b>Status:</b> ${status}<br>
      ${attachmentHtml}
      <b>Content:</b><br>${html}${imgHtml}
    </div>`;
  });

  const htmlOutput = HtmlService.createHtmlOutput(previewHtml)
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Email Preview");
}


// 
function cleanExportedHtml(html) {
  // Xóa thẻ tiêu đề đầu tiên (title của Google Docs)
  html = html.replace(/<p[^>]*class="title"[^>]*>[\s\S]*?<\/p>/i, "");

  return html.trim();
}


function exportDocToCleanHtml(docId, outFileName, folderId) {
  if (!docId) throw new Error("docId is empty");
  if (!folderId) throw new Error("folderId is empty");

  // Kiểm tra tồn tại
  try { DriveApp.getFileById(docId); }
  catch (e) { throw new Error("Doc not found: " + docId); }

  try { DriveApp.getFolderById(folderId); }
  catch (e) { throw new Error("Folder not found: " + folderId); }

  // --- EXPORT HTML ---
  const url = `https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=text/html`;
  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token }
  });

  let html = response.getContentText();

  // Clean tab title
  html = cleanExportedHtml(html);

  // --- CLEAN SUBJECT/TITLE ---
  html = cleanHtmlTitle(html);

  // Lưu file
  const blob = Utilities.newBlob(html, "text/html", outFileName);

  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  Logger.log("Clean HTML created: " + file.getId());
  return file.getId();
}


// CLEAN FUNCTION
function cleanHtmlTitle(html) {
  // Bỏ title
  html = html.replace(/<title[^>]*>[\s\S]*?<\/title>/gi, "");

  // Bỏ meta og:title
  html = html.replace(/<meta[^>]*og:title[^>]*>/gi, "");

  // Xóa <h1 class="title">
  html = html.replace(/<h1[^>]*class="[^"]*title[^"]*"[^>]*>[\s\S]*?<\/h1>/gi, "");

  // --- CLEAN DÒNG ĐẦU TIÊN CHỨA MALE/FEMALE ---
  html = removeMaleFemaleFirstLine(html);

  return html.trim();
}

function removeMaleFemaleFirstLine(html) {
  const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  if (!bodyMatch) return html; // không tìm thấy body

  let bodyContent = bodyMatch[1];

  // Tách thành từng dòng HTML logic
  let parts = bodyContent.split(/(?=<)/g); // tách theo tag mở

  if (parts.length > 0) {
    // Nếu dòng đầu chứa male/female → xoá
    if (/male|female/i.test(parts[0])) {
      parts.shift();
    }
  }

  const newBody = parts.join("");

  // Ghép HTML lại
  return html.replace(bodyMatch[0], `<body>${newBody}</body>`);
}


function exportTwoDocsToHtml() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Convert HTML");

  const maleId = sheet.getRange("A2").getDisplayValue().trim();
  const femaleId = sheet.getRange("B2").getDisplayValue().trim();
  const folderId = sheet.getRange("C2").getDisplayValue().trim();

  // Validate cơ bản
  if (!maleId) throw new Error("❌ ID Doc male trống");
  if (!femaleId) throw new Error("❌ ID Doc female trống");
  if (!folderId) throw new Error("❌ Folder ID trống");

  // Export 2 file
  exportDocToCleanHtml(maleId, "male.html", folderId);
  exportDocToCleanHtml(femaleId, "female.html", folderId);

  Logger.log("🎉 Done! Created male.html & female.html");
}
