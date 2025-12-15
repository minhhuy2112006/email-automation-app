//////////////////////
// Helper functions //
//////////////////////

function getSettings() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Setting Overal");

  const values = sheet.getRange(2, 1, 1, 12).getValues()[0]; // thêm 1 cột

  const useInline = (values[4] || "").toString().trim().toLowerCase() === "yes";
  const imageFolderId = useInline ? (values[8] || "") : "";

  const useAttachment = (values[9] || "").toString().trim().toLowerCase() === "yes";
  const attachmentFolderId = useAttachment ? (values[10] || "") : "";

  // --- TEMPLATE MODE ---
  let templateMode = (values[11] || "gender").toString().trim().toLowerCase();
  if (templateMode !== "single" && templateMode !== "gender") {
    templateMode = "gender"; // fallback an toàn
  }

  return {
    subject: values[0],
    send_time: values[1],
    template_male: values[2],     // dùng cho male hoặc single
    template_female: values[3],
    use_inline_image: useInline,
    batch_size: Number(values[5]) || 10,
    min_delay_seconds: Number(values[6]) || 20,
    max_delay_seconds: Number(values[7]) || 45,
    image_folder_id: imageFolderId,
    use_attachment: useAttachment,
    attachment_folder_id: attachmentFolderId,
    template_mode: templateMode
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

function loadTemplates(settings) {
  let htmlMale = null;
  let htmlFemale = null;
  let htmlSingle = null;

  if (settings.template_mode === "single") {
    if (!settings.template_male) {
      throw new Error("❌ SINGLE mode nhưng template_male trống");
    }
    htmlSingle = fixGoogleDocsBullet(
      loadHtmlFromDrive(settings.template_male)
    );
  } else {
    if (!settings.template_male && !settings.template_female) {
      throw new Error("❌ GENDER mode nhưng cả male & female template đều trống");
    }

    if (settings.template_male) {
      htmlMale = fixGoogleDocsBullet(
        loadHtmlFromDrive(settings.template_male)
      );
    }

    if (settings.template_female) {
      htmlFemale = fixGoogleDocsBullet(
        loadHtmlFromDrive(settings.template_female)
      );
    }
  }

  return { htmlMale, htmlFemale, htmlSingle };
}


function getTemplateForRecipient(settings, templates, gender) {
  if (settings.template_mode === "single") {
    if (!templates.htmlSingle) {
      throw new Error("Single mode but template missing");
    }
    return templates.htmlSingle;
  }

  const g = gender.toLowerCase();

  if ((g === "female" || g === "nữ" || g === "nu")) {
    if (!templates.htmlFemale) {
      throw new Error("Female template missing");
    }
    return templates.htmlFemale;
  }

  if ((g === "male" || g === "nam")) {
    if (!templates.htmlMale) {
      throw new Error("Male template missing");
    }
    return templates.htmlMale;
  }

  throw new Error("Invalid gender: " + gender);
}



function applyTemplate(html, data) {
  if (!html) return "";

  return html.replace(/\{(\w+)\}/g, (_, key) => {
    return data[key] !== undefined ? data[key] : `{${key}}`;
  });
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

function getAttachmentFileNames(folderId) {
  if (!folderId) return [];

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const names = [];

    while (files.hasNext()) {
      const f = files.next();
      names.push(f.getName());
    }
    return names;
  } catch (e) {
    return [];
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
  const templates = loadTemplates(settings);

  const sendTime = parseSendTime(settings.send_time);
  if (!sendTime) return;
  if (new Date() < sendTime) return;

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

  // Lấy các dòng có thể gửi
  const candidates = rows
    .map((r, i) => ({ r, row: i + 2 }))
    .filter(e => {
      const s = (e.r[idx.status] || "").toString().trim().toLowerCase();
      return s === "" || s === "pending" || s === "ready";
    });

  if (candidates.length === 0) return;

  const batch = candidates.slice(0, settings.batch_size);

  // Attachment chung
  let sharedAttachment = null;
  if (settings.use_attachment && settings.attachment_folder_id) {
    sharedAttachment = getAttachmentFile(settings.attachment_folder_id);
  }

  for (let i = 0; i < batch.length; i++) {
    const { r: row, row: rowIndex } = batch[i];

    const email = row[idx.email];
    const name = row[idx.name];
    const rawGender = (row[idx.gender] || "").toString().trim().toLowerCase();
    const academicYear = row[idx.academic_year] || "";

    let status = "Successful";
    let errorMessage = "";

    // =====================
    // VALIDATE GENDER + TEMPLATE
    // =====================
    if (settings.template_mode === "gender") {
      if (!["male", "female", "nam", "nữ", "nu"].includes(rawGender)) {
        status = "Failed";
        errorMessage = "Invalid gender value: " + rawGender;
      }

      if (status === "Successful") {
        const isFemale = rawGender === "female" || rawGender === "nữ" || rawGender === "nu";
        const isMale   = rawGender === "male" || rawGender === "nam";

        if (isFemale && !templates.htmlFemale) {
          status = "Failed";
          errorMessage = "Female gender but female template missing";
        }

        if (isMale && !templates.htmlMale) {
          status = "Failed";
          errorMessage = "Male gender but male template missing";
        }
      }
    }

    // ⛔ FAIL → LOG + ARCHIVE + CONTINUE
    if (status === "Failed") {
      recipientsSheet.getRange(rowIndex, idx.status + 1).setValue("Failed");
      setStatusColor(recipientsSheet, rowIndex, idx.status + 1, "Failed");

      logSheet.appendRow([new Date(), name, email, "Failed", errorMessage]);

      archiveSheet.appendRow([
        new Date(),
        name,
        academicYear,
        row[idx.gender],
        email,
        "Failed",
        errorMessage
      ]);

      continue; // ⛔ KHÔNG GỬI
    }

    const gender = rawGender;
    const template = getTemplateForRecipient(settings, templates, gender);

    let html = template.replace(/\{0\}/g, name);

    let options = { htmlBody: html };

    // Inline image
    if (settings.use_inline_image && settings.image_folder_id) {
      const imgBlob = getRecipientImageByName(name, settings.image_folder_id);
      if (!imgBlob) {
        status = "Failed";
        errorMessage = "Image not found";
      } else {
        options.inlineImages = { image: imgBlob };
        html += '<br><img src="cid:image" style="max-width:300px;">';
        options.htmlBody = html;
      }
    }

    if (sharedAttachment && status === "Successful") {
      options.attachments = [sharedAttachment];
    }

    if (status === "Successful") {
      try {
        GmailApp.sendEmail(email, settings.subject, "", options);
      } catch (err) {
        if (err.toString().toLowerCase().includes("service invoked too many times")) {
          GmailApp.createDraft(email, settings.subject, "", options);
          status = "Draft";
          errorMessage = "Quota exceeded";
        } else {
          status = "Failed";
          errorMessage = err.toString();
        }
      }
    }

    recipientsSheet.getRange(rowIndex, idx.status + 1).setValue(status);
    setStatusColor(recipientsSheet, rowIndex, idx.status + 1, status);

    logSheet.appendRow([new Date(), name, email, status, errorMessage]);

    archiveSheet.appendRow([
      new Date(),
      name,
      academicYear,
      row[idx.gender],
      email,
      status,
      errorMessage
    ]);

    if (i < batch.length - 1) {
      Utilities.sleep(
        Math.min(
          Math.max(settings.min_delay_seconds, 1),
          Math.min(settings.max_delay_seconds, 15)
        ) * 1000
      );
    }
  }
  // ================================
  // CLEAN SENT ROWS (SUCCESS + DRAFT)
  // ================================
  const lastRow = recipientsSheet.getLastRow();
  if (lastRow <= 1) return;

  const bodyRange = recipientsSheet.getRange(2, 1, lastRow - 1, header.length);
  const allRows = bodyRange.getValues();

  const keepRows = allRows.filter(r => {
    const s = (r[idx.status] || "").toString().trim().toLowerCase();
    return s === "" || s === "pending" || s === "failed";
  });

  // ⛔ Không cần làm gì nếu không có thay đổi
  if (keepRows.length === allRows.length) {
    return;
  }

  // 🔥 Xóa body cũ
  bodyRange.clearContent();

  // ✍️ Ghi lại nếu còn rows
  if (keepRows.length > 0) {
    recipientsSheet
      .getRange(2, 1, keepRows.length, keepRows[0].length)
      .setValues(keepRows);
  }
}

function previewScheduledBatch() {
  const ss = SpreadsheetApp.getActive();
  const recipientsSheet = ss.getSheetByName("Recipients");
  const settings = getSettings();
  const templates = loadTemplates(settings);

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

  // --- Attachment info ---
  let attachmentInfo = "";
  if (settings.use_attachment && settings.attachment_folder_id) {
    try {
      const folder = DriveApp.getFolderById(settings.attachment_folder_id);
      const files = folder.getFiles();
      const names = [];
      while (files.hasNext()) names.push(files.next().getName());
      if (names.length > 0) {
        attachmentInfo = `<div style="margin-top:4px; white-space: normal; word-break: break-word; line-height: 1.4;"><b>Attachment:</b> <span>${names.join(", ")}</span></div>`;
      }
    } catch (e) {
      attachmentInfo = `<div><b>Attachment:</b> (cannot read files)</div>`;
    }
  }

  let previewHtml = "<h2></h2><hr>";

  // --- chỉ preview top 10 ---
  const previewRows = rows.slice(0, 10);

  previewRows.forEach((row, i) => {
    const email = row[idx.email];
    const name = row[idx.name];
    const rawGender = (row[idx.gender] || "").toString().toLowerCase();

    const gender = ["male", "female", "nam", "nữ", "nu", "single"].includes(rawGender)
      ? rawGender
      : "default";

    const template = getTemplateForRecipient(settings, templates, gender);
    let html = template.replace(/\{0\}/g, name);

    // --- Inline image preview ---
    let imgHtml = "";
    if (settings.use_inline_image && settings.image_folder_id) {
      const imgBlob = getRecipientImageByName(name, settings.image_folder_id);
      if (imgBlob) {
        const base64 = Utilities.base64Encode(imgBlob.getBytes());
        imgHtml = `<br><img src="data:${imgBlob.getContentType()};base64,${base64}" style="max-width:300px;">`;
      }
    }

    // Set status Ready (tạm)
    recipientsSheet.getRange(i + 2, idx.status + 1).setValue("Ready");

    previewHtml += `<div style="border:1px solid #ccc; padding:10px; margin-bottom:30px;">
      <b>To:</b> ${email}<br>
      <b>Name:</b> ${name}<br>
      <b>Status:</b> Ready
      ${attachmentInfo}
      <br>
      ${html}
      ${imgHtml}
    </div>`;
  });

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(previewHtml).setWidth(900).setHeight(650),
    "Email Preview"
  );
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
  let mode = (sheet.getRange("D2").getDisplayValue() || "gender")
    .toLowerCase()
    .trim();

  if (mode !== "single" && mode !== "gender") {
    mode = "gender"; // fallback
  }

  if (!folderId) throw new Error("❌ Folder ID trống");

  if (mode === "single") {
    if (!maleId) throw new Error("❌ ID Doc template chung trống");
    exportDocToCleanHtml(maleId, "template.html", folderId);
  } else {
    if (!maleId || !femaleId) {
      throw new Error("❌ Thiếu ID Doc male hoặc female");
    }
    exportDocToCleanHtml(maleId, "male.html", folderId);
    exportDocToCleanHtml(femaleId, "female.html", folderId);
  }

  Logger.log("✅ Convert HTML DONE (" + mode + ")");
}


// Get field ID
function showGetIdDialog() {
  const html = HtmlService.createHtmlOutput(`
    <label>Paste Google Doc/Drive link:</label>
    <input type="text" id="link" style="width:300px;">
    <button onclick="getId()">Get ID</button>
    <p id="result"></p>
    <script>
      function getId() {
        const link = document.getElementById('link').value;
        const match = link.match(/\\/d\\/([a-zA-Z0-9-_]+)/);
        document.getElementById('result').innerText = match ? match[1] : 'Invalid link';
      }
    </script>
  `).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Get Google Doc ID');
}

