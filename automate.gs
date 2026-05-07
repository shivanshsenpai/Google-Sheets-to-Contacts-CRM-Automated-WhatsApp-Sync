/**
 * UNIFIED WHATSAPP CRM & CONTACT SYNC
 * Flow: Paste in Customer -> Auto Clean/Categorize (OnEdit) -> Background Sync (Timer) OR Manual Sync
 */

// ==========================================
// 1. GLOBAL SETTINGS & CONFIGURATIONS
// ==========================================
const S_CUSTOMER = "customer";
const S_PROCESSED = "Processed_data";
const S_CONFIG = "Config";
const EXPORT_FOLDER_NAME = "WhatsApp_Exports"; // Sanitized folder name

// NOTE: CATEGORIES was missing in the original script. 
// Populate this array dynamically from your Config sheet or define it here.
const CATEGORIES = ["Category A", "Category B", "Others"]; 

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🟢 WhatsApp CRM')
    .addItem('📥 1. Export Selected to WhatsApp CSVs', 'showCsvPopup')
    .addItem('📱 2. Manual Sync to Google Contacts', 'showSyncPopup')
    .addSeparator()
    .addItem('⚙️ Initialize Automation & Background Sync', 'setupTriggers')
    .addToUi();
}

/**
 * INSTALLABLE TRIGGERS
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger('handleOnEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  ScriptApp.newTrigger('processBackgroundSync')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  SpreadsheetApp.getUi().alert("✅ Automation Active!\n\n1. Pasted data is instantly cleaned and categorized.\n2. Contacts will sync securely in the background every 5 minutes.");
}

// ==========================================
// 2. THE MAIN ENGINE (ON EDIT - BATCH PROCESSING)
// ==========================================

function handleOnEdit(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() !== S_CUSTOMER || range.getRow() === 1) return;
  
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  
  if (numCols < 5 && range.getColumn() > 1) return; 

  const allValues = sheet.getRange(startRow, 1, numRows, 5).getValues();
  const procSheet = ss.getSheetByName(S_PROCESSED) || ss.insertSheet(S_PROCESSED);
  const configSheet = ss.getSheetByName(S_CONFIG);
  
  if (!configSheet) return;

  const configData = configSheet.getDataRange().getValues();
  const pendingDataToAppend = [];

  for (let i = 0; i < allValues.length; i++) {
    const rowData = allValues[i];
    const [custName, city, orderTime, prodName, phone] = rowData;

    if (!phone || phone.toString().trim() === "") continue;
    if (phone.toString().trim() === "-" || custName === "-" || prodName === "-") continue;
    if (isDuplicatePhone(procSheet, phone)) continue;

    let assignedCategory = "Others";
    let categoryInitial = "OTH";
    let configRowIndex = -1;

    for (let j = 1; j < configData.length; j++) {
      const catName = configData[j][0].toString().trim();
      const initial = configData[j][1].toString().trim();
      const keywords = configData[j][2].toString().toLowerCase().split("|");
      
      const regexPattern = "\\b(" + keywords.map(k => k.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join("|") + ")\\b";
      const ruleRegex = new RegExp(regexPattern, "i");

      if (ruleRegex.test(prodName)) {
        if (catName.toLowerCase() === "exclude") {
          assignedCategory = "EXCLUDED";
          break;
        }
        assignedCategory = catName;
        categoryInitial = initial;
        configRowIndex = j + 1;
        break;
      }
    }

    if (assignedCategory === "EXCLUDED") continue;

    let groupNum = 1;
    let recCount = 0;
    
    if (configRowIndex !== -1) {
      groupNum = parseInt(configSheet.getRange(configRowIndex, 4).getValue()) || 1;
      recCount = parseInt(configSheet.getRange(configRowIndex, 5).getValue()) || 0;
      
      if (recCount >= 200) {
        groupNum++;
        recCount = 0;
      }
      recCount++;
      
      configSheet.getRange(configRowIndex, 4).setValue(groupNum);
      configSheet.getRange(configRowIndex, 5).setValue(recCount);
    }

    const cleanPhone = phone.toString().trim().replace(/^(\+?91)/, "");
    const formattedName = custName.toString().replace(/\s+/g, "_");
    const broadcastName = `${categoryInitial}_G${groupNum}_C${recCount}_${formattedName}`;

    pendingDataToAppend.push([custName, city, orderTime, prodName, phone, assignedCategory, broadcastName, cleanPhone, "Pending"]);
  }

  if (pendingDataToAppend.length > 0) {
    if (procSheet.getLastRow() === 0) {
      procSheet.appendRow(["Customer_Name", "City", "Order_Time", "Product_Name", "Phone_Number", "category", "Broadcast_Name", "WhatsApp_Number", "Synced_to_Contacts"]);
    }
    procSheet.getRange(procSheet.getLastRow() + 1, 1, pendingDataToAppend.length, 9).setValues(pendingDataToAppend);
  }
}

// ==========================================
// 3. BACKGROUND SYNC ENGINE (PREVENTS TIMEOUTS)
// ==========================================

function processBackgroundSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const procSheet = ss.getSheetByName(S_PROCESSED);
  if (!procSheet) return;
  
  const data = procSheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const startTime = new Date().getTime();
  const maxExecutionTime = 240000; 
  
  const headers = data[0];
  const bNameIdx = headers.indexOf("Broadcast_Name");
  const phoneIdx = headers.indexOf("WhatsApp_Number");
  const syncIdx = headers.indexOf("Synced_to_Contacts");
  
  if (bNameIdx === -1 || phoneIdx === -1 || syncIdx === -1) return;

  const groupCache = {};
  try {
    const groupsResponse = People.ContactGroups.list({pageSize: 1000});
    const existingGroups = groupsResponse.contactGroups || [];
    for (let g = 0; g < existingGroups.length; g++) {
      groupCache[existingGroups[g].name] = existingGroups[g].resourceName;
    }
  } catch (e) {
    console.error("People API Error: Failed to fetch contact groups.");
    return; 
  }

  for (let i = 1; i < data.length; i++) {
    if (new Date().getTime() - startTime > maxExecutionTime) break; 
    
    const isSynced = data[i][syncIdx];
    
    if (isSynced === "Pending" || isSynced === "Syncing...") {
      const broadcastName = String(data[i][bNameIdx]).trim();
      const phone = String(data[i][phoneIdx]).trim();
      
      if (!broadcastName || !phone) continue;
      
      const nameParts = broadcastName.split("_");
      const labelName = nameParts.length >= 2 ? `${nameParts[0]}_${nameParts[1]}` : "Unknown_Group";
      
      let groupId = groupCache[labelName];
      if (!groupId) {
        try {
          const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
          groupId = newGroup.resourceName; 
          groupCache[labelName] = groupId;
        } catch (e) {
          procSheet.getRange(i + 1, syncIdx + 1).setValue("Error: Group creation failed");
          continue;
        }
      }
      
      const contact = {
        names: [{ givenName: broadcastName }],
        phoneNumbers: [{ value: phone, type: "mobile" }],
        memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
      };
      
      let success = false;
      
      for (let retry = 0; retry < 3; retry++) {
        try {
          People.People.createContact(contact);
          success = true;
          break; 
        } catch (e) {
          const lastError = e.message.toLowerCase();
          if (lastError.includes("invalid") || lastError.includes("bad request") || lastError.includes("not found")) {
            procSheet.getRange(i + 1, syncIdx + 1).setValue("API Error: Invalid payload");
            break; 
          }
          Utilities.sleep(2000 * (retry + 1)); 
        }
      }

      if (success) {
        procSheet.getRange(i + 1, syncIdx + 1).setValue("Yes");
        Utilities.sleep(300); 
      } else if (!procSheet.getRange(i + 1, syncIdx + 1).getValue().toString().includes("API Error")) {
        procSheet.getRange(i + 1, syncIdx + 1).setValue("Error: Blocked/Rate Limit");
        break; 
      }
    }
  }
}

// ==========================================
// 4. HELPER FUNCTIONS
// ==========================================

function isDuplicatePhone(sheet, phone) {
  if (sheet.getLastRow() < 2) return false;
  const data = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues().flat();
  return data.includes(phone);
}

function getFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ==========================================
// 5. CSV EXPORT LOGIC
// ==========================================

function showCsvPopup() {
  const checkboxesHtml = CATEGORIES.sort().map(cat =>
    `<label class="cb-container"><input type="checkbox" value="${cat}" class="cat-checkbox"> ${cat}</label>`
  ).join('');
  const html = HtmlService.createHtmlOutput(getPopupHtml('📥 WhatsApp CSV Export', checkboxesHtml, 'runExport()', 'Generate CSV Files')).setWidth(450).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'WhatsApp CSV Exporter');
}

function executeWhatsAppExport(selectedCategories) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(S_PROCESSED);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const bNameIdx = headers.indexOf("Broadcast_Name");
  const phoneIdx = headers.indexOf("WhatsApp_Number");
  const catIdx = headers.indexOf("category");

  const csvGroups = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const category = row[catIdx];

    if (selectedCategories.includes(category)) {
      const bName = row[bNameIdx];
      const phone = row[phoneIdx];
      
      const nameParts = String(bName).split("_");
      const filePrefix = (nameParts.length >= 2) ? `${nameParts[0]}_${nameParts[1]}` : "Others";

      if (!csvGroups[filePrefix]) {
        csvGroups[filePrefix] = ["Name,Phone"];
      }
      csvGroups[filePrefix].push(`"${bName}","${phone}"`);
    }
  }

  const folder = getFolder(EXPORT_FOLDER_NAME);
  let fileCount = 0;

  for (let prefix in csvGroups) {
    const fileName = `${prefix}_WhatsApp.csv`;
    
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) { existingFiles.next().setTrashed(true); }

    folder.createFile(Utilities.newBlob(csvGroups[prefix].join("\n"), MimeType.CSV, fileName));
    fileCount++;
  }

  SpreadsheetApp.getUi().alert(`✅ Export Complete!\nGenerated ${fileCount} group files (capped at max 200 per file).`);
}

// ==========================================
// 6. MANUAL SYNC LOGIC (UI CHUNKING)
// ==========================================

function showSyncPopup() {
  const checkboxesHtml = CATEGORIES.sort().map(cat =>
    `<label class="cb-container"><input type="checkbox" value="${cat}" class="cat-checkbox"> ${cat}</label>`
  ).join('');
  const html = HtmlService.createHtmlOutput(getPopupHtml('📱 Manual Contacts Sync', checkboxesHtml, 'startSync()', 'Sync Checked Categories')).setWidth(450).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Google Contacts Sync');
}

function executeContactSync(selectedCategories, startIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(S_PROCESSED);
  if (!sheet) throw new Error("Sheet 'Processed_data' not found!");

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'complete', syncedThisBatch: 0 };

  const headers = data[0];
  const bNameIdx = headers.indexOf("Broadcast_Name");
  const phoneIdx = headers.indexOf("WhatsApp_Number");
  const catIdx = headers.indexOf("category"); 
  const syncIdx = headers.indexOf("Synced_to_Contacts");

  const groupCache = {};
  try {
    const groupsResponse = People.ContactGroups.list({pageSize: 1000});
    const existingGroups = groupsResponse.contactGroups || [];
    for (let g = 0; g < existingGroups.length; g++) {
      groupCache[existingGroups[g].name] = existingGroups[g].resourceName;
    }
  } catch (e) {
    throw new Error("Could not load Contact Groups.");
  }

  let syncedThisBatch = 0;
  const startTime = new Date().getTime();

  for (let i = startIndex; i < data.length; i++) {
    
    if (new Date().getTime() - startTime > 240000) { 
      return { status: 'partial', nextRow: i, syncedThisBatch: syncedThisBatch };
    }

    const row = data[i];
    const category = row[catIdx] ? row[catIdx].toString().trim() : "";
    
    if (!selectedCategories.includes(category)) continue;

    const broadcastName = String(row[bNameIdx]).trim();
    const phone = String(row[phoneIdx]).trim();
    const isSynced = row[syncIdx];

    if (isSynced !== "Yes" && broadcastName !== "" && phone !== "") {
      const nameParts = broadcastName.split("_");
      const labelName = nameParts.length >= 2 ? `${nameParts[0]}_${nameParts[1]}` : "Unknown_Group";

      let groupId = groupCache[labelName];
      if (!groupId) {
        const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
        groupId = newGroup.resourceName; 
        groupCache[labelName] = groupId; 
      }

      const newContact = {
        names: [{ givenName: broadcastName }],
        phoneNumbers: [{ value: phone, type: "mobile" }],
        memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
      };

      let success = false;
      let lastError = "";
      
      for (let retry = 0; retry < 3; retry++) {
        try {
          People.People.createContact(newContact);
          success = true;
          break; 
        } catch (e) {
          lastError = e.message;
          const lowerErr = lastError.toLowerCase();
          if (lowerErr.includes("invalid") || lowerErr.includes("bad request") || lowerErr.includes("not found")) {
            sheet.getRange(i + 1, syncIdx + 1).setValue("API Error: Invalid Payload");
            break; 
          }
          Utilities.sleep(2000 * (retry + 1)); 
        }
      }

      if (success) {
        sheet.getRange(i + 1, syncIdx + 1).setValue("Yes");
        Utilities.sleep(500); 
        syncedThisBatch++;
      } else if (!sheet.getRange(i + 1, syncIdx + 1).getValue().toString().includes("API Error")) {
        sheet.getRange(i + 1, syncIdx + 1).setValue("Error: Blocked/Rate Limit");
        throw new Error("Connection limits reached. Please try again later.");
      }
    }
  }

  return { status: 'complete', syncedThisBatch: syncedThisBatch };
}

// ==========================================
// 7. UI POPUP HTML GENERATOR
// ==========================================

function getPopupHtml(title, checkboxes, fnName, btnText) {
  return `
  <!DOCTYPE html>
  <html>
    <head>
      <style>
        body { font-family: sans-serif; padding: 15px; }
        .scroll-box { max-height: 300px; overflow-y: auto; border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; }
        .cb-container { display: block; margin-bottom: 5px; cursor: pointer; }
        button { width: 100%; padding: 10px; background: #25D366; color: white; border: none; cursor: pointer; font-weight: bold; }
        .btn-select { background: #f0f0f0; color: #333; margin-bottom: 10px; font-size: 12px; }
      </style>
    </head>
    <body>
      <h3>${title}</h3>
      <button class="btn-select" onclick="selectAll()">Check/Uncheck All</button>
      <div class="scroll-box">${checkboxes}</div>
      <button id="runBtn" onclick="${fnName}">${btnText}</button>
      <script>
        let allChecked = false;
        let totalSynced = 0;
        
        function selectAll() {
          allChecked = !allChecked;
          document.querySelectorAll('.cat-checkbox').forEach(cb => cb.checked = allChecked);
        }
        
        function runExport() {
          const cats = Array.from(document.querySelectorAll('.cat-checkbox:checked')).map(cb => cb.value);
          if (cats.length === 0) return alert("Select at least one category.");
          document.getElementById('runBtn').innerText = "Processing...";
          google.script.run.withSuccessHandler(() => google.script.host.close()).executeWhatsAppExport(cats);
        }

        function startSync() {
          const cats = Array.from(document.querySelectorAll('.cat-checkbox:checked')).map(cb => cb.value);
          if(cats.length === 0) { alert('Select at least one category.'); return; }
          document.getElementById('runBtn').disabled = true; 
          document.getElementById('runBtn').style.backgroundColor = "#999";
          totalSynced = 0;
          runSyncBatch(1, cats);
        }

        function runSyncBatch(startRowIndex, cats) {
          const btn = document.getElementById('runBtn');
          if (startRowIndex === 1) btn.innerText = "⏳ Syncing... DO NOT close window!";
          
          google.script.run
            .withSuccessHandler(function(response) {
              totalSynced += response.syncedThisBatch; 
              if (response.status === 'partial') {
                btn.innerText = "⏳ Synced " + totalSynced + " so far... Continuing...";
                runSyncBatch(response.nextRow, cats);
              } else {
                alert("✅ All Done! Successfully synced " + totalSynced + " total contacts.");
                google.script.host.close();
              }
            })
            .withFailureHandler(function(err) {
              alert("Error processing sync.");
              btn.innerText = "🚀 Retry Sync";
              btn.disabled = false; 
              btn.style.backgroundColor = "#25D366";
            }).executeContactSync(cats, startRowIndex);
        }
      </script>
    </body>
  </html>`;
}
