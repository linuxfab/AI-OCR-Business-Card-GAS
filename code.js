
// --- 配置區 ---

const GOOGLE_SHEET_ID = PropertiesService.getScriptProperties().getProperty("GOOGLE_SHEET_ID");
const SHEET_NAME = "工作表1"; 

// 🎯【新功能】必填：請填入存放 AppSheet 圖片的「資料夾 ID」
// 範例："1zKx... (從網址列複製)"
const IMAGE_FOLDER_ID = PropertiesService.getScriptProperties().getProperty("IMAGE_FOLDER_ID");

const DATA_COLUMN_MAPPING = {
  START_COLUMN: 2, 
  NUMBER_OF_COLUMNS_TO_WRITE: 5 
};

// --- 配置區結束 ---

function processBusinessCard(imageFileName, rowNumber) {
  Logger.log('=== AppSheet 呼叫開始 (指定資料夾版) ===');
  
  const rowIndex = parseInt(rowNumber);
  if (isNaN(rowIndex) || rowIndex < 2) {
    Logger.log('✗ 傳入的行號無效');
    return;
  }
  
  try {
    // 1. 取得圖片檔案 (改用資料夾 ID 搜尋)
    const imageFile = getFileFromDrive(imageFileName);
    
    if (!imageFile) {
      Logger.log('✗ 失敗：在指定資料夾中找不到檔案: ' + imageFileName);
      Logger.log('  請確認 1. 資料夾 ID 正確 2. 檔案確實存在於該資料夾');
      return;
    }
    
    // 2. 呼叫 Gemini API
    Logger.log('正在呼叫 Gemini API...');
    const ocrResults = callGeminiOCR(imageFile);
    
    if (!ocrResults) {
      Logger.log('✗ Gemini 分析失敗');
      return;
    }
    
    Logger.log('Gemini 回傳: ' + JSON.stringify(ocrResults));

    // 3. 準備寫入
    const dataToWrite = [
      ocrResults.Name || "",
      ocrResults.Phone || "",
      ocrResults.Email || "",
      ocrResults.Address || "",
      ocrResults.Company || ""
    ];
    
    // 4. 寫入 Sheet
    const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const startCol = DATA_COLUMN_MAPPING.START_COLUMN;
    const numCols = DATA_COLUMN_MAPPING.NUMBER_OF_COLUMNS_TO_WRITE;
    
    sheet.getRange(rowIndex, startCol, 1, numCols).setValues([dataToWrite]);
    SpreadsheetApp.flush();
    
    Logger.log('✓ 成功寫入資料');
    
  } catch (error) {
    Logger.log('✗ 發生錯誤: ' + error.toString());
    Logger.log('堆疊: ' + error.stack);
  }
}

/**
 * 修正後的搜尋函數：直接去指定資料夾找，不再全域搜尋
 */
function getFileFromDrive(filePath) {
  try {
    const cleanFileName = filePath.split('/').pop(); 
    Logger.log('前往資料夾 ID: ' + IMAGE_FOLDER_ID);
    Logger.log('搜尋檔案名稱: ' + cleanFileName);
    
    if (!IMAGE_FOLDER_ID || IMAGE_FOLDER_ID === "請在此貼上您的資料夾ID") {
      throw new Error("請先在程式碼上方設定 IMAGE_FOLDER_ID");
    }

    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    const files = folder.getFilesByName(cleanFileName);
    
    if (files.hasNext()) {
      Logger.log('✓ 找到檔案了');
      return files.next();
    } else {
      Logger.log('✗ 資料夾內無此檔案');
      return null;
    }
  } catch (e) {
    Logger.log('取得檔案錯誤: ' + e.toString());
    throw e;
  }
}

function callGeminiOCR(file) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_KEY");
  if (!apiKey) throw new Error("找不到 API Key (請檢查 Script Properties)");

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const blob = file.getBlob();
  const base64Image = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();
  
  // 1. 🎯 定義 JSON 結構 (Schema) - 確保輸出一致性
  const businessCardSchema = {
    "type": "object",
    "properties": {
      "Name": { "type": "string", "description": "從名片中提取的人名，如果找不到則為空字串。" },
      "Phone": { "type": "string", "description": "從名片中提取的電話號碼，如果找不到則為空字串。" },
      "Email": { "type": "string", "description": "從名片中提取的電子郵件地址，如果找不到則為空字串。" },
      "Address": { "type": "string", "description": "從名片中提取的公司地址，如果找不到則為空字串。" },
      "Company": { "type": "string", "description": "從名片中提取的公司名稱，如果找不到則為空字串。" }
    },
    "required": ["Name", "Phone", "Email", "Address", "Company"] // 確保所有欄位都存在於輸出中
  };
  
  // 2. 簡化的 Prompt - 只給予任務指令
  const promptText = `
    Analyze this business card image and extract the required fields (Name, Phone, Email, Address, Company). 
    Use the empty string ("") if a field is not found.
  `;

  const payload = {
    "contents": [{
      "parts": [
        { "text": promptText },
        { "inline_data": { "mime_type": mimeType, "data": base64Image } }
      ]
    }],
    "generationConfig": { 
      // 3. ⭐ 透過 generationConfig 強制指定 JSON Schema
      "response_mime_type": "application/json", 
      "responseJsonSchema": businessCardSchema
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(apiUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error("Gemini API Error: " + response.getContentText());
  }

  const responseJson = JSON.parse(response.getContentText());
  
  // 4. 簡化解析邏輯 (假設 Gemini 會遵守 Schema 並只輸出 JSON)
  const jsonOutputText = responseJson.candidates[0].content.parts[0].text;
  
  // 注意：即使強制要求 JSON，API 仍可能將 JSON 包裹在 Markdown 塊中。
  // 我們再次使用更強韌的解析方式，確保腳本不會因為多餘的 ```json 而崩潰。
  try {
      const cleanJsonText = jsonOutputText.trim().replace(/^```json\s*|(?:\s*```)?$/g, '');
      return JSON.parse(cleanJsonText);
  } catch (e) {
      Logger.log("警告：JSON 解析失敗，可能是 API 輸出格式不標準。原始輸出：" + jsonOutputText);
      throw new Error("無法解析 Gemini 回傳的 JSON 結構。");
  }
}