  /**
   * HỆ THỐNG QUẢN LÝ KHO V12.19 - FINAL LOGIC (UNCLE EDITION)
   * Update: Optimized Performance, Column H for Batch Price, Infinite Loop Fix.
   */

  const CONFIG = {
    SHEET_NVL: "Danh mục NVL",         
    SHEET_PROD: "Danh mục Product",    
    SHEET_KIT_KITCHEN: "Kit Semi Store",
    SHEET_KIT_PIZZA: "Pzz Semi Store",
    SHEET_KIT_SERVICE: "Semi Store Service", 
    SHEET_ML_NVL: "Học máy",    
    SHEET_ML_PROD: "Học máy Product",  
    SHEET_ML_SEMI: "Học máy Semi Store",
    SHEET_ML_TRANSFER: "Học máy Transfer",
    SHEET_STORE_LIST: "Học máy Store",
    SHEET_TRANSFER_DATA: "Transfer",
    SHEET_SPOILAGE: "Hủy NVL",           
    SHEET_SPOILAGE_SEMI: "Hủy Semi Store",
    SHEET_SPOILAGE_PROD: "Hủy Product",
    SHEET_BOM_PRODUCT: "BOM Product",
    SHEET_BOM_CACHE: "DB_BOM_CACHE",
  };
/**
 * ========================================
 * PUSH CODE LÊN GITHUB TỰ ĐỘNG
 * ========================================
 * Chức năng:
 * - Đọc toàn bộ file từ Apps Script
 * - Đẩy lên GitHub qua REST API
 * - Tự động tạo commit message với timestamp
 * - Hỗ trợ nhiều file (code.gs, HTML, CSS, JS)
 * ========================================
 */

// ⚙️ CẤU HÌNH GITHUB
const GITHUB_CONFIG = {
  OWNER: "nhanca666",
  TOKEN: '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',
  REPO: 'nhanca666/gas-inventory-system',
  BRANCH: 'main'
};
// ============ doGet - Entry Point cho Web App ============
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('V12.8 Analyst Center | TRUNG TÂM PHÂN TÍCH').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
/**
 * HÀM CHÍNH: PUSH CODE LÊN GITHUB
 */
function pushToGitHub() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Xác nhận trước khi push
    const confirm = ui.alert(
      '🚀 Push to GitHub', 
      'Bạn có chắc muốn đẩy code lên GitHub?\n\nToàn bộ thay đổi sẽ được commit.', 
      ui.ButtonSet.YES_NO
    );
    
    if (confirm !== ui.Button.YES) {
      ui.alert('❌ Đã hủy push.');
      return;
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast('🔄 Đang đọc files...', 'Push to GitHub', 5);
    
    // Danh sách file cần push
    const filesToPush = [
      { name: 'code.gs', type: 'gs' },
      { name: 'Index.html', type: 'html' },
      { name: 'Drawer.html', type: 'html' },
      { name: 'Footer.html', type: 'html' },
      { name: 'Javascript.html', type: 'html' },
      { name: 'main.html', type: 'html' },
      { name: 'section.html', type: 'html' },
      { name: 'Stylesheet.html', type: 'html' }
    ];
    
    let successCount = 0;
    let errorFiles = [];
    
    // Timestamp cho commit message
    const timestamp = new Date().toISOString();
    const commitPrefix = `// Last sync: ${timestamp}\n// Auto-pushed from Apps Script\n// User: ${Session.getActiveUser().getEmail()}\n\n`;
    
    // Đẩy từng file
    filesToPush.forEach(function(fileConfig) {
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `📤 Đang push ${fileConfig.name}...`, 
          'GitHub Sync', 
          3
        );
        
        // Đọc nội dung file
        let content = '';
        if (fileConfig.type === 'gs') {
          // Đọc file .gs (Apps Script)
          content = getScriptContent(fileConfig.name);
        } else {
          // Đọc file HTML
          content = HtmlService.createHtmlOutputFromFile(fileConfig.name).getContent();
        }
        
        // Thêm header thông tin sync vào đầu file code.gs
        if (fileConfig.name === 'code.gs') {
          content = commitPrefix + content;
        }
        
        // Push lên GitHub
        const result = pushFileToGitHub(fileConfig.name, content, timestamp);
        
        if (result.success) {
          successCount++;
        } else {
          errorFiles.push(fileConfig.name + ': ' + result.message);
        }
        
      } catch (e) {
        errorFiles.push(fileConfig.name + ': ' + e.toString());
      }
    });
    
    // Thông báo kết quả
    if (successCount === filesToPush.length) {
      ui.alert(
        '✅ Push thành công!', 
        `Đã đẩy ${successCount}/${filesToPush.length} files lên GitHub.\n\n` +
        `📅 Timestamp: ${timestamp}\n` +
        `🔗 Repo: github.com/${GITHUB_CONFIG.OWNER}/${GITHUB_CONFIG.REPO}`, 
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '⚠️ Push một phần', 
        `Thành công: ${successCount}/${filesToPush.length}\n\n` +
        `Lỗi:\n${errorFiles.join('\n')}`, 
        ui.ButtonSet.OK
      );
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Lỗi: ' + e.toString());
  }
}

/**
 * HÀM HỖ TRỢ: ĐẨY 1 FILE LÊN GITHUB
 */
function pushFileToGitHub(fileName, content, timestamp) {
  try {
    const apiBase = `https://api.github.com/repos/${GITHUB_CONFIG.OWNER}/${GITHUB_CONFIG.REPO}`;
    
    // Bước 1: Lấy SHA của file hiện tại (nếu có)
    let currentSHA = null;
    try {
      const getUrl = `${apiBase}/contents/${fileName}?ref=${GITHUB_CONFIG.BRANCH}`;
      const getResponse = UrlFetchApp.fetch(getUrl, {
        headers: {
          'Authorization': 'token ' + GITHUB_CONFIG.TOKEN,
          'Accept': 'application/vnd.github.v3+json'
        },
        muteHttpExceptions: true
      });
      
      if (getResponse.getResponseCode() === 200) {
        const fileData = JSON.parse(getResponse.getContentText());
        currentSHA = fileData.sha;
      }
    } catch (e) {
      // File chưa tồn tại - sẽ tạo mới
    }
    
    // Bước 2: Encode content sang Base64
    const base64Content = Utilities.base64Encode(content);
    
    // Bước 3: Tạo commit message
    const commitMessage = `[AUTO] Update ${fileName} - ${timestamp}`;
    
    // Bước 4: Tạo payload
    const payload = {
      message: commitMessage,
      content: base64Content,
      branch: GITHUB_CONFIG.BRANCH
    };
    
    // Thêm SHA nếu file đã tồn tại (để update)
    if (currentSHA) {
      payload.sha = currentSHA;
    }
    
    // Bước 5: Push lên GitHub
    const putUrl = `${apiBase}/contents/${fileName}`;
    const putResponse = UrlFetchApp.fetch(putUrl, {
      method: 'PUT',
      headers: {
        'Authorization': 'token ' + GITHUB_CONFIG.TOKEN,
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseCode = putResponse.getResponseCode();
    
    if (responseCode === 200 || responseCode === 201) {
      return { success: true, message: 'OK' };
    } else {
      const errorData = JSON.parse(putResponse.getContentText());
      return { 
        success: false, 
        message: errorData.message || 'Unknown error' 
      };
    }
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ĐỌC NỘI DUNG FILE APPS SCRIPT - PHIÊN BẢN FIX HOÀN CHỈNH
 * @param {string} fileName - Tên file cần đọc (vd: 'code.gs')
 * @return {string} Nội dung file
 */
function getScriptContent(fileName) {
  try {
    Logger.log(`📖 Đang đọc file: ${fileName}`);
    
    // MAP TÊN FILE .GS → .HTML
    const fileMap = {
      'code.gs': 'Backend',           // ← QUAN TRỌNG: Không có .html
      'Main_Section.html': 'Main_Section',
      'index.html': 'index',
      'Drawer.html': 'Drawer',
      'section.html': 'section',
      'Footer.html': 'Footer',
      'Javascript.html': 'Javascript',
      'Stylesheet.html': 'Stylesheet',
      'test.gs': 'test'
    };
    
    const htmlFileName = fileMap[fileName] || fileName.replace('.gs', '').replace('.html', '');
    
    Logger.log(`🔄 Map ${fileName} → ${htmlFileName}.html`);
    
    // ĐỌC FILE HTML
    const html = HtmlService.createHtmlOutputFromFile(htmlFileName).getContent();
    
    // NẾU LÀ FILE .GS → TRÍCH XUẤT CODE TỪ <script>
    if (fileName.endsWith('.gs')) {
      Logger.log('🔍 Tìm code trong <script>...</script>');
      
      const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/);
      
      if (!scriptMatch || !scriptMatch[1]) {
        throw new Error(`❌ Không tìm thấy code trong <script> của ${htmlFileName}.html`);
      }
      
      const code = scriptMatch[1].trim();
      Logger.log(`✅ Đọc thành công: ${code.length} ký tự`);
      
      return code;
    }
    
    // NẾU LÀ FILE .HTML → TRẢ VỀ NGUYÊN BẢN
    Logger.log(`✅ Đọc thành công: ${html.length} ký tự`);
    return html;
    
  } catch (e) {
    Logger.log(`❌ Lỗi đọc ${fileName}: ${e.toString()}`);
    throw new Error(`❌ Lỗi đọc ${fileName}: ${e.toString()}`);
  }
}

/**
 * HÀM KIỂM TRA: TEST GITHUB CONNECTION
 */
function testGitHubConnection() {
  try {
    const url = `https://api.github.com/repos/${GITHUB_CONFIG.OWNER}/${GITHUB_CONFIG.REPO}`;
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'token ' + GITHUB_CONFIG.TOKEN,
        'Accept': 'application/vnd.github.v3+json'
      }
    });
    
    const repo = JSON.parse(response.getContentText());
    
    SpreadsheetApp.getUi().alert(
      '✅ Kết nối thành công!',
      `Repo: ${repo.full_name}\n` +
      `Branch: ${GITHUB_CONFIG.BRANCH}\n` +
      `Private: ${repo.private}\n` +
      `Last push: ${repo.pushed_at}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '❌ Lỗi kết nối',
      'Chi tiết: ' + e.toString() + '\n\n' +
      '💡 Kiểm tra:\n' +
      '1. GitHub Token còn hiệu lực?\n' +
      '2. Token có quyền "repo"?\n' +
      '3. Repository name đúng?',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * HÀM HỖ TRỢ: LẤY THÔNG TIN COMMIT MỚI NHẤT
 */
function getLatestCommit() {
  try {
    const url = `https://api.github.com/repos/${GITHUB_CONFIG.OWNER}/${GITHUB_CONFIG.REPO}/commits?per_page=1`;
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'token ' + GITHUB_CONFIG.TOKEN,
        'Accept': 'application/vnd.github.v3+json'
      }
    });
    
    const commits = JSON.parse(response.getContentText());
    const latest = commits[0];
    
    return {
      sha: latest.sha.substring(0, 7),
      message: latest.commit.message,
      author: latest.commit.author.name,
      date: latest.commit.author.date
    };
    
  } catch (e) {
    return null;
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 GITHUB SYNC')
    .addItem('📤 Push to GitHub', 'pushToGitHub')
    .addItem('🧪 Test Connection', 'testGitHubConnection')
    .addItem('📊 View Latest Commit', 'viewLatestCommit')
    .addSeparator()
    .addItem('💬 Open Claude Project', 'openClaudeProject')
    .addToUi();
}
function openClaudeProject() {
  const url = 'https://claude.ai/project/019b4ff4-028a-7786-b048-da3dda956661';
  const html = '<script>window.open("' + url + '");google.script.host.close();</script>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html),
    'Opening Claude...'
  );
}

function viewLatestCommit() {
  const commit = getLatestCommit();
  if (commit) {
    SpreadsheetApp.getUi().alert(
      '📝 Latest Commit',
      `SHA: ${commit.sha}\n` +
      `Message: ${commit.message}\n` +
      `Author: ${commit.author}\n` +
      `Date: ${commit.date}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
  // --- CORE SYSTEM DATA ---
  function getSystemData() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Hàm ép kiểu số an toàn (xử lý dấu phẩy Việt Nam)
      const safeFloat = (val) => {
          if (typeof val === 'number') return val;
          if (!val) return 0;
          return Number(String(val).replace(/,/g, '.')) || 0;
      };

      const getData = (name) => {
        const s = ss.getSheetByName(name);
        return s ? s.getDataRange().getValues().slice(1) : [];
      };
      
      const cleanRawRow = (row) => row.map(cell => (cell instanceof Date ? formatDate(cell) : cell));
      
      // 1. Master Data
      // NVL: Giá vốn ở Cột D (Index 3)
      const listNVL = getData(CONFIG.SHEET_NVL).map((r, i) => ({ 
        rowId: i + 2, 
        name: String(r[0]),           // A: Tên
        code: cleanCode(r[1]),        // B: Mã
        unit: String(r[2]),           // C: ĐVT Gốc
        cost: safeFloat(r[3]),        // D: Giá Vốn
        stdUnit: String(r[4]),        // E: ĐVT Quy đổi (Quan trọng)
        rate: Number(r[5]) || 1,      // F: Hệ số
        buyingPrice: safeFloat(r[6]), // G: Giá Mua (MỚI)
        supplier: String(r[7]),       // H: Nhà cung cấp (MỚI)
        leadtime: String(r[8]),       // I: Leadtime (MỚI)
        noDelivery: String(r[9]),     // J: No Delivery (MỚI)
        group: String(r[10]),         // K: Group hàng (MỚI)
        status: String(r[11]),        // L: Trạng thái (MỚI)
        type: 'NVL', 
        rawData: cleanRawRow(r) 
      }));

      // PROD: Giá vốn ở Cột G (Index 6)
      const listProd = getData(CONFIG.SHEET_PROD).map((r, i) => ({ 
        rowId: i + 2, 
        code: cleanCode(r[0]), 
        name: String(r[1]),    
        category: String(r[2]),
        class: String(r[3]),   
        team: String(r[4]),    
        sapCode: String(r[5]), 
        cost: safeFloat(r[6]), 
        status: String(r[7]) || 'Active', // Lấy thêm trạng thái
        rate: 1, unit: 'Cái', standardUnit: 'Cái', type: 'PROD', 
        rawData: cleanRawRow(r) 
      }));

      // SEMI STORE: Đọc giá từ Sheet (Không tính toán lại để load nhanh)
      const listSemi = [];
      [CONFIG.SHEET_KIT_KITCHEN, CONFIG.SHEET_KIT_PIZZA, CONFIG.SHEET_KIT_SERVICE].forEach(sheetName => {
          getData(sheetName).forEach(r => {
              if (String(r[1]).toLowerCase() === 'parent') {
                  listSemi.push({ 
                      rowId: -1, 
                      code: cleanCode(r[0]), 
                      name: String(r[3]), 
                      unit: String(r[5]), 
                      
                      // [CHÚ IT UPDATE] Đọc thẳng giá từ Sheet
                      cost: safeFloat(r[6]),       // Cột G: Giá vốn đơn vị
                      batchPrice: safeFloat(r[7]), // Cột H: Giá Batch (Mới)
                      
                      yield: safeFloat(r[4]),      // Cột E: Định lượng
                      rate: 1, 
                      type: 'SEMI', 
                      sourceSheet: sheetName, 
                      rawData: cleanRawRow(r) 
                  });
              }
          });
      });

      // Merge Master Data
      const masterData = [...listNVL, ...listProd];
      listSemi.forEach(semiItem => {
          const existing = masterData.find(m => String(m.code) === String(semiItem.code));
          if (existing) { 
              existing.sourceSheet = semiItem.sourceSheet; 
              if (existing.type === 'NVL') existing.type = 'SEMI'; 
              // Cập nhật giá cho mã đã tồn tại
              existing.cost = semiItem.cost;
              existing.batchPrice = semiItem.batchPrice;
              existing.yield = semiItem.yield;
          } 
          else { masterData.push(semiItem); }
      });

      // 2. Stores & ML
      const stores = getData(CONFIG.SHEET_STORE_LIST).filter(r => r[0]).map(r => ({ code: String(r[0]).trim(), keywords: [String(r[0]).trim().toLowerCase()] }));
      const mlData = [];
      const processML = (sheetName, type, contextIdx) => {
          getData(sheetName).forEach(r => { if(r[0] && r[2]) mlData.push({ term: String(r[0]).toLowerCase().trim(), dept: contextIdx !== null ? String(r[contextIdx]).trim() : '', store: (type === 'ML_TRANS' && contextIdx !== null) ? String(r[contextIdx]).trim() : '', code: cleanCode(r[2]), type: type }); });
      };
      processML(CONFIG.SHEET_ML_NVL, 'ML_NVL', 1);
      processML(CONFIG.SHEET_ML_PROD, 'ML_PROD', 1);
      processML(CONFIG.SHEET_ML_SEMI, 'ML_SEMI', 1);
      processML(CONFIG.SHEET_ML_TRANSFER, 'ML_TRANS', 1);

      // 3. History Transfer
      const transferHistory = [];
      const sheetHistory = ss.getSheetByName(CONFIG.SHEET_TRANSFER_DATA);
      if (sheetHistory) {
        const lastRow = sheetHistory.getLastRow();
        const startRow = Math.max(2, lastRow - 2000);
        if (lastRow >= 2) {
          const numRows = lastRow - startRow + 1;
          const values = sheetHistory.getRange(startRow, 1, numRows, sheetHistory.getLastColumn()).getValues();
          for (let i = 0; i < values.length; i++) {
            const r = values[i];
            const statusVal = r[16]; 
            if (statusVal === true || statusVal === false || String(statusVal).toUpperCase() === 'TRUE' || String(statusVal).toUpperCase() === 'FALSE') {
              let totalBaseQty = 0;
              let valM = r[12]; // Col M
              if (typeof valM === 'number') totalBaseQty = valM;
              else if (valM) totalBaseQty = Number(String(valM).replace(/,/g, '').trim());
              if (isNaN(totalBaseQty)) totalBaseQty = 0;

              let rawUnit = safeValue(r[11]); // Col L
              let isDivided = (r[7] !== "" && r[7] != null);
              const code = safeValue(r[6]);
              const masterItem = masterData.find(m => String(m.code) === String(code));
              const rate = masterItem ? masterItem.rate : (Number(r[9]) || 1);
              const standardUnit = masterItem ? (masterItem.standardUnit || masterItem.unit) : rawUnit;
              let displayQty = isDivided ? (totalBaseQty / rate) : totalBaseQty;
              let displayUnit = isDivided ? standardUnit : rawUnit;

              transferHistory.push({
                id: startRow + i, date: formatDate(safeValue(r[0])),
                sender: (r[1] && r[3]) ? `${r[1]} \u2194 ${r[3]}` : safeValue(r[1]),
                realSender: safeValue(r[1]), realReceiver: safeValue(r[3]),
                type: String(safeValue(r[2])).toUpperCase(), receiver: safeValue(r[3]), store: safeValue(r[4]),
                itemName: safeValue(r[5]), code: code, rate: rate, team: safeValue(r[13]),
                qty: displayQty, unit: displayUnit, originalQty: totalBaseQty, originalUnit: rawUnit, standardUnit: standardUnit,
                amount: Number(safeValue(r[15])) || 0, status: statusVal === true || String(statusVal).toUpperCase() === 'TRUE', note: safeValue(r[16]),
                rateState: isDivided ? 2 : 0
              });
            }
          }
          transferHistory.reverse();
        }
      }
      
      return { success: true, masterData, stores, mlData, transferHistory, sheetNames: { kit: CONFIG.SHEET_KIT_KITCHEN, pzz: CONFIG.SHEET_KIT_PIZZA, svc: CONFIG.SHEET_KIT_SERVICE } };
    } catch (e) { return { success: false, message: "Server Error: " + e.toString() }; }
  }

  function findNextEmptyRow(sheet) {
    const colA = sheet.getRange("A1:A").getValues();
    for (let i = colA.length - 1; i >= 0; i--) { if (colA[i][0] !== "" && colA[i][0] != null) return i + 2; }
    return 2;
  }

  function smartRound(num) {
    if (num === null || num === undefined || String(num).trim() === '') return 0;
    let val = Number(num); if (isNaN(val)) return 0;
    // [CHÚ IT UPDATE] Luôn giữ 3 số lẻ (0.001) cho mọi trường hợp để đảm bảo chính xác cho Transfer/Hủy
    return Math.round(val * 1000) / 1000;
  }

  /* [TỐI ƯU TỐC ĐỘ] Lưu Hủy: Sử dụng Bulk Insert (Array) cho cả 3 Sheet, tốc độ lưu < 2 giây */
function saveSpoilageData(p){try{const ss=SpreadsheetApp.getActiveSpreadsheet(),sSp=ss.getSheetByName(CONFIG.SHEET_SPOILAGE),sSm=ss.getSheetByName(CONFIG.SHEET_SPOILAGE_SEMI),sPr=ss.getSheetByName(CONFIG.SHEET_SPOILAGE_PROD);if(!sSp||!sSm||!sPr)return{success:false,message:"Thiếu Sheet Hủy!"};const sys=getSystemData(),mMap=new Map(),rMap=getRecipeMap(ss),nMap=createNameMap(ss);sys.masterData.forEach(m=>mMap.set(String(m.code).trim(),m));let d=p.date;if(d.includes("-")){let x=d.split('-');d=`${x[2]}/${x[1]}/${x[0]}`}const rSp=[],rSm=[],rPr=[];p.items.forEach(i=>{const c=String(i.code).trim(),mI=mMap.get(c),uP=mI?Number(mI.cost)||0:0,amt=Math.round(i.qty*uP);if(i.itemType==='PROD'){rPr.push([d,i.name,i.code,"",i.qty,i.note||"",p.dept,"","",mI?mI.category:"",mI?mI.class:"",uP])}else if(i.itemType==='SEMI'){const leaves=bomEngine(c,i.qty,rMap,mMap);const isBk=(leaves.length===1&&String(leaves[0].code).trim()===c);let wn=isBk?" | ⚠️ CHƯA BOM":"";rSp.push([d,c,i.name,"",i.unit,i.factor||1,"",i.unit,i.qty,"",(i.note||"Hủy BTP")+wn,p.dept,"",uP,amt]);if(!isBk){const cons={};leaves.forEach(l=>{cons[l.code]=(cons[l.code]||0)+l.qty_raw});for(const[lC,lQ]of Object.entries(cons)){rSm.push([d,lC,"","","","",smartRound(lQ),"","","",`Bung từ ${i.qty} ${i.name}`,p.dept])}}}else{rSp.push([d,c,i.name,"",i.unit,i.factor||1,"",i.unit,i.qty,"",(i.note||""),p.dept,"",uP,amt])}});if(rSp.length>0)sSp.getRange(findNextEmptyRow(sSp),1,rSp.length,15).setValues(rSp);if(rSm.length>0)sSm.getRange(findNextEmptyRow(sSm),1,rSm.length,12).setValues(rSm);if(rPr.length>0)sPr.getRange(findNextEmptyRow(sPr),1,rPr.length,12).setValues(rPr);return{success:true,message:"✅ Đã lưu phiếu Hủy thành công & Khớp cột!"}}catch(e){return{success:false,message:"Lỗi: "+e.toString()}}}

  function saveTransferData(items) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_TRANSFER_DATA); // [1]
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet Transfer" };

    // 1. Chuẩn bị Map Giá Vốn (Chỉ đọc 1 lần)
    const costMap = new Map();
    [CONFIG.SHEET_NVL, CONFIG.SHEET_PROD].forEach(name => {
      const s = ss.getSheetByName(name);
      if (s) {
        const vals = s.getDataRange().getValues();
        // [2] Xác định cột giá dựa trên tên Sheet
        const isNVL = name === CONFIG.SHEET_NVL; 
        vals.forEach(r => {
          // NVL: Code cột B(1), Giá cột D(3). PROD: Code cột A(0), Giá cột G(6)
          if (isNVL) costMap.set(String(r[3]).trim(), Number(r[4]) || 0);
          else costMap.set(String(r).trim(), Number(r[5]) || 0);
        });
      }
    });

    const outputRows = [];
    const startRow = findNextEmptyRow(sheet); // [6] Chỉ tìm dòng trống 1 lần đầu tiên

    // 2. Xử lý dữ liệu trong bộ nhớ (RAM)
    items.forEach(item => {
      // Xử lý ngày tháng
      let d = item.date;
      if (d.includes("-")) { let p = d.split('-'); d = `${p[7]}/${p[3]}/${p}`; }

      // Logic tính toán Total Base Qty (Quan trọng để trừ kho đúng)
      // [8] Nếu là State 2 (Chia) hoặc 3 (Hack Unit), nhân ngược lại ra số gốc
      let rate = Number(item.rate) || 1;
      let totalBaseQty = 0;
      let qtyDisplay = item.qty;

      if (item.rateState === 2 || item.rateState === 3) {
         // Trường hợp nhập theo Thùng/Quy đổi
         totalBaseQty = item.qty * rate; 
      } else {
         // Trường hợp nhập Lẻ hoặc Nhân
         totalBaseQty = item.qty;
      }
      
      // An toàn: Nếu Frontend có gửi originalQty thì ưu tiên kiểm tra, nhưng logic trên là "Chốt chặn" cuối cùng.
      
      // Lấy giá vốn đơn vị
      const unitCost = costMap.get(String(item.code).trim()) || 0;
      // Tính thành tiền: Phải nhân với TỔNG SỐ LƯỢNG GỐC (Total Base Qty)
      const totalAmount = Math.round(unitCost * totalBaseQty);

      // Chuẩn bị dòng dữ liệu (Mapping theo đúng cột trong Sheet Transfer)
      // [9]-[10] Cấu trúc cột
      outputRows.push([
        d,                                      // A: Ngày
        item.sender,                            // B: Người gửi
        item.type,                              // C: Loại (IN/OUT)
        item.receiver,                          // D: Người nhận
        item.storeName,                         // E: Store
        item.name,                              // F: Tên hàng
        item.code,                              // G: Mã hàng
        (item.rateState === 2 || item.rateState === 3) ? qtyDisplay : "", // H: SL Quy đổi
        item.standardUnit,                      // I: ĐVT Quy đổi
        rate,                                   // J: Hệ số
        (item.rateState === 2 || item.rateState === 3) ? "" : qtyDisplay, // K: SL Lẻ
        item.originalUnit || item.baseUnit || item.unit, // L: ĐVT Gốc (Bắt buộc)
        totalBaseQty,                           // M: Tổng SL Gốc (QUAN TRỌNG NHẤT)
        item.team || "",                        // N: Team
        unitCost,                               // O: Giá vốn
        totalAmount,                            // P: Thành tiền
        false                                   // Q: Checkbox Status
      ]);
    });

    // 3. Ghi xuống Sheet 1 lần duy nhất (Tốc độ cao)
    if (outputRows.length > 0) {
      sheet.getRange(startRow, 1, outputRows.length, 17).setValues(outputRows);
    }

    return { success: true, message: `Đã lưu thành công ${outputRows.length} dòng!` };

  } catch (e) {
    return { success: false, message: "Lỗi Server: " + e.toString() };
  }
}

  /* [ĐỒNG BỘ] Cập nhật phiếu: Logic khớp 100% với saveTransferData để tránh lệch kho */
function updateTransferFull(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_TRANSFER_DATA); // [cite: 73]
    const r = p.id;
    if (!r || r < 2) return { success: false, message: "ID dòng lỗi" };

    // 1. Chuẩn hóa ngày tháng (Sửa lỗi Index mảng p)
    let d = p.date;
    if (d && d.includes("-")) {
      let x = d.split('-');
      d = `${x[2]}/${x[1]}/${x[0]}`; // Định dạng chuẩn DD/MM/YYYY
    }

    // 2. Logic tính toán TotalBaseQty (Giữ nguyên logic p.rateState 2 & 3 như bạn yêu cầu)
    let rate = Number(p.rate) || 1;
    let totalBaseQty = 0;
    if (p.rateState === 2 || p.rateState === 3) {
      totalBaseQty = p.qty * rate;
    } else {
      totalBaseQty = p.qty;
    }

    // 3. Tra cứu giá vốn (Sửa đúng Index theo cấu trúc Master Data)
    let bC = 0;
    const tC = String(p.code || '').trim();
    if (tC) {
      const sNvl = ss.getSheetByName(CONFIG.SHEET_NVL);
      const sPr = ss.getSheetByName(CONFIG.SHEET_PROD);
      
      // Kiểm tra trong NVL: Mã [1], Giá [3]
      if (sNvl) {
        const foundNvl = sNvl.getDataRange().getValues().find(row => String(row[1]).trim() === tC);
        if (foundNvl) bC = Number(foundNvl[3]) || 0;
      }
      
      // Nếu không thấy trong NVL, kiểm tra trong PROD: Mã [0], Giá [6]
      if (bC === 0 && sPr) {
        const foundPr = sPr.getDataRange().getValues().find(row => String(row[0]).trim() === tC);
        if (foundPr) bC = Number(foundPr[6]) || 0;
      }
    }

    // 4. Chuẩn bị dòng dữ liệu (Ghi từ cột A đến P - 16 cột)
    let vals = [[
      d,                                    // A: Ngày
      p.sender || '',                       // B: Người gửi
      p.type || 'UNK',                      // C: Loại
      p.receiver || '',                     // D: Người nhận
      p.store || '',                        // E: Store
      p.itemName || p.name || '',           // F: Tên
      p.code || '',                         // G: Mã
      (p.rateState === 2 || p.rateState === 3) ? p.qty : "", // H: SL Quy đổi
      p.standardUnit || '',                 // I: ĐVT Quy đổi
      rate,                                 // J: Hệ số
      (p.rateState === 2 || p.rateState === 3) ? "" : p.qty, // K: SL Lẻ
      p.originalUnit || p.unit || '',       // L: ĐVT Gốc
      totalBaseQty,                         // M: TỔNG SL GỐC
      p.team || '',                         // N: Team
      bC,                                   // O: Giá vốn đơn vị
      Math.round(bC * totalBaseQty)         // P: Thành tiền
    ]];

    // 5. Ghi dữ liệu - [cite: 179-180]
    sheet.getRange(r, 1, 1, 16).setValues(vals);

    return { success: true, message: "✅ Đã cập nhật phiếu & Đồng bộ logic thành công!" };

  } catch (e) {
    return { success: false, message: "Lỗi Server Update: " + e.toString() };
  }
}


  /**
   * [V12.34 FIX MASTER DATA]
   * - Fix lỗi Semi: Sửa = Xóa Cũ + Thêm Mới (Để cập nhật BOM con).
   * - Fix lỗi NVL/Prod: Cập nhật đầy đủ các trường (NCC, Giá, Group...) khi sửa.
   */
  function updateMasterData(action, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName;
    if (payload.type === 'PROD') sheetName = CONFIG.SHEET_PROD;
    else if (payload.type === 'SEMI') {
      const d = String(payload.dept || '').toLowerCase();
      if (d.includes('pizza') || d.includes('pzz')) sheetName = CONFIG.SHEET_KIT_PIZZA;
      else if (d.includes('service') || d.includes('svc')) sheetName = CONFIG.SHEET_KIT_SERVICE;
      else sheetName = CONFIG.SHEET_KIT_KITCHEN;
    } else { sheetName = CONFIG.SHEET_NVL; }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet: " + sheetName };

    // --- CASE SPECIAL: SEMI (XỬ LÝ DÒNG CHA & CON) ---
    if (action === 'EDIT' && payload.type === 'SEMI' && payload.rowId) {
      const oldCode = sheet.getRange(payload.rowId, 1).getValue();
      const data = sheet.getDataRange().getValues();
      // Xóa từ dưới lên để không lệch Index dòng
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]).trim() === String(oldCode).trim()) { sheet.deleteRow(i + 1); }
      }
      action = 'ADD'; 
    }

    if (action === 'ADD') {
      const row = findNextEmptyRow(sheet);
      if (payload.type === 'PROD') {
        sheet.getRange(row, 1, 1, 8).setValues([[payload.code, payload.name, payload.category||'', payload.class||'', payload.team||'Service', payload.sapCode||'', payload.cost||0, 'Active']]);
      }
      else if (payload.type === 'SEMI') {
        const ingredients = payload.ingredients || [], rowsToInsert = [];
        let totalBatchCost = 0;
        // Map giá vốn NVL: Mã [1], Giá [3] [cite: 11-12]
        const costMap = new Map();
        const sNVL = ss.getSheetByName(CONFIG.SHEET_NVL);
        if(sNVL) sNVL.getDataRange().getValues().forEach(r => costMap.set(String(r[1]).trim(), Number(r[3])||0));

        ingredients.forEach(ing => {
          if (ing.code && ing.qty > 0) {
            const unitPrice = costMap.get(String(ing.code).trim()) || 0;
            const lineTotal = unitPrice * Number(ing.qty);
            totalBatchCost += lineTotal;
            rowsToInsert.push([payload.code, 'Child', ing.code, ing.name, ing.qty, ing.unit, lineTotal, '']);
          }
        });
        const yieldVal = Number(payload.yield) || 1;
        const unitCost = (yieldVal > 0) ? (totalBatchCost / yieldVal) : 0;
        rowsToInsert.unshift([payload.code, 'Parent', payload.code, payload.name, yieldVal, payload.unit, Number(unitCost.toFixed(3)), Math.round(totalBatchCost)]);
        if (rowsToInsert.length > 0) sheet.getRange(row, 1, rowsToInsert.length, 8).setValues(rowsToInsert);
      }
      else { // NVL [cite: 5]
        sheet.getRange(row, 1, 1, 12).setValues([[payload.name, payload.code, payload.unit, payload.cost||0, payload.stdUnit||'', payload.rate||1, payload.buyingPrice||0, payload.supplier||'', payload.leadtime||'', payload.noDelivery||'', payload.group||'', payload.status||'Active']]);
      }
    }

    if (action === 'EDIT' && payload.rowId && payload.type !== 'SEMI') {
      const r = payload.rowId;
      if (payload.type === 'PROD') {
        sheet.getRange(r, 1, 1, 6).setValues([[payload.code, payload.name, payload.category||'', payload.class||'', payload.team||'', payload.sapCode||'']]);
      } else {
        sheet.getRange(r, 1, 1, 3).setValues([[payload.name, payload.code, payload.unit]]);
        sheet.getRange(r, 5, 1, 8).setValues([[payload.stdUnit||'', payload.rate||1, payload.buyingPrice||0, payload.supplier||'', payload.leadtime||'', payload.noDelivery||'', payload.group||'', payload.status||'Active']]);
      }
    }

    // Cập nhật Cache thông minh [cite: 123]
    if (payload.type === 'SEMI' || payload.type === 'PROD') {
      try { updateSingleBOMCache(payload.code); } catch (e) { console.log("Lỗi Cache: " + e.toString()); }
    }
    return { success: true, message: `✅ Đã cập nhật ${payload.code}` };
  } catch (e) { return { success: false, message: "Lỗi: " + e.toString() }; }
}

  function deleteMasterData(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName = (payload.type === 'PROD') ? CONFIG.SHEET_PROD : CONFIG.SHEET_NVL;
    const sheet = ss.getSheetByName(sheetName);
    if (payload.rowId) { sheet.deleteRow(payload.rowId); return { success: true, message: "Đã xóa mã: " + payload.code }; }
    return { success: false, message: "Không tìm thấy ID!" };
  }

  function updateSystemData(action, payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (action === 'ADD_ML') {
      let sheetName = CONFIG.SHEET_ML_NVL;
      if (payload.type === 'ML_TRANS') sheetName = CONFIG.SHEET_ML_TRANSFER;
      else if (payload.type === 'ML_PROD') sheetName = CONFIG.SHEET_ML_PROD;
      else if (payload.type === 'ML_SEMI') sheetName = CONFIG.SHEET_ML_SEMI;
      const sheet = ss.getSheetByName(sheetName);
      if(sheet) sheet.appendRow([payload.term, payload.context, payload.code]);
      return { success: true, message: "Đã cập nhật ML" };
    }
    if (action === 'DELETE_ML') { 
        let sheetName = CONFIG.SHEET_ML_NVL;
        if (payload.type === 'ML_TRANS') sheetName = CONFIG.SHEET_ML_TRANSFER;
        else if (payload.type === 'ML_PROD') sheetName = CONFIG.SHEET_ML_PROD;
        else if (payload.type === 'ML_SEMI') sheetName = CONFIG.SHEET_ML_SEMI;
        const sheet = ss.getSheetByName(sheetName);
        if(sheet) {
            const data = sheet.getDataRange().getValues();
            for(let i=data.length-1; i>=0; i--) { if(String(data[i][0]) === payload.term && String(data[i][2]) === payload.code) { sheet.deleteRow(i+1); break; } }
        }
        return { success: true, message: "Đã xóa ML" };
    }
    return { success: false };
  }

  function updateTransferStatus(rowId, newStatus) { 
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.SHEET_TRANSFER_DATA);
      sheet.getRange(rowId, 17).setValue(newStatus); 
      return {success:true}; 
  }

  function deleteTransferRow(rowId) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEET_TRANSFER_DATA);
        if (!sheet) return { success: false, message: "Sheet Transfer not found" };
        if (!rowId || rowId < 2 || rowId > sheet.getLastRow()) return { success: false, message: "Invalid Row ID" };
        sheet.deleteRow(rowId);
        return { success: true, message: "Đã xóa dòng thành công!" };
      } catch (e) { return { success: false, message: "Server Error: " + e.toString() }; }
  }


  function getBOMDataSimple(ss, sheetName, code) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return null;
      const data = sheet.getDataRange().getValues();
      const codeStr = String(code).trim();
      let header = null;
      let ingredients = [];
      for (let r of data) {
          if (String(r[0]).trim() === codeStr && String(r[1]).toLowerCase() === 'parent') { header = { yield: Number(r[4]) || 1 }; break; }
      }
      if (!header) return null;
      for (let r of data) {
          if (String(r[0]).trim() === codeStr && String(r[1]).toLowerCase() === 'child') { ingredients.push({ code: r[2], name: r[3], qty: Number(r[4]), unit: r[5] }); }
      }
      return { header, ingredients };
  }

  function getBOMDetail(c, s) { 
      const ss = SpreadsheetApp.getActiveSpreadsheet(); 
      
      // [LOGIC MỚI] Nếu nguồn là Danh mục Product -> Đọc từ Sheet BOM Product
      let targetSheetName = s;
      if (s === CONFIG.SHEET_PROD) targetSheetName = CONFIG.SHEET_BOM_PRODUCT;
      
      let sheet = ss.getSheetByName(targetSheetName || CONFIG.SHEET_KIT_KITCHEN);
      if (!sheet) return { header: { yield: 1, unit: 'Batch' }, details: [] }; 
      
      const data = sheet.getDataRange().getValues(); 
      const codeStr = cleanCode(c); 
      let header = { yield: 1, unit: 'Batch' };
      const ing = []; 
      
      for (let r of data) { 
          if (cleanCode(r[0]) === codeStr) { 
              if (String(r[1]).toLowerCase() === 'parent') { 
                  header.yield = Number(r[4]) || 1;
                  header.unit = r[5] || 'Batch'; 
              } else if (String(r[1]).toLowerCase() === 'child') { 
                  ing.push({ code: r[2], name: r[3], qty: Number(r[4]), unit: r[5] });
              } 
          } 
      } 
      return { header, details: ing };
  }

  // --- CORE BOM SAVING (LƯU GIÁ VÀO CỘT G VÀ H) ---
  /* [MODULE MASTER] HÀM LƯU ĐỊNH MỨC BOM - BẢN NGUYÊN KHỐI CHUẨN V12.19 [cite: 1, 2025-12-21] */
function saveBOM(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. KIỂM TRA LOẠI BOM VÀ XÁC ĐỊNH SHEET [cite: 1, 2025-12-21]
    const isProductBOM = (p.sourceSheet === CONFIG.SHEET_PROD);
    const targetSheetName = isProductBOM ? CONFIG.SHEET_BOM_PRODUCT : p.sourceSheet;
    const sheetDetail = ss.getSheetByName(targetSheetName);
    if (!sheetDetail) return { success: false, message: "Thiếu Sheet: " + targetSheetName };

    const itemCode = cleanCode(p.itemCode);

    // 2. XÓA DỮ LIỆU BOM CŨ [cite: 1, 2025-12-21]
    const data = sheetDetail.getDataRange().getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (cleanCode(data[i][0]) === itemCode) {
        sheetDetail.deleteRow(i + 1);
      }
    }

    // 3. TÍNH TOÁN GIÁ VÀ CHUẨN BỊ DỮ LIỆU MỚI [cite: 1, 2025-12-21]
    const sysData = getSystemData();
    const allMaster = sysData.masterData;
    let totalBatchCost = 0;
    const newRows = [];

    // Tạo dòng Header cho Sheet Chi tiết
    let headerRow = [itemCode, 'Parent', itemCode, p.itemName, p.yield, p.yieldUnit, 0, 0];
    
    p.ingredients.forEach(ing => {
      const ingMaster = allMaster.find(m => String(m.code) === String(ing.code) || (m.type === 'PROD' && String(m.sapCode) === String(ing.code)));
      let unitCost = 0;
      if (ingMaster) {
        unitCost = (ingMaster.batchPrice && ingMaster.batchPrice > 0) ? (ingMaster.batchPrice / (ingMaster.yield || 1)) : (Number(ingMaster.cost) || 0);
      }
      const lineTotal = unitCost * Number(ing.qty);
      totalBatchCost += lineTotal;
      newRows.push([itemCode, 'Child', ing.code, ing.name, ing.qty, ing.unit, lineTotal, '']);
    });

    const yieldVal = Number(p.yield) || 1;
    const finalUnitCost = (yieldVal > 0) ? (totalBatchCost / yieldVal) : 0;

    headerRow[6] = Number(finalUnitCost.toFixed(2));
    headerRow[7] = Math.round(totalBatchCost);
    newRows.unshift(headerRow);

    // 4. GHI DỮ LIỆU VÀ CẬP NHẬT CACHE TỰ ĐỘNG [cite: 1, 2025-12-21]
    if (newRows.length > 0) {
      sheetDetail.getRange(sheetDetail.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      
      // Cập nhật Cache riêng lẻ (Partial Update) - Dùng đúng biến itemCode [cite: 1, 2025-12-21]
      try {
        if (itemCode) { 
          updateSingleBOMCache(itemCode); 
        }
      } catch (cacheErr) { 
        console.log("Lỗi Cache: " + cacheErr.toString()); 
      }
    }

    // 5. CẬP NHẬT GIÁ VỐN VÀO DANH MỤC PRODUCT (NẾU CÓ) [cite: 1, 2025-12-21]
    if (isProductBOM) {
      const sheetProd = ss.getSheetByName(CONFIG.SHEET_PROD);
      const prodData = sheetProd.getDataRange().getValues();
      for (let i = 0; i < prodData.length; i++) {
        if (String(prodData[i][0]) === itemCode) {
          sheetProd.getRange(i + 1, 7).setValue(Number(finalUnitCost.toFixed(2)));
          break;
        }
      }
    }

    // 6. PHẢN HỒI KẾT THÚC ĐỂ TẮT LOADING TRÊN WEB [cite: 1, 1809, 2025-12-21]
    return { 
      success: true, 
      message: "Đã lưu BOM món " + p.itemName + " thành công!",
      newCost: finalUnitCost,
      newBatchPrice: totalBatchCost 
    };

  } catch (err) {
    return { success: false, message: "Lỗi hệ thống: " + err.toString() };
  }
} // <--- Đủ dấu đóng kết thúc hàm [cite: 1, 2025-12-21]

  function traceIngredientUsage(nvlCode) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      // [MỚI] Bổ sung SHEET_BOM_PRODUCT vào danh sách đi tìm
      const sheets = [
        { name: CONFIG.SHEET_KIT_KITCHEN, team: 'KITCHEN' },
        { name: CONFIG.SHEET_KIT_PIZZA, team: 'PIZZA' },
        { name: CONFIG.SHEET_KIT_SERVICE, team: 'SERVICE' },
        { name: CONFIG.SHEET_BOM_PRODUCT, team: 'PRODUCT' } // <-- Thêm dòng này
      ];
      
      const results = [];
      const target = String(nvlCode).trim();
      
      sheets.forEach(conf => {
          const sheet = ss.getSheetByName(conf.name);
          if (sheet) {
              const data = sheet.getDataRange().getValues();
              data.forEach(r => {
                  // Cột B (index 1) là 'Child', Cột C (index 2) là Mã Con
                  if (String(r[1]).toLowerCase() === 'child' && String(r[2]).trim() === target) {
                      results.push({ 
                          parentCode: r[0], // Mã Cha
                          team: conf.team,
                          qty: Number(r[4]) || 0, // SL
                          unit: String(r[5]) || '' // ĐVT
                      });
                  }
              });
          }
      });

      // Lọc trùng lặp (Unique)
      const unique = [];
      const map = new Map();
      for (const item of results) {
          // Tạo khóa unique là Mã cha + Team
          const key = item.parentCode + '-' + item.team;
          if(!map.has(key)){
              map.set(key, true);
              unique.push(item);
          }
      }
      return { success: true, data: unique };
    } catch (e) { return { success: false, message: e.toString() }; }
  }

  function cleanCode(code) { return String(code || "").split(",")[0].trim(); }
  function formatDate(date) { if (!date) return ""; try { return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "dd/MM/yyyy"); } catch (e) { return ""; } }
  function safeValue(val) { return val || ""; }

  // --- BULK UPDATE TOOL (Cập nhật cả cột G và H) ---
  /**
   * HÀM 1: LẤY DANH SÁCH MÓN CHA ĐANG SỬ DỤNG NVL (Sửa lỗi hiển thị số 0)
   * @param {string} ingredientCode - Mã nguyên liệu (Con)
   */
  function getParentItemsUsingIngredient(ingredientCode) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("BOM"); // ⚠️ CHÚ Ý: Đổi tên Sheet cho đúng file của cháu
    var data = sheet.getDataRange().getValues();
    
    var result = [];
    
    // Giả định cấu trúc cột BOM (Cháu đếm lại cột trong file Excel nhé A=0, B=1...)
    // Ví dụ: A=Mã Cha, B=Tên Cha, C=Mã Con, D=Tên Con, E=Định Lượng, F=ĐVT
    const COL_PARENT_CODE = 0; 
    const COL_PARENT_NAME = 1;
    const COL_CHILD_CODE = 2;
    const COL_QTY = 4; // ⚠️ QUAN TRỌNG: Kiểm tra lại cột E (Định lượng) có đúng là index 4 không?
    const COL_UNIT = 5;

    for (var i = 1; i < data.length; i++) { // Bỏ qua header
      var row = data[i];
      // So sánh Mã Con, chuyển về String để tránh lỗi số/chữ
      if (String(row[COL_CHILD_CODE]) === String(ingredientCode)) {
        result.push({
          parentCode: row[COL_PARENT_CODE],
          parentName: row[COL_PARENT_NAME],
          // SỬA LỖI Ở ĐÂY: Đảm bảo parse số, nếu lỗi thì về 0
          qty: parseFloat(row[COL_QTY]) || 0, 
          unit: row[COL_UNIT]
        });
      }
    }
    
    return result;
  }

  /**
   * HÀM 2: TÌM VÀ THAY THẾ NVL TRONG BOM (Sửa lỗi nút không chạy)
   * @param {string} oldCode - Mã cũ cần thay
   * @param {string} newCode - Mã mới thay thế vào
   */
  function replaceIngredientInBom(oldCode, newCode) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("BOM"); // ⚠️ CHÚ Ý: Đổi tên Sheet
    var range = sheet.getDataRange();
    var data = range.getValues();
    
    const COL_CHILD_CODE = 2; // Cột chứa Mã Nguyên Liệu (Con) - Index 2 = Cột C
    var changeCount = 0;

    // Quét và thay thế trong mảng (Memory) cho nhanh
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][COL_CHILD_CODE]) === String(oldCode)) {
        data[i][COL_CHILD_CODE] = newCode; // Thay thế mã
        changeCount++;
      }
    }

    // Ghi ngược lại xuống Sheet (Chỉ ghi nếu có thay đổi)
    if (changeCount > 0) {
      range.setValues(data);
      return { success: true, message: "Đã thay thế thành công " + changeCount + " dòng!" };
    } else {
      return { success: false, message: "Không tìm thấy mã cũ " + oldCode + " trong BOM." };
    }
  }

  function importSalesHistory(payload) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName("Data Doanh Thu");
      if (!sheet) {
          sheet = ss.insertSheet("Data Doanh Thu");
          sheet.appendRow(["Ngày Import", "Khoảng Thời Gian", "Chi Nhánh", "Mã SP", "Tên SP", "Tổng SL", "SL Delivery", "SL Dine-in", "SL Take-away", "Người Nhập"]);
      }
      
      const importDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      const period = payload.period || "";
      let userEmail = "";
      try { userEmail = Session.getActiveUser().getEmail(); } catch(e) { userEmail = "Unknown User"; }
      
      const rows = [];
      // Hàm ép kiểu an toàn
      const safeNum = (n) => typeof n === 'number' ? n : (Number(n) || 0);

      payload.items.forEach(item => {
          rows.push([
              importDate,
              period,
              item.store || "",
              String(item.code || ""), // Ép về String để giữ số 0 đầu
              item.name || "",
              safeNum(item.qty),
              safeNum(item.qtyDeli),
              safeNum(item.qtyDine),
              safeNum(item.qtyTake),
              userEmail
          ]);
      });

      if (rows.length > 0) {
          sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
      
      return { success: true, message: "Đã lưu thành công " + rows.length + " dòng!" };
    } catch (e) {
      return { success: false, message: "Lỗi Backend: " + e.toString() };
    }
  }

  /**
   * HÀM MỚI: Thêm nhanh danh sách Product mới từ file Import Doanh Thu
   * Tự động map các cột: Code, Name, Category, Class, Price
   */
  function quickAddMissingProducts(items) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PROD); // Danh mục Product
    if (!sheet) return { success: false, message: "Không tìm thấy Sheet Danh mục Product!" };

    try {
      const rowsToAdd = [];
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

      items.forEach(item => {
        // Cấu trúc cột Sheet Product: 
        // A: Code | B: Name | C: Category | D: Class | E: Team | F: SAP Code | G: Cost | H: Status
        rowsToAdd.push([
          String(item.code),       // A: Code (Dùng làm mã nội bộ luôn)
          item.name,               // B: Tên
          item.category || "",     // C: Category (Lấy từ file Excel)
          item.class || "",        // D: Class (Lấy từ file Excel)
          "Service",               // E: Team (Mặc định Service vì bán hàng)
          String(item.code),       // F: SAP Code
          0,                       // G: Giá vốn (Tạm để 0)
          "Active"                 // H: Trạng thái
        ]);
      });

      if (rowsToAdd.length > 0) {
        // Tìm dòng trống tiếp theo
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      }

      return { success: true, message: `Đã thêm thành công ${rowsToAdd.length} mã mới!` };
    } catch (e) {
      return { success: false, message: "Lỗi Backend: " + e.toString() };
    }
  }

  // [SOURCE: KAIZEN BRAIN SYSTEM]
// ==========================================================

// ⚠️ THAY ID FILE TXT CỦA CHÁU VÀO ĐÂY
const KAIZEN_CONFIG = {
  BRAIN_FILE_ID: "14K3qOvEtsLfmo_XJfV2yxbYJBknulgmL" 
};

/**
 * HÀM 1: ĐỌC TRI THỨC (Dùng để lấy Context ném cho AI đầu buổi)
 */
function getKaizenBrain() {
  try {
    const file = DriveApp.getFileById(KAIZEN_CONFIG.BRAIN_FILE_ID);
    return { 
      success: true, 
      content: file.getBlob().getDataAsString() 
    };
  } catch (e) {
    return { success: false, message: "Lỗi đọc não: " + e.toString() };
  }
}

// Hàm này chỉ dùng để Cháu kiểm tra Log thôi nhé
function debugReadBrain() {
  const data = getKaizenBrain();
  console.log("=== KẾT QUẢ ĐỌC NÃO ===");
  console.log(data); 
  console.log("=======================");
}

/**
 * HÀM 2: NẠP TRI THỨC MỚI (AI gọi hàm này qua User)
 * @param {string} newRuleContent - Nội dung quy tắc mới
 */
function appendKaizenRule(newRuleContent) {
  try {
    const file = DriveApp.getFileById(KAIZEN_CONFIG.BRAIN_FILE_ID);
    const currentContent = file.getBlob().getDataAsString();
    
    // Timestamp định dạng VN
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    // Tạo Block nội dung mới, có phân cách rõ ràng
    const updateBlock = `\n\n# [UPDATE OTA: ${time}] ---------------------------\n${newRuleContent}`;
    
    // Ghi đè: Nội dung cũ + Block mới
    file.setContent(currentContent + updateBlock);
    
    return { success: true, message: `Đã nạp tri thức lúc ${time}!` };
  } catch (e) {
    return { success: false, message: "Lỗi nạp tri thức: " + e.toString() };
  }
}

/**
 * HÀM 3: TÁI CẤU TRÚC NÃO BỘ (Dùng khi Clean dọn dẹp)
 * Hàm này sẽ XÓA SẠCH cũ và GHI MỚI toàn bộ.
 */
function rewriteKaizenBrain(fullContent) {
  try {
    const file = DriveApp.getFileById(KAIZEN_CONFIG.BRAIN_FILE_ID);
    
    // Ghi đè toàn bộ nội dung (setContent thay vì append)
    file.setContent(fullContent);
    
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    return { success: true, message: `Đã tái cấu trúc não bộ thành công lúc ${time}!` };
  } catch (e) {
    return { success: false, message: "Lỗi ghi đè não: " + e.toString() };
  }
}

function rebuildBOMCache(productID) {
  // 1. Lấy dữ liệu BOM gốc (Dạng cây)
  let bomTree = getBOMTree(productID); 

  // 2. Biến chứa kết quả phẳng
  let flatList = [];

  // 3. Hàm đệ quy để đào sâu và làm phẳng (CHÚ IT ĐÃ ĐỘ LẠI)
  function flatten(node, currentYield, pathString) {
    
    // Kiểm tra an toàn: nếu node không có children thì dừng
    if (!node.children || node.children.length === 0) return;

    // Duyệt qua từng thành phần con
    node.children.forEach(child => {
      
      // Tính Yield tích lũy
      let accumulatedYield = currentYield * (child.inputQty / child.outputQty);
      
      // --- [CHÚ IT FIX HIỂN THỊ TẠI ĐÂY] ---
      
      // A. Lấy ĐVT (Nếu hàm getBOMTree chưa trả về unit, cháu nhớ kiểm tra lại hàm đó)
      let unit = child.unit || ""; 
      
      // B. Làm tròn số lượng cho gọn (3 số lẻ), đổi dấu chấm thành phẩy cho chuẩn VN
      let qtyPretty = (Math.round(child.usage * 1000) / 1000).toString().replace('.', ',');

      // C. Tạo chuỗi hiển thị bước hiện tại: Ví dụ "(50 Gr) Phô mai"
      let stepInfo = `[${qtyPretty} ${unit}] ${child.name}`;

      // D. Nối chuỗi: Dùng dấu mũi tên đậm "➔" thay vì dấu ">" nhìn cho xịn
      let newPath = pathString + " ➔ " + stepInfo;
      
      // -------------------------------------

      if (child.type === 'RAW') {
        // ĐIỂM DỪNG: Nếu là Raw, ghi vào danh sách kết quả
        flatList.push({
          product_id: productID,
          raw_id: child.id,
          total_qty: child.usage * accumulatedYield, // CON SỐ VÀNG
          path_log: newPath, // Đường dẫn đẹp đã được tạo ở trên
          updated: new Date()
        });
      } else if (child.type === 'SEMI') {
        // ĐỆ QUY: Nếu là Semi, đào tiếp với đường dẫn mới
        flatten(child, accumulatedYield, newPath);
      }
    });
  }
  
  // 4. Bắt đầu chạy
  // Chuỗi khởi đầu: Tên món chính (Ví dụ: "Pizza Hải Sản")
  if (bomTree) {
      flatten(bomTree, 1.0, `(1) ${bomTree.name}`);
  }
  
  // 5. Lưu flatList vào Sheet 'DB_BOM_CACHE'
  // Chú ý: Hàm saveToSheet của cháu cần xử lý xóa dữ liệu cũ của productID này trước khi ghi mới
  saveToSheet('DB_BOM_CACHE', flatList);
}

/**
 * ------------------------------------------------------------------------
 * [CORE] HÀM TÁI TẠO CACHE BOM (ĐÃ FIX LỖI 1.2 TỶ)
 * Chạy hàm này từ menu "SYSTEM ADMIN" để làm sạch dữ liệu.
 * ------------------------------------------------------------------------
 */
function regenerateAllBOMCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCache = ss.getSheetByName("DB_BOM_CACHE");
  const sheetProd = ss.getSheetByName("Danh mục Product");
  
  if (!sheetCache || !sheetProd) {
    SpreadsheetApp.getUi().alert("❌ Thiếu Sheet Cache hoặc Danh mục Product!");
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("⏳ Đang tải bản đồ công thức & Tên hàng...", "System Admin");

  // 1. Load Dữ liệu
  const allRecipes = getRecipeMap(ss); 
  const nameMap = createNameMap(ss); // <--- [MỚI] Lấy từ điển tên
  
  const prodData = sheetProd.getDataRange().getValues();
  let cacheData = [];
  const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  SpreadsheetApp.getActiveSpreadsheet().toast("🚀 Đang tính toán lộ trình (Trace Path)...", "System Admin");

  // 2. Tính toán
  for (let i = 1; i < prodData.length; i++) {
    let prodCode = String(prodData[i][0]).trim();
    if (!prodCode) continue;

    // Truyền nameMap vào hàm bung BOM
    let bomResult = explodeBOMForCache(prodCode, 1, allRecipes, nameMap); 
    
    for (let rawCode in bomResult) {
      let item = bomResult[rawCode];
      cacheData.push([
        prodCode,           
        rawCode,            
        item.qty,           
        item.path, // Lúc này path đã là Tên -> Tên -> Tên         
        timeStr             
      ]);
    }
  }

  // 3. Ghi kết quả
  if (cacheData.length > 0) {
    sheetCache.getRange("A2:E").clearContent();
    sheetCache.getRange(2, 1, cacheData.length, 5).setValues(cacheData);
    SpreadsheetApp.getUi().alert(`✅ Đã cập nhật xong!\nKiểm tra cột Trace_Path xem đã hiện Tên chưa nhé.`);
  } else {
    SpreadsheetApp.getUi().alert("⚠️ Không có dữ liệu để tính.");
  }
}

/**
 * [KAIZEN V3] HÀM BUNG BOM AN TOÀN (CHỐNG VÒNG LẶP & TRÀN BỘ NHỚ)
 * Update: Thêm cơ chế "Visited Stack" để phát hiện vòng lặp A -> B -> A
 */
function explodeBOMForCache(rootCode, demandQty, allRecipes, nameMap) {
  let results = {}; 

  const getName = (code) => nameMap[code] || code;
  // Format số: 1,234.567 (Bỏ bớt số 0 thừa nếu cần)
  const fmt = (num) => Number(num).toLocaleString('vi-VN', {maximumFractionDigits: 3});

  let rootName = getName(rootCode);
  
  // Hàm đệ quy
  function traverse(currentCode, currentQty, history, visited = []) {
    if (visited.includes(currentCode)) return; // Chống loop

    let recipe = allRecipes[currentCode];
    let currentName = getName(currentCode);

    // Cập nhật lịch sử
    let currentStep = { 
      name: currentName, 
      qty: currentQty,
      isSemi: (recipe && recipe.components && recipe.components.length > 0)
    };
    let newHistory = [...history, currentStep];

    // ĐIỂM CUỐI (NVL hoặc Semi cụt)
    if (!currentStep.isSemi) {
      if (!results[currentCode]) {
        results[currentCode] = { qty: 0, branches: [] };
      }
      
      results[currentCode].qty += currentQty;
      
      // [FIX] TẠO DÒNG NHÁNH (Số lượng đứng trước)
      let branchStr = newHistory.slice(1).map((step, index) => {
        let prefix = (index === 0) ? "   └─ " : " ➔ ";
        // KAIZEN: (SL) Tên
        return `${prefix}(${fmt(step.qty)}) ${step.name}`;
      }).join("");

      if (!results[currentCode].branches.includes(branchStr)) {
        results[currentCode].branches.push(branchStr);
      }
      return;
    }

    // BUNG TIẾP
    let batchSize = recipe.batchOutput; 
    if (!batchSize || batchSize <= 0) batchSize = 1;
    let ratio = currentQty / batchSize;
    let newVisited = [...visited, currentCode];

    recipe.components.forEach(comp => {
      let childNeed = comp.qty * ratio; 
      traverse(comp.code, childNeed, newHistory, newVisited);
    });
  }

  // Bắt đầu chạy
  traverse(String(rootCode).trim(), demandQty, []);
  
  // TỔNG HỢP KẾT QUẢ
  let finalResults = {};
  for (let code in results) {
    let item = results[code];
    
    // [FIX] Dòng tiêu đề cũng đưa số lượng lên trước
    let treeView = `BẮT ĐẦU: (${fmt(demandQty)}) ${rootName}\n` + item.branches.join('\n');
    
    finalResults[code] = {
      qty: item.qty,
      path: treeView
    };
  }

  return finalResults;
}


// ==========================================
// KHU VỰC 9: TRACEABILITY & SHADOW ENGINE
// ==========================================

/**
 * [V12.31 DUAL TRACE] HÀM TRUY VẾT THÔNG MINH 2 CHIỀU
 * - Input: Mã bất kỳ (NVL hoặc Product)
 * - Logic: 
 * + Nếu là NVL -> Tìm xem nó nằm trong Món nào? (Where Used)
 * + Nếu là Product -> Tìm xem nó gồm những NVL nào? (BOM Explosion)
 */
function getTraceDataFromCache(code) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = ss.getSheetByName(CONFIG.SHEET_BOM_CACHE);
    
    if (!sheetCache) return { success: false, message: "Chưa có Cache! Hãy chạy Tái tạo." };
    
    const data = sheetCache.getDataRange().getValues();
    // Cache Structure: [0]Product | [1]Raw | [2]Qty | [3]Path | [4]Time
    
    const target = String(code).trim();
    let results = [];
    let mode = ""; // Để ghi log xem đang tra xuôi hay ngược

    // 1. CHUẨN BỊ MAP TÊN (Để hiển thị cho đẹp)
    const masterMap = getMasterMap(ss);

    // 2. QUÉT CACHE
    for (let i = 1; i < data.length; i++) {
      let prodCode = String(data[i][0]).trim();
      let rawCode = String(data[i][1]).trim();
      let qty = Number(data[i][2]) || 0;
      let path = String(data[i][3]);
      let time = data[i][4];

      // TRƯỜNG HỢP A: Bấm vào NVL (Tìm Món ăn sử dụng)
      if (rawCode === target) {
        mode = "WHERE_USED";
        let prodInfo = masterMap[prodCode] || { name: prodCode };
        
        // Lấy ĐVT của cái NVL đang soi (Target)
        let rawUnit = masterMap[rawCode] ? masterMap[rawCode].unit : ""; 

        results.push({
          code: prodCode,             
          name: prodInfo.name,        
          qty: qty,                   
          unit: rawUnit,              // <--- THÊM DÒNG NÀY
          path: path,                 
          type: "Dùng trong"
        });
      }
      
      // TRƯỜNG HỢP B: Bấm vào Product (Tìm thành phần NVL)
      else if (prodCode === target) {
        mode = "BOM_EXPLODE";
        let rawInfo = masterMap[rawCode] || { name: rawCode };
        
        // Lấy ĐVT của từng thành phần con
        let childUnit = rawInfo.unit || "";

        results.push({
          code: rawCode,              
          name: rawInfo.name,         
          qty: qty,                   
          unit: childUnit,            // <--- THÊM DÒNG NÀY
          path: path,                 
          type: "Thành phần"
        });
      }
    }
    
    // Sắp xếp: Cái nào số lượng lớn lên đầu
    results.sort((a, b) => b.qty - a.qty);

    if (results.length === 0) {
      return { success: false, message: `Không tìm thấy dữ liệu Trace cho mã ${target} trong Cache.` };
    }

    return { success: true, data: results, mode: mode };
    
  } catch (e) {
    return { success: false, message: "Lỗi Trace: " + e.toString() };
  }
}

/**
 * [SHADOW ENGINE V12.29 - FULL AUDIT MODE]
 * Tính toán tổng nhu cầu NVL dựa trên Doanh thu & Cache
 */
function runShadowAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSales = ss.getSheetByName("Data Doanh Thu");
  const sheetCache = ss.getSheetByName(CONFIG.SHEET_BOM_CACHE);
  let sheetResult = ss.getSheetByName("SHADOW_RESULT");

  if (!sheetSales || !sheetCache) return;

  // 1. TẠO SHEET KẾT QUẢ
  if (!sheetResult) {
    sheetResult = ss.insertSheet("SHADOW_RESULT");
    sheetResult.appendRow(["Mã NVL", "Tên NVL (Gợi ý)", "Tổng Cần (Shadow)", "ĐVT", "Chi tiết truy vết (Full Source)"]);
    sheetResult.getRange("A1:E1").setFontWeight("bold").setBackground("#d9ead3");
    sheetResult.setColumnWidth(2, 250);
    sheetResult.setColumnWidth(5, 400);
    sheetResult.setRowHeight(1, 40);
  } else {
    if(sheetResult.getLastRow() > 1) 
      sheetResult.getRange(2, 1, sheetResult.getLastRow()-1, 5).clearContent();
  }

  // 2. GỘP DOANH THU
  const salesData = sheetSales.getDataRange().getValues();
  let salesMap = {}; 
  for (let i = 1; i < salesData.length; i++) {
    let code = String(salesData[i][3]).trim(); // Col D
    let qty = Number(salesData[i][5]) || 0;    // Col F
    if (code && qty > 0) salesMap[code] = (salesMap[code] || 0) + qty;
  }

  // 3. TÍNH TOÁN
  const cacheData = sheetCache.getDataRange().getValues();
  let shadowResult = {}; 

  for (let i = 1; i < cacheData.length; i++) {
    let prodCode = String(cacheData[i][0]).trim();
    let rawCode = String(cacheData[i][1]).trim();
    let unitUsage = Number(cacheData[i][2]) || 0;
    
    if (salesMap[prodCode]) {
      let salesQty = salesMap[prodCode];
      let totalRawNeed = salesQty * unitUsage;

      if (!shadowResult[rawCode]) shadowResult[rawCode] = { qty: 0, sources: [] };
      
      shadowResult[rawCode].qty += totalRawNeed;
      shadowResult[rawCode].sources.push({
        prod: prodCode,
        sales: salesQty,
        total: totalRawNeed
      });
    }
  }

  // 4. LẤY THÔNG TIN MASTER
  let masterInfo = getMasterMap(ss); 

  // 5. XUẤT KẾT QUẢ
  let outputRows = [];
  
  for (let raw in shadowResult) {
    let item = shadowResult[raw];
    let info = masterInfo[raw] || { name: "Unknown", unit: "" };
    
    item.sources.sort((a, b) => b.total - a.total);
    
    let traceInfo = item.sources.map(src => {
      return `[${src.prod}] (x${src.sales}) ➔ ${Math.round(src.total)}`;
    }).join("\n");
    
    outputRows.push([
      raw,                                
      info.name,                          
      Math.round(item.qty * 1000) / 1000, 
      info.unit,                          
      traceInfo                           
    ]);
  }

  if (outputRows.length > 0) {
    outputRows.sort((a,b) => String(a[0]).localeCompare(String(b[0])));
    sheetResult.getRange(2, 1, outputRows.length, 5).setValues(outputRows);
    sheetResult.getRange(2, 5, outputRows.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheetResult.getRange(2, 1, outputRows.length, 5).setVerticalAlignment("top");
    SpreadsheetApp.getUi().alert(`✅ ĐÃ CHẠY AUDIT!\nTổng mã NVL: ${outputRows.length}`);
  } else {
    SpreadsheetApp.getUi().alert("⚠️ Không có số liệu tính toán.");
  }
}

/**
 * [HELPER] MAP TÊN & ĐVT
 */
function getMasterMap(ss) {
  let map = {};
  let sNVL = ss.getSheetByName(CONFIG.SHEET_NVL);
  if (sNVL) {
    let d = sNVL.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      let code = String(d[i][1]).trim();
      let name = String(d[i][0]).trim();
      let unit = String(d[i][2]).trim();
      if (code) map[code] = { name: name, unit: unit };
    }
  }
  let sProd = ss.getSheetByName(CONFIG.SHEET_PROD);
  if (sProd) {
    let d = sProd.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      let code = String(d[i][0]).trim();
      let name = String(d[i][1]).trim();
      if (code && !map[code]) map[code] = { name: name, unit: "Cái" };
    }
  }
  return map;
}

/**
 * [KAIZEN V2] HÀM ĐỌC CÔNG THỨC TOÀN DIỆN (BAO GỒM CẢ PRODUCT & SEMI)
 * Update: Đã thêm phần đọc sheet 'BOM Product' để liên kết Món ăn -> Semi
 */
function getRecipeMap(ss) {
  let recipes = {};

  const clean = (val) => String(val || "").trim();
  const safeNum = (val) => {
      if (val === "" || val === null || val === undefined) return 0;
      if (typeof val === 'number') return val;
      let cleanStr = String(val).replace(/,/g, '.').trim(); 
      let res = Number(cleanStr);
      return isNaN(res) ? 0 : res;
  };

  // --- PHẦN 1: ĐỌC CÔNG THỨC SEMI (3 Sheet Bếp) ---
  const semiSheets = [CONFIG.SHEET_KIT_KITCHEN, CONFIG.SHEET_KIT_PIZZA, CONFIG.SHEET_KIT_SERVICE];
  
  semiSheets.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const data = s.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let r = data[i];
      let parentCode = clean(r[0]);         // Cột A: Mã Cha
      let type = clean(r[1]).toLowerCase(); // Cột B: Loại
      let val = safeNum(r[4]);              // Cột E: SL/Yield

      if (!parentCode) continue;
      if (!recipes[parentCode]) recipes[parentCode] = { batchOutput: 1, components: [] };

      if (type === 'parent') {
        // Lấy sản lượng mẻ nấu (Batch Output)
        recipes[parentCode].batchOutput = (val > 0) ? val : 1;
      } 
      else if (type === 'child') {
        let childCode = clean(r[2]); // Cột C: Mã Con
        if (childCode) {
          recipes[parentCode].components.push({ code: childCode, qty: val });
        }
      }
    }
  });

  // --- PHẦN 2: ĐỌC CÔNG THỨC PRODUCT (QUAN TRỌNG: Món -> Semi) ---
  // Chú thêm đoạn này để hệ thống hiểu 1 Pizza gồm những gì
  const sheetBOM = ss.getSheetByName(CONFIG.SHEET_BOM_PRODUCT); // "BOM Product"
  if (sheetBOM) {
    const bomData = sheetBOM.getDataRange().getValues();
    // Cấu trúc file BOM Product của cháu:
    // A: Mã Parent (10000014) | B: Loại | C: Mã Component (214) | E: SL (80)
    
    for (let i = 1; i < bomData.length; i++) {
      let r = bomData[i];
      let parentCode = clean(r[0]); // Cột A
      let childCode = clean(r[2]);  // Cột C
      let qty = safeNum(r[4]);      // Cột E

      if (!parentCode || !childCode) continue;

      // Nếu chưa có trong danh sách công thức thì tạo mới
      // Với Product, Batch Output mặc định luôn là 1 (1 cái bánh)
      if (!recipes[parentCode]) {
        recipes[parentCode] = { batchOutput: 1, components: [] };
      }

      recipes[parentCode].components.push({
        code: childCode,
        qty: qty
      });
    }
  }

  return recipes;
}

/**
 * [KAIZEN V12.5 - FINAL TRACE ENGINE]
 * Khắc phục triệt để lỗi đứt gãy dữ liệu BOM
 */
  function buildPath(code, currentQty, currentLevel, currentPathNodes) {
    // a. Chặn vòng lặp vô tận & Giới hạn độ sâu (Max 10 tầng)
    if (currentLevel > 10) {
       allPaths.push({ nodes: [...currentPathNodes, { qty: 0, name: "⚠️ LỖI: QUÁ NHIỀU TẦNG (LOOP)", level: currentLevel, type: "ERROR" }].reverse() });
       return;
    }

    const cleanCode = String(code).trim();
    const itemName = nameMap[cleanCode] || "Mã " + cleanCode;
    // Xác định loại hàng từ Master Data (Chính xác hơn chỉ dựa vào việc có BOM hay không)
    const masterType = typeMap.get(cleanCode) || "NVL"; 
    const hasRecipe = bomMap[cleanCode] ? true : false;
    
    // b. Format số lượng chuẩn (Luôn giữ 3 số lẻ để khớp với Frontend)
    const qtyStr = Number(currentQty).toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 3});

    const newNode = {
  // Lỗi #2: Chuyển 'en-US' sang 'vi-VN' để dùng dấu chấm phân cách hàng nghìn
  qty: Number(currentQty).toLocaleString('vi-VN', {minimumFractionDigits: 0, maximumFractionDigits: 3}), 
  
  // Lỗi #3: Đổi raw_qty thành qty_raw theo đúng quy tắc Suffix "Dual State"
  qty_raw: Number(currentQty), 
  
  name: itemName,
  unit: item.unit || "Gr",
  level: currentLevel
};

    const newPath = [...currentPathNodes, newNode];

    // c. LOGIC QUYẾT ĐỊNH (CORE)
    if (hasRecipe) {
      // TRƯỜNG HỢP 1: Có công thức -> Đào tiếp
      bomMap[cleanCode].components.forEach(item => {
        let childQty = (currentQty * item.qty) / (bomMap[cleanCode].batchOutput || 1);
        buildPath(item.code, childQty, currentLevel + 1, newPath);
      });
    } else {
      // TRƯỜNG HỢP 2: Không có công thức
      if (masterType === 'SEMI' || masterType === 'PROD') {
          // ⚠️ RỦI RO PHÁT HIỆN: Là Semi/Prod mà không có BOM -> ĐỨT GÃY!
          // Ghi nhận dòng lỗi để Frontend hiển thị màu Đỏ
          const errorNode = {
             qty: "MISSING",
             name: `⚠️ CẢNH BÁO: ${itemName} (Chưa khai báo BOM)`,
             level: currentLevel + 1,
             type: "BROKEN" 
          };
          allPaths.push({ nodes: [...newPath, errorNode].reverse() });
      } else {
          // TRƯỜNG HỢP 3: Là NVL thật sự -> Điểm cuối an toàn
          allPaths.push({ nodes: [...newPath].reverse() });
      }
    }
  }

/**
 * HÀM 1: TẠO TỪ ĐIỂN TÊN (Name Mapping)
 * Giúp đổi Mã (100xxx) -> Tên món ăn/nguyên liệu
 */
function createNameMap(ss) {
  let map = {};
  const clean = (val) => String(val || "").trim();

  // 1. Quét Danh mục NVL
  const sheetNVL = ss.getSheetByName(CONFIG.SHEET_NVL);
  if (sheetNVL) {
    const data = sheetNVL.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let code = clean(data[i][1]);
      if (code) map[code] = clean(data[i][0]); 
    }
  }

  // 2. Quét Danh mục Product
  const sheetProd = ss.getSheetByName(CONFIG.SHEET_PROD);
  if (sheetProd) {
    const data = sheetProd.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let code = clean(data[i][0]);
      if (code) map[code] = clean(data[i][1]);
    }
  }

  // 3. Quét 3 Sheet Semi (Kitchen, Pizza, Service)
  const semiSheets = [CONFIG.SHEET_KIT_KITCHEN, CONFIG.SHEET_KIT_PIZZA, CONFIG.SHEET_KIT_SERVICE];
  semiSheets.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const data = s.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (clean(data[i][1]).toLowerCase() === 'parent') {
        let code = clean(data[i][0]);
        let nameItem = clean(data[i][3]);
        if (code) map[code] = nameItem;
      }
    }
  });
  return map;
}



function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function roundNumber(v, d) { return Number(Math.round(v + "e" + d) + "e-" + d) || 0; }

/* [MODULE CACHE] CẬP NHẬT CACHE RIÊNG LẺ (PARTIAL UPDATE) */
function updateSingleBOMCache(targetCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = ss.getSheetByName("DB_BOM_CACHE");
    if (!sheetCache) return;

    // 1. Lấy cấu trúc cây mới nhất của mã hàng này
    const newTreeData = getTraceDataFromCache(targetCode); // Tận dụng hàm phân tích cây có sẵn
    if (!newTreeData || !newTreeData.length) return;

    // 2. Tìm vị trí cũ trong Sheet Cache để ghi đè
    const dataCache = sheetCache.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 0; i < dataCache.length; i++) {
      if (dataCache[i][0] === targetCode) {
        rowIndex = i + 1;
        break;
      }
    }

    const rowData = [targetCode, JSON.stringify(newTreeData), new Date()];

    // 3. Ghi đè nếu đã có, hoặc thêm mới nếu chưa tồn tại
    if (rowIndex > 0) {
      sheetCache.getRange(rowIndex, 1, 1, 3).setValues([rowData]);
    } else {
      sheetCache.appendRow(rowData);
    }
    console.log("Đã cập nhật Cache cho mã: " + targetCode);
  } catch (e) {
    console.error("Lỗi UpdateSingleCache: " + e.toString());
  }
}

function bulkUpdateSemiCosts(){try{const ss=SpreadsheetApp.getActiveSpreadsheet(),sN=ss.getSheetByName(CONFIG.SHEET_NVL),sP=ss.getSheetByName(CONFIG.SHEET_PROD),sBp=ss.getSheetByName(CONFIG.SHEET_BOM_PRODUCT),costs=new Map();if(!sN||!sP||!sBp)return{success:false,message:"Thiếu Sheet Danh mục/BOM!"};sN.getDataRange().getValues().forEach((r,i)=>{if(i>0)costs.set(String(r[1]).trim(),Number(r[3])||0)});const semiSheets=[CONFIG.SHEET_KIT_KITCHEN,CONFIG.SHEET_KIT_PIZZA,CONFIG.SHEET_KIT_SERVICE];semiSheets.forEach(name=>{const s=ss.getSheetByName(name);if(!s)return;const d=s.getDataRange().getValues();let pIdx=-1,pTot=0,pY=1;for(let i=0;i<d.length;i++){const type=String(d[i][1]).trim(),code=String(d[i][2]).trim();if(type==='Parent'){if(pIdx!==-1){d[pIdx][7]=Math.round(pTot);d[pIdx][6]=Number((pY>0?pTot/pY:0).toFixed(3));costs.set(String(d[pIdx][0]).trim(),d[pIdx][6])}pIdx=i;pTot=0;pY=Number(d[i][4])||1}else if(type==='Child'){const q=Number(d[i][4])||0,uP=costs.get(code)||0,lT=uP*q;d[i][6]=lT;pTot+=lT}}if(pIdx!==-1){d[pIdx][7]=Math.round(pTot);d[pIdx][6]=Number((pY>0?pTot/pY:0).toFixed(3));costs.set(String(d[pIdx][0]).trim(),d[pIdx][6])}s.getDataRange().setValues(d)});const bD=sBp.getDataRange().getValues(),pCosts=new Map();let pC="",pT=0;for(let i=1;i<bD.length;i++){const r=bD[i],type=String(r[1]).toLowerCase();if(type==='parent'){if(pC)pCosts.set(pC,pT);pC=String(r[0]).trim();pT=0}else if(type==='child'){pT+=(costs.get(String(r[2]).trim())||0)*(Number(r[4])||0)}}if(pC)pCosts.set(pC,pT);const pData=sP.getDataRange().getValues();for(let i=1;i<pData.length;i++){const c=String(pData[i][0]).trim();if(pCosts.has(c))pData[i][6]=Math.round(pCosts.get(c))}sP.getDataRange().setValues(pData);regenerateAllBOMCache();return{success:true,message:"✅ Domino thành công: Giá đã cập nhật từ Gốc đến Ngọn!"}}catch(e){return{success:false,message:"Lỗi: "+e.toString()}}}

function bomEngine(targetCode,targetQty,rMap,mMap,level=0,visited=[]){if(level>10||visited.includes(targetCode))return[];const recipe=rMap[targetCode];if(!recipe||!recipe.components||recipe.components.length===0){const mas=mMap.get(targetCode);return[{code:targetCode,name:mas?mas.name:targetCode,qty_raw:targetQty,qty_fmt:smartRound(targetQty).toString(),unit:mas?mas.unit:'',isLeaf:true}]}let results=[];const ratio=targetQty/(recipe.batchOutput||1);recipe.components.forEach(comp=>{const childQty=comp.qty*ratio;const subResults=bomEngine(comp.code,childQty,rMap,mMap,level+1,[...visited,targetCode]);results=results.concat(subResults)});return results}

function getSpoilageHistory(fromDateStr, toDateStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parseDateVN = (dStr) => {
      if (!dStr || dStr instanceof Date) return dStr;
      const p = String(dStr).split('/');
      return p.length === 3 ? new Date(p[2], p[1] - 1, p[0]) : null;
    };
    const fDate = parseDateVN(fromDateStr), tDate = parseDateVN(toDateStr);
    if (fDate) fDate.setHours(0,0,0,0); if (tDate) tDate.setHours(23,59,59,999);
    const results = [];

    // 1. Quét Sheet Hủy NVL & SEMI 
    const sSp = ss.getSheetByName(CONFIG.SHEET_SPOILAGE);
    if (sSp && sSp.getLastRow() > 1) {
      sSp.getRange(2, 1, sSp.getLastRow() - 1, 15).getValues().forEach(r => {
        const d = parseDateVN(r[0]);
        if (!d || (fDate && d < fDate) || (tDate && d > tDate)) return;
        results.push({
          date: r[0] instanceof Date ? formatDate(r[0]) : r[0],
          code: String(r[1]), name: String(r[2]), qty: Number(r[8]) || 0,
          unit: String(r[7]), note: String(r[10]), dept: String(r[11]),
          cost: Number(r[13]) || 0, amount: Number(r[14]) || 0,
          type: String(r[10]).includes("BTP") ? "SEMI" : "NVL"
        });
      });
    }

    // 2. Quét Sheet Hủy PRODUCT 
    const sPr = ss.getSheetByName(CONFIG.SHEET_SPOILAGE_PROD);
    if (sPr && sPr.getLastRow() > 1) {
      sPr.getRange(2, 1, sPr.getLastRow() - 1, 12).getValues().forEach(r => {
        const d = parseDateVN(r[0]);
        if (!d || (fDate && d < fDate) || (tDate && d > tDate)) return;
        const q = Number(r[4]), c = Number(r[11]) || 0;
        results.push({
          date: r[0] instanceof Date ? formatDate(r[0]) : r[0],
          code: String(r[2]), name: String(r[1]), qty: q,
          unit: 'Cái', note: String(r[5]), dept: String(r[6]),
          cost: c, amount: Math.round(q * c), type: 'PROD'
        });
      });
    }
    return { success: true, data: results.sort((a,b) => parseDateVN(b.date) - parseDateVN(a.date)) };
  } catch (e) { return { success: false, message: "Lỗi: " + e.toString() }; }
}
// ============ exportSpoilageToExcel ============
// Xuất báo cáo hủy hàng ra Excel - Lưu vào folder cố định
// Layout: Cột A-D = Hợp lệ | Cột F+ = Lỗi cần kiểm tra
function exportSpoilageToExcel(data, config) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shouldExplodeBOM = config.explodeBOM === true;
    const unitMode = config.unitMode || 'ORIGINAL';
    
    // FOLDER ID cố định để lưu file
    const FOLDER_ID = '1OAiOKp4HbLIXDKzlo4IkOrGFX2wjCvuM';
    
    // 1. TẠO MAP MASTER DATA
    const masterMap = new Map();
    const sheetNVL = ss.getSheetByName(CONFIG.SHEET_NVL);
    if (sheetNVL) {
      sheetNVL.getDataRange().getValues().slice(1).forEach(r => {
        const code = String(r[1]).trim();
        if (code) {
          masterMap.set(code, {
            name: String(r[0]),
            unit: String(r[2]),
            cost: Number(r[3]) || 0,
            stdUnit: String(r[4]),
            rate: Number(r[5]) || 1
          });
        }
      });
    }
    
    // 2. XỬ LÝ DỮ LIỆU - CHIA 2 NHÓM
    const validRows = [];   // Cột A-D: Hợp lệ
    const errorRows = [];   // Cột F+: Lỗi
    
    data.forEach(item => {
      const needExplode = item.needExplode === true;
      const itemType = item.type || 'NVL';
      
      // TRƯỜNG HỢP 1: Cần bung BOM (SEMI/PROD)
      if (needExplode && shouldExplodeBOM) {
        const sheetSemiSpoilage = ss.getSheetByName(CONFIG.SHEET_SPOILAGE_SEMI);
        if (sheetSemiSpoilage) {
          const semiData = sheetSemiSpoilage.getDataRange().getValues().slice(1);
          let foundExploded = false;
          
          semiData.forEach(row => {
            const nvlCode = String(row[1]).trim();
            const nvlQty = Number(row[6]) || 0;
            const nvlNote = String(row[10] || '');
            const nvlDept = String(row[11] || '');
            
            if (nvlNote.includes('Bung từ') && 
                (nvlNote.includes(item.originalName) || nvlNote.includes(String(item.qty)))) {
              
              const nvlMaster = masterMap.get(nvlCode);
              let displayQty = nvlQty;
              let displayUnit = nvlMaster ? nvlMaster.unit : '';
              
              if (unitMode === 'CONVERTED' && nvlMaster && nvlMaster.rate > 1) {
                displayQty = nvlQty / nvlMaster.rate;
                displayUnit = nvlMaster.stdUnit || displayUnit;
              }
              
              const finalCode = formatCodeVN(nvlCode);
              const finalUnit = displayUnit;
              const finalName = nvlMaster ? nvlMaster.name : item.originalName;
              
              // VALIDATION
              if (!finalCode || finalCode === '' || !finalUnit || finalUnit === '') {
                let errReason = !finalCode ? 'THIẾU MÃ HÀNG' : 'THIẾU ĐƠN VỊ TÍNH';
                errorRows.push([finalCode, finalName, finalUnit, roundNum(displayQty, 3), errReason]);
              } else {
                validRows.push([finalCode, finalName, finalUnit, roundNum(displayQty, 3)]);
              }
              foundExploded = true;
            }
          });
          
          if (!foundExploded) {
            errorRows.push([formatCodeVN(item.code), item.name, item.unit, item.qty, 'CHƯA CÓ BOM']);
          }
        }
      }
      // TRƯỜNG HỢP 2: Không cần bung
      else {
        const master = masterMap.get(String(item.code).replace(',', '.'));
        let displayQty = item.qty;
        let displayUnit = item.unit || (master ? master.unit : '');
        
        if (unitMode === 'CONVERTED' && master && master.rate > 1) {
          displayQty = item.qty / master.rate;
          displayUnit = master.stdUnit || displayUnit;
        }
        
        const finalCode = formatCodeVN(item.code);
        const finalUnit = displayUnit;
        
        // VALIDATION
        if (!finalCode || finalCode === '' || !finalUnit || finalUnit === '') {
          let errReason = !finalCode ? 'THIẾU MÃ HÀNG' : 'THIẾU ĐƠN VỊ TÍNH';
          errorRows.push([finalCode, item.name, finalUnit, roundNum(displayQty, 3), errReason]);
        } else {
          validRows.push([finalCode, item.name, finalUnit, roundNum(displayQty, 3)]);
        }
      }
    });
    
    // 3. TẠO FILE EXCEL
    const dateStr = config.fromDate ? config.fromDate.replace(/-/g, '') : new Date().toISOString().slice(0,10).replace(/-/g,'');
    const suffix = shouldExplodeBOM ? '_BUNG_BOM' : '';
    const unitSuffix = unitMode === 'CONVERTED' ? '_CONVERTED' : '';
    const fileName = 'HUY_HANG_SAP_' + dateStr + unitSuffix + suffix;
    
    const newSS = SpreadsheetApp.create(fileName);
    const sheet = newSS.getActiveSheet();
    sheet.setName('Import SAP');
    
    // 4. GHI DỮ LIỆU THEO LAYOUT MỚI
    // Header cột A-D (Hợp lệ)
    const validHeaders = ['ItemCode', 'ItemName', 'UomCode', 'Quantity'];
    sheet.getRange(1, 1, 1, 4).setValues([validHeaders]);
    sheet.getRange(1, 1, 1, 4).setBackground('#4CAF50').setFontColor('white').setFontWeight('bold');
    
    // Header cột F-J (Lỗi) - Cách 1 cột trống (E)
    const errorHeaders = ['ItemCode', 'ItemName', 'UomCode', 'Quantity', 'ErrorReason'];
    sheet.getRange(1, 6, 1, 5).setValues([errorHeaders]);
    sheet.getRange(1, 6, 1, 5).setBackground('#F44336').setFontColor('white').setFontWeight('bold');
    
    // Ghi dữ liệu hợp lệ (A-D)
    if (validRows.length > 0) {
      sheet.getRange(2, 1, validRows.length, 4).setValues(validRows);
    }
    
    // Ghi dữ liệu lỗi (F-J)
    if (errorRows.length > 0) {
      sheet.getRange(2, 6, errorRows.length, 5).setValues(errorRows);
      // Highlight cột ErrorReason
      sheet.getRange(2, 10, errorRows.length, 1).setBackground('#FFCDD2').setFontColor('#B71C1C').setFontWeight('bold');
    }
    
    // Auto-fit columns
    for (let i = 1; i <= 10; i++) sheet.autoResizeColumn(i);
    sheet.setFrozenRows(1);
    
    // Thêm border phân cách giữa 2 bảng (cột E)
    sheet.getRange(1, 5, Math.max(validRows.length, errorRows.length) + 1, 1).setBackground('#ECEFF1');
    
    // 5. DI CHUYỂN FILE VÀO FOLDER CỐ ĐỊNH
    const file = DriveApp.getFileById(newSS.getId());
    const folder = DriveApp.getFolderById(FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // Xóa khỏi root
    
    return {
      success: true,
      message: 'Tạo file thành công! Hợp lệ: ' + validRows.length + ' | Lỗi: ' + errorRows.length,
      url: newSS.getUrl(),
      fileName: fileName,
      validCount: validRows.length,
      errorCount: errorRows.length
    };
    
  } catch (e) {
    console.error('Export Error:', e);
    return { success: false, message: 'Lỗi: ' + e.toString() };
  }
}

// Helper: Format mã theo chuẩn VN (1402194.1 → 1402194,1)
function formatCodeVN(code) {
  if (!code) return '';
  let s = String(code).trim();
  if (/^\d+\.\d+$/.test(s)) s = s.replace('.', ',');
  return s;
}

// Helper: Làm tròn số
function roundNum(n, d) {
  if (!n || isNaN(n)) return 0;
  return Number(Math.round(n + 'e' + d) + 'e-' + d);
}

/**
 * [MỚI] HÀM LẤY TIẾN ĐỘ
 */
function getExportProgress() {
  const cache = CacheService.getScriptCache();
  return cache.get('EXPORT_PROGRESS') || '';
}

/**
 * [GIAI ĐOẠN 1 - FIX 2] TÍNH BOM HỦY HÀNG THEO SỐ LƯỢNG
 */
function calculateSpoilageBOM(itemCode, spoiledQty, spoiledUnit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Lấy Recipe Map
    const recipeMap = getRecipeMap(ss);
    const masterMap = new Map();
    
    // Load Master Data
    const sysData = getSystemData();
    sysData.masterData.forEach(m => masterMap.set(String(m.code).trim(), m));
    
    // Lấy BOM của món hủy
    const recipe = recipeMap[String(itemCode).trim()];
    if (!recipe || !recipe.components || recipe.components.length === 0) {
      return { success: false, message: 'Mã này không có BOM' };
    }
    
    // Tính ratio
    const batchOutput = recipe.batchOutput || 1;
    const ratio = spoiledQty / batchOutput;
    
    // Tính chi tiết
    let totalCost = 0;
    const details = recipe.components.map(comp => {
      const compMaster = masterMap.get(String(comp.code).trim());
      const compCost = compMaster ? (Number(compMaster.cost) || 0) : 0;
      
      const neededQty = comp.qty * ratio;
      const lineCost = neededQty * compCost;
      totalCost += lineCost;
      
      return {
        code: comp.code,
        name: compMaster ? compMaster.name : comp.code,
        batchQty: comp.qty,
        neededQty: Math.round(neededQty * 1000) / 1000,
        unit: compMaster ? compMaster.unit : '',
        unitCost: compCost,
        lineCost: Math.round(lineCost)
      };
    });
    
    return { 
      success: true, 
      data: details,
      totalCost: Math.round(totalCost),
      ratio: Math.round(ratio * 1000) / 1000
    };
    
  } catch (e) {
    return { success: false, message: 'Lỗi: ' + e.toString() };
  }
}
// ============================================================
// GITHUB SYNC ENGINE
// ============================================================

/**
 * [MASTER FUNCTION] Đẩy toàn bộ code lên GitHub
 * Gọi hàm này sau mỗi lần chỉnh sửa code
 */
function pushToGitHub() {
  try {
    const files = {
      'code.gs': getThisScriptContent(),
      'Drawer.html': getFileContent('Drawer'),
      'Footer.html': getFileContent('Footer'),
      'Index.html': getFileContent('Index'),
      'Javascript.html': getFileContent('Javascript'),
      'main.html': getFileContent('main'),
      'section.html': getFileContent('section'),
      'Stylesheet.html': getFileContent('Stylesheet')
    };
    
    let successCount = 0;
    let errorFiles = [];
    
    for (const [filename, content] of Object.entries(files)) {
      try {
        pushSingleFile(filename, content);
        successCount++;
        Utilities.sleep(500); // Tránh rate limit
      } catch (e) {
        errorFiles.push(`${filename}: ${e.toString()}`);
      }
    }
    
    const message = `✅ Pushed ${successCount}/${Object.keys(files).length} files to GitHub\n` +
                   (errorFiles.length > 0 ? `\n⚠️ Errors:\n${errorFiles.join('\n')}` : '');
    
    Logger.log(message);
    return message;
    
  } catch (e) {
    return `❌ Push failed: ${e.toString()}`;
  }
}

/**
 * [HELPER] Push 1 file lên GitHub
 */
function pushSingleFile(filename, content) {
  const url = `https://api.github.com/repos/${GITHUB_CONFIG.REPO}/contents/${filename}`;
  
  // Bước 1: Kiểm tra file đã tồn tại chưa
  let sha = null;
  try {
    const checkResponse = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'token ' + GITHUB_CONFIG.TOKEN,
        'Accept': 'application/vnd.github.v3+json'
      },
      muteHttpExceptions: true
    });
    
    if (checkResponse.getResponseCode() === 200) {
      sha = JSON.parse(checkResponse.getContentText()).sha;
    }
  } catch (e) {
    // File chưa tồn tại, sha = null
  }
  
  // Bước 2: Push/Update file
  const payload = {
    message: `Update ${filename} - ${new Date().toISOString()}`,
    content: Utilities.base64Encode(content),
    branch: GITHUB_CONFIG.BRANCH
  };
  
  if (sha) {
    payload.sha = sha; // Update file cũ
  }
  
  const response = UrlFetchApp.fetch(url, {
    method: 'put',
    headers: {
      'Authorization': 'token ' + GITHUB_CONFIG.TOKEN,
      'Content-Type': 'application/json',
      'Accept': 'application/vnd.github.v3+json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200 && response.getResponseCode() !== 201) {
    throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
  }
}

/**
 * [HELPER] Đọc nội dung code.gs (chính file này)
 */
function getThisScriptContent() {
  try {
    // HACK: Sử dụng eval để lấy source code
    const functionNames = [
      'doGet', 'getSystemData', 'saveSpoilageData', 'saveTransferData',
      'updateTransferFull', 'updateMasterData', 'deleteMasterData',
      'updateSystemData', 'getBOMDetail', 'saveBOM', 'traceIngredientUsage'
      // ... thêm tên các function quan trọng khác
    ];
    
    let sourceCode = '// AUTO-GENERATED FROM APPS SCRIPT\n\n';
    sourceCode += '// ⚠️ THIS IS A MIRROR - DO NOT EDIT DIRECTLY\n';
    sourceCode += '// Edit in Apps Script, then run pushToGitHub()\n\n';
    sourceCode += `// Last sync: ${new Date().toISOString()}\n\n`;
    
    // Lấy source từ toString() của từng function
    functionNames.forEach(funcName => {
      try {
        const func = eval(funcName);
        if (typeof func === 'function') {
          sourceCode += `// ============ ${funcName} ============\n`;
          sourceCode += func.toString() + '\n\n';
        }
      } catch (e) {
        // Function không tồn tại, skip
      }
    });
    
    // Thêm CONFIG
    sourceCode = `const CONFIG = ${JSON.stringify(CONFIG, null, 2)};\n\n` + sourceCode;
    
    return sourceCode;
    
  } catch (e) {
    return `// Error reading code.gs: ${e.toString()}\n// Please create Backend.html as fallback`;
  }
}

/**
 * [HELPER] Đọc nội dung file HTML (giữ nguyên từ code cũ)
 */
function getFileContent(fileName) {
  try {
    return HtmlService.createHtmlOutputFromFile(fileName).getContent();
  } catch (e) {
    return `// Error reading ${fileName}: ${e.toString()}`;
  }
}

/**
 * [TRIGGER] Tự động sync mỗi 6 giờ (setup ở Bước 3)
 */
function autoSyncToGitHub() {
  pushToGitHub();
}


/**
 * KIỂM TRA FILE CODE.GS HIỆN TẠI
 */
function debugCodeGsFile() {
  try {
    Logger.log('🔍 Đang kiểm tra file code.gs...');
    Logger.log('================================');
    
    // Thử đọc qua getScriptContent
    let content = '';
    try {
      content = getScriptContent('code.gs');
      Logger.log('✅ Đọc được qua getScriptContent()');
    } catch (e) {
      Logger.log('❌ Không đọc được qua getScriptContent(): ' + e.toString());
      
      // Thử cách khác
      try {
        const html = HtmlService.createHtmlOutputFromFile('Backend').getContent();
        const match = html.match(/<script>([\s\S]*?)<\/script>/);
        if (match) {
          content = match[1].trim();
          Logger.log('✅ Đọc được qua Backend.html');
        }
      } catch (e2) {
        Logger.log('❌ Cũng không đọc được qua Backend.html: ' + e2.toString());
      }
    }
    
    if (!content) {
      Logger.log('');
      Logger.log('⚠️ KHÔNG THỂ ĐỌC FILE CODE.GS!');
      Logger.log('');
      Logger.log('📋 Kiểm tra:');
      Logger.log('1. File Backend.html có tồn tại không?');
      Logger.log('2. Backend.html có chứa code trong <script>...</script> không?');
      Logger.log('3. Code có quá dài (>5MB) không?');
      
      SpreadsheetApp.getUi().alert(
        '❌ Không đọc được code.gs',
        'Không thể đọc file code.gs.\n\n' +
        'Kiểm tra log để biết chi tiết.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Phân tích nội dung
    const lines = content.split('\n').length;
    const chars = content.length;
    const kb = (chars / 1024).toFixed(2);
    
    Logger.log('');
    Logger.log('📊 THÔNG TIN FILE:');
    Logger.log(`  - Số dòng: ${lines}`);
    Logger.log(`  - Số ký tự: ${chars}`);
    Logger.log(`  - Kích thước: ${kb} KB`);
    Logger.log('');
    
    // Kiểm tra các hàm quan trọng
    const functionsToCheck = [
      'onOpen', 'XN', 'XK', 'KC', 
      'calculateKitPrice', 'getCurrentStock',
      'pushToGitHub', 'handleGitHubSync'
    ];
    
    Logger.log('🔍 KIỂM TRA CÁC HÀM:');
    
    const found = [];
    const missing = [];
    
    functionsToCheck.forEach(funcName => {
      // Tìm pattern: function funcName( hoặc const funcName = 
      const pattern1 = new RegExp(`function\\s+${funcName}\\s*\\(`);
      const pattern2 = new RegExp(`const\\s+${funcName}\\s*=`);
      const pattern3 = new RegExp(`var\\s+${funcName}\\s*=`);
      
      if (pattern1.test(content) || pattern2.test(content) || pattern3.test(content)) {
        found.push(funcName);
        Logger.log(`  ✅ ${funcName}`);
      } else {
        missing.push(funcName);
        Logger.log(`  ❌ ${funcName}`);
      }
    });
    
    Logger.log('');
    Logger.log('================================');
    Logger.log(`📊 Kết quả: ${found.length}/${functionsToCheck.length} hàm tìm thấy`);
    
    // Hiển thị kết quả
    let message = `📊 File code.gs:\n\n`;
    message += `  • ${lines} dòng\n`;
    message += `  • ${chars} ký tự\n`;
    message += `  • ${kb} KB\n\n`;
    
    if (lines < 1000 || chars < 30000) {
      message += `⚠️ File có vẻ CHƯA ĐẦY ĐỦ!\n\n`;
    }
    
    message += `🔍 Functions:\n`;
    message += `  ✅ Tìm thấy: ${found.length}\n`;
    message += `  ❌ Thiếu: ${missing.length}\n\n`;
    
    if (missing.length > 0) {
      message += `Thiếu:\n${missing.map(f => `  • ${f}`).join('\n')}`;
    }
    
    SpreadsheetApp.getUi().alert(
      missing.length === 0 ? '✅ Code.gs OK' : '⚠️ Code.gs thiếu code',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return {
      lines,
      chars,
      kb,
      found,
      missing,
      isComplete: missing.length === 0 && lines > 1000
    };
    
  } catch (e) {
    Logger.log('❌ Lỗi: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert('❌ Lỗi: ' + e.toString());
  }
}
