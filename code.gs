const CONFIG = {
  "SHEET_NVL": "Danh m?c NVL",
  "SHEET_PROD": "Danh m?c Product",
  "SHEET_KIT_KITCHEN": "Kit Semi Store",
  "SHEET_KIT_PIZZA": "Pzz Semi Store",
  "SHEET_KIT_SERVICE": "Semi Store Service",
  "SHEET_ML_NVL": "H?c m?y",
  "SHEET_ML_PROD": "H?c m?y Product",
  "SHEET_ML_SEMI": "H?c m?y Semi Store",
  "SHEET_ML_TRANSFER": "H?c m?y Transfer",
  "SHEET_STORE_LIST": "H?c m?y Store",
  "SHEET_TRANSFER_DATA": "Transfer",
  "SHEET_SPOILAGE": "H?y NVL",
  "SHEET_SPOILAGE_SEMI": "H?y Semi Store",
  "SHEET_SPOILAGE_PROD": "H?y Product",
  "SHEET_BOM_PRODUCT": "BOM Product",
  "SHEET_BOM_CACHE": "DB_BOM_CACHE"
};

// AUTO-GENERATED FROM APPS SCRIPT

// ?? THIS IS A MIRROR - DO NOT EDIT DIRECTLY
// Edit in Apps Script, then run pushToGitHub()

// Last sync: 2025-12-25T02:59:55.930Z

// ============ doGet ============
function doGet(e) {
  // [GI? NGUY?N CODE C?] - Ph?n render HTML
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('V12.19 Analyst Center')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // [TH?M M?I] - API tr? code (d?ng 40-66)
  if (e.parameter.action === 'getCode') {
  const fileMap = {
    // Backend - Placeholder v? kh?ng ??c ???c .gs file
    'Backend': `/* 
 * CODE BACKEND (code.gs)
 * ?? xem full code, vui l?ng t?o file Backend.html
 * ho?c xem tr?c ti?p trong Apps Script Editor.
 * 
 * Script ID: ${ScriptApp.getScriptId()}
 * Version: V12.19
 */`,
    
    // C?c file HTML (GI? NGUY?N)
    'Drawer': getFileContent('Drawer'),
    'Footer': getFileContent('Footer'),
    'Index': getFileContent('Index'),
    'Javascript': getFileContent('Javascript'),
    'main': getFileContent('main'),
    'section': getFileContent('section'),
    'Stylesheet': getFileContent('Stylesheet')
  };
  
  const requestedFile = e.parameter.file || 'Backend';
  const content = fileMap[requestedFile] || 'File not found';
  
  const response = {
    file: requestedFile,
    content: content,
    timestamp: new Date().toISOString(),
    version: 'V12.19'
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * [HELPER] ??c n?i dung file HTML
 */
function getFileContent(fileName) {
  try {
    return HtmlService.createHtmlOutputFromFile(fileName).getContent();
  } catch (e) {
    return `// Error reading ${fileName}: ${e.toString()}`;
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('? SYSTEM ADMIN')
    .addItem('? T?i t?o to?n b? BOM Cache', 'regenerateAllBOMCache')
    .addSeparator()
    .addItem('? Push to GitHub', 'pushToGitHub') // ? TH?M D?NG N?Y
    .addToUi();
}

  // --- CORE SYSTEM DATA ---
  function getSystemData() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // H?m ?p ki?u s? an to?n (x? l? d?u ph?y Vi?t Nam)
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
      // NVL: Gi? v?n ? C?t D (Index 3)
      const listNVL = getData(CONFIG.SHEET_NVL).map((r, i) => ({ 
        rowId: i + 2, 
        name: String(r[0]),           // A: T?n
        code: cleanCode(r[1]),        // B: M?
        unit: String(r[2]),           // C: ?VT G?c
        cost: safeFloat(r[3]),        // D: Gi? V?n
        stdUnit: String(r[4]),        // E: ?VT Quy ??i (Quan tr?ng)
        rate: Number(r[5]) || 1,      // F: H? s?
        buyingPrice: safeFloat(r[6]), // G: Gi? Mua (M?I)
        supplier: String(r[7]),       // H: Nh? cung c?p (M?I)
        leadtime: String(r[8]),       // I: Leadtime (M?I)
        noDelivery: String(r[9]),     // J: No Delivery (M?I)
        group: String(r[10]),         // K: Group h?ng (M?I)
        status: String(r[11]),        // L: Tr?ng th?i (M?I)
        type: 'NVL', 
        rawData: cleanRawRow(r) 
      }));

      // PROD: Gi? v?n ? C?t G (Index 6)
      const listProd = getData(CONFIG.SHEET_PROD).map((r, i) => ({ 
        rowId: i + 2, 
        code: cleanCode(r[0]), 
        name: String(r[1]),    
        category: String(r[2]),
        class: String(r[3]),   
        team: String(r[4]),    
        sapCode: String(r[5]), 
        cost: safeFloat(r[6]), 
        status: String(r[7]) || 'Active', // L?y th?m tr?ng th?i
        rate: 1, unit: 'C?i', standardUnit: 'C?i', type: 'PROD', 
        rawData: cleanRawRow(r) 
      }));

      // SEMI STORE: ??c gi? t? Sheet (Kh?ng t?nh to?n l?i ?? load nhanh)
      const listSemi = [];
      [CONFIG.SHEET_KIT_KITCHEN, CONFIG.SHEET_KIT_PIZZA, CONFIG.SHEET_KIT_SERVICE].forEach(sheetName => {
          getData(sheetName).forEach(r => {
              if (String(r[1]).toLowerCase() === 'parent') {
                  listSemi.push({ 
                      rowId: -1, 
                      code: cleanCode(r[0]), 
                      name: String(r[3]), 
                      unit: String(r[5]), 
                      
                      // [CH? IT UPDATE] ??c th?ng gi? t? Sheet
                      cost: safeFloat(r[6]),       // C?t G: Gi? v?n ??n v?
                      batchPrice: safeFloat(r[7]), // C?t H: Gi? Batch (M?i)
                      
                      yield: safeFloat(r[4]),      // C?t E: ??nh l??ng
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
              // C?p nh?t gi? cho m? ?? t?n t?i
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
    // [CH? IT UPDATE] Lu?n gi? 3 s? l? (0.001) cho m?i tr??ng h?p ?? ??m b?o ch?nh x?c cho Transfer/H?y
    return Math.round(val * 1000) / 1000;
  }

  /* [T?I ?U T?C ??] L?u H?y: S? d?ng Bulk Insert (Array) cho c? 3 Sheet, t?c ?? l?u < 2 gi?y */
function saveSpoilageData(p){try{const ss=SpreadsheetApp.getActiveSpreadsheet(),sSp=ss.getSheetByName(CONFIG.SHEET_SPOILAGE),sSm=ss.getSheetByName(CONFIG.SHEET_SPOILAGE_SEMI),sPr=ss.getSheetByName(CONFIG.SHEET_SPOILAGE_PROD);if(!sSp||!sSm||!sPr)return{success:false,message:"Thi?u Sheet H?y!"};const sys=getSystemData(),mMap=new Map(),rMap=getRecipeMap(ss),nMap=createNameMap(ss);sys.masterData.forEach(m=>mMap.set(String(m.code).trim(),m));let d=p.date;if(d.includes("-")){let x=d.split('-');d=`${x[2]}/${x[1]}/${x[0]}`}const rSp=[],rSm=[],rPr=[];p.items.forEach(i=>{const c=String(i.code).trim(),mI=mMap.get(c),uP=mI?Number(mI.cost)||0:0,amt=Math.round(i.qty*uP);if(i.itemType==='PROD'){rPr.push([d,i.name,i.code,"",i.qty,i.note||"",p.dept,"","",mI?mI.category:"",mI?mI.class:"",uP])}else if(i.itemType==='SEMI'){const leaves=bomEngine(c,i.qty,rMap,mMap);const isBk=(leaves.length===1&&String(leaves[0].code).trim()===c);let wn=isBk?" | ?? CH?A BOM":"";rSp.push([d,c,i.name,"",i.unit,i.factor||1,"",i.unit,i.qty,"",(i.note||"H?y BTP")+wn,p.dept,"",uP,amt]);if(!isBk){const cons={};leaves.forEach(l=>{cons[l.code]=(cons[l.code]||0)+l.qty_raw});for(const[lC,lQ]of Object.entries(cons)){rSm.push([d,lC,"","","","",smartRound(lQ),"","","",`Bung t? ${i.qty} ${i.name}`,p.dept])}}}else{rSp.push([d,c,i.name,"",i.unit,i.factor||1,"",i.unit,i.qty,"",(i.note||""),p.dept,"",uP,amt])}});if(rSp.length>0)sSp.getRange(findNextEmptyRow(sSp),1,rSp.length,15).setValues(rSp);if(rSm.length>0)sSm.getRange(findNextEmptyRow(sSm),1,rSm.length,12).setValues(rSm);if(rPr.length>0)sPr.getRange(findNextEmptyRow(sPr),1,rPr.length,12).setValues(rPr);return{success:true,message:"? ?? l?u phi?u H?y th?nh c?ng & Kh?p c?t!"}}catch(e){return{success:false,message:"L?i: "+e.toString()}}}

  function saveTransferData(items) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_TRANSFER_DATA); // [1]
    if (!sheet) return { success: false, message: "Kh?ng t?m th?y Sheet Transfer" };

    // 1. Chu?n b? Map Gi? V?n (Ch? ??c 1 l?n)
    const costMap = new Map();
    [CONFIG.SHEET_NVL, CONFIG.SHEET_PROD].forEach(name => {
      const s = ss.getSheetByName(name);
      if (s) {
        const vals = s.getDataRange().getValues();
        // [2] X?c ??nh c?t gi? d?a tr?n t?n Sheet
        const isNVL = name === CONFIG.SHEET_NVL; 
        vals.forEach(r => {
          // NVL: Code c?t B(1), Gi? c?t D(3). PROD: Code c?t A(0), Gi? c?t G(6)
          if (isNVL) costMap.set(String(r[3]).trim(), Number(r[4]) || 0);
          else costMap.set(String(r).trim(), Number(r[5]) || 0);
        });
      }
    });

    const outputRows = [];
    const startRow = findNextEmptyRow(sheet); // [6] Ch? t?m d?ng tr?ng 1 l?n ??u ti?n

    // 2. X? l? d? li?u trong b? nh? (RAM)
    items.forEach(item => {
      // X? l? ng?y th?ng
      let d = item.date;
      if (d.includes("-")) { let p = d.split('-'); d = `${p[7]}/${p[3]}/${p}`; }

      // Logic t?nh to?n Total Base Qty (Quan tr?ng ?? tr? kho ??ng)
      // [8] N?u l? State 2 (Chia) ho?c 3 (Hack Unit), nh?n ng??c l?i ra s? g?c
      let rate = Number(item.rate) || 1;
      let totalBaseQty = 0;
      let qtyDisplay = item.qty;

      if (item.rateState === 2 || item.rateState === 3) {
         // Tr??ng h?p nh?p theo Th?ng/Quy ??i
         totalBaseQty = item.qty * rate; 
      } else {
         // Tr??ng h?p nh?p L? ho?c Nh?n
         totalBaseQty = item.qty;
      }
      
      // An to?n: N?u Frontend c? g?i originalQty th? ?u ti?n ki?m tra, nh?ng logic tr?n l? "Ch?t ch?n" cu?i c?ng.
      
      // L?y gi? v?n ??n v?
      const unitCost = costMap.get(String(item.code).trim()) || 0;
      // T?nh th?nh ti?n: Ph?i nh?n v?i T?NG S? L??NG G?C (Total Base Qty)
      const totalAmount = Math.round(unitCost * totalBaseQty);

      // Chu?n b? d?ng d? li?u (Mapping theo ??ng c?t trong Sheet Transfer)
      // [9]-[10] C?u tr?c c?t
      outputRows.push([
        d,                                      // A: Ng?y
        item.sender,                            // B: Ng??i g?i
        item.type,                              // C: Lo?i (IN/OUT)
        item.receiver,                          // D: Ng??i nh?n
        item.storeName,                         // E: Store
        item.name,                              // F: T?n h?ng
        item.code,                              // G: M? h?ng
        (item.rateState === 2 || item.rateState === 3) ? qtyDisplay : "", // H: SL Quy ??i
        item.standardUnit,                      // I: ?VT Quy ??i
        rate,                                   // J: H? s?
        (item.rateState === 2 || item.rateState === 3) ? "" : qtyDisplay, // K: SL L?
        item.originalUnit || item.baseUnit || item.unit, // L: ?VT G?c (B?t bu?c)
        totalBaseQty,                           // M: T?ng SL G?c (QUAN TR?NG NH?T)
        item.team || "",                        // N: Team
        unitCost,                               // O: Gi? v?n
        totalAmount,                            // P: Th?nh ti?n
        false                                   // Q: Checkbox Status
      ]);
    });

    // 3. Ghi xu?ng Sheet 1 l?n duy nh?t (T?c ?? cao)
    if (outputRows.length > 0) {
      sheet.getRange(startRow, 1, outputRows.length, 17).setValues(outputRows);
    }

    return { success: true, message: `?? l?u th?nh c?ng ${outputRows.length} d?ng!` };

  } catch (e) {
    return { success: false, message: "L?i Server: " + e.toString() };
  }
}

  /* [??NG B?] C?p nh?t phi?u: Logic kh?p 100% v?i saveTransferData ?? tr?nh l?ch kho */
function updateTransferFull(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_TRANSFER_DATA); // [cite: 73]
    const r = p.id;
    if (!r || r < 2) return { success: false, message: "ID d?ng l?i" };

    // 1. Chu?n h?a ng?y th?ng (S?a l?i Index m?ng p)
    let d = p.date;
    if (d && d.includes("-")) {
      let x = d.split('-');
      d = `${x[2]}/${x[1]}/${x[0]}`; // ??nh d?ng chu?n DD/MM/YYYY
    }

    // 2. Logic t?nh to?n TotalBaseQty (Gi? nguy?n logic p.rateState 2 & 3 nh? b?n y?u c?u)
    let rate = Number(p.rate) || 1;
    let totalBaseQty = 0;
    if (p.rateState === 2 || p.rateState === 3) {
      totalBaseQty = p.qty * rate;
    } else {
      totalBaseQty = p.qty;
    }

    // 3. Tra c?u gi? v?n (S?a ??ng Index theo c?u tr?c Master Data)
    let bC = 0;
    const tC = String(p.code || '').trim();
    if (tC) {
      const sNvl = ss.getSheetByName(CONFIG.SHEET_NVL);
      const sPr = ss.getSheetByName(CONFIG.SHEET_PROD);
      
      // Ki?m tra trong NVL: M? [1], Gi? [3]
      if (sNvl) {
        const foundNvl = sNvl.getDataRange().getValues().find(row => String(row[1]).trim() === tC);
        if (foundNvl) bC = Number(foundNvl[3]) || 0;
      }
      
      // N?u kh?ng th?y trong NVL, ki?m tra trong PROD: M? [0], Gi? [6]
      if (bC === 0 && sPr) {
        const foundPr = sPr.getDataRange().getValues().find(row => String(row[0]).trim() === tC);
        if (foundPr) bC = Number(foundPr[6]) || 0;
      }
    }

    // 4. Chu?n b? d?ng d? li?u (Ghi t? c?t A ??n P - 16 c?t)
    let vals = [[
      d,                                    // A: Ng?y
      p.sender || '',                       // B: Ng??i g?i
      p.type || 'UNK',                      // C: Lo?i
      p.receiver || '',                     // D: Ng??i nh?n
      p.store || '',                        // E: Store
      p.itemName || p.name || '',           // F: T?n
      p.code || '',                         // G: M?
      (p.rateState === 2 || p.rateState === 3) ? p.qty : "", // H: SL Quy ??i
      p.standardUnit || '',                 // I: ?VT Quy ??i
      rate,                                 // J: H? s?
      (p.rateState === 2 || p.rateState === 3) ? "" : p.qty, // K: SL L?
      p.originalUnit || p.unit || '',       // L: ?VT G?c
      totalBaseQty,                         // M: T?NG SL G?C
      p.team || '',                         // N: Team
      bC,                                   // O: Gi? v?n ??n v?
      Math.round(bC * totalBaseQty)         // P: Th?nh ti?n
    ]];

    // 5. Ghi d? li?u - [cite: 179-180]
    sheet.getRange(r, 1, 1, 16).setValues(vals);

    return { success: true, message: "? ?? c?p nh?t phi?u & ??ng b? logic th?nh c?ng!" };

  } catch (e) {
    return { success: false, message: "L?i Server Update: " + e.toString() };
  }
}


  /**
   * [V12.34 FIX MASTER DATA]
   * - Fix l?i Semi: S?a = X?a C? + Th?m M?i (?? c?p nh?t BOM con).
   * - Fix l?i NVL/Prod: C?p nh?t ??y ?? c?c tr??ng (NCC, Gi?, Group...) khi s?a.
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
    if (!sheet) return { success: false, message: "Kh?ng t?m th?y Sheet: " + sheetName };

    // --- CASE SPECIAL: SEMI (X? L? D?NG CHA & CON) ---
    if (action === 'EDIT' && payload.type === 'SEMI' && payload.rowId) {
      const oldCode = sheet.getRange(payload.rowId, 1).getValue();
      const data = sheet.getDataRange().getValues();
      // X?a t? d??i l?n ?? kh?ng l?ch Index d?ng
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
        // Map gi? v?n NVL: M? [1], Gi? [3] [cite: 11-12]
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

    // C?p nh?t Cache th?ng minh [cite: 123]
    if (payload.type === 'SEMI' || payload.type === 'PROD') {
      try { updateSingleBOMCache(payload.code); } catch (e) { console.log("L?i Cache: " + e.toString()); }
    }
    return { success: true, message: `? ?? c?p nh?t ${payload.code}` };
  } catch (e) { return { success: false, message: "L?i: " + e.toString() }; }
}

  function deleteMasterData(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName = (payload.type === 'PROD') ? CONFIG.SHEET_PROD : CONFIG.SHEET_NVL;
    const sheet = ss.getSheetByName(sheetName);
    if (payload.rowId) { sheet.deleteRow(payload.rowId); return { success: true, message: "?? x?a m?: " + payload.code }; }
    return { success: false, message: "Kh?ng t?m th?y ID!" };
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
      return { success: true, message: "?? c?p nh?t ML" };
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
        return { success: true, message: "?? x?a ML" };
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
        return { success: true, message: "?? x?a d?ng th?nh c?ng!" };
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
      
      // [LOGIC M?I] N?u ngu?n l? Danh m?c Product -> ??c t? Sheet BOM Product
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

  // --- CORE BOM SAVING (L?U GI? V?O C?T G V? H) ---
  /* [MODULE MASTER] H?M L?U ??NH M?C BOM - B?N NGUY?N KH?I CHU?N V12.19 [cite: 1, 2025-12-21] */
function saveBOM(p) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. KI?M TRA LO?I BOM V? X?C ??NH SHEET [cite: 1, 2025-12-21]
    const isProductBOM = (p.sourceSheet === CONFIG.SHEET_PROD);
    const targetSheetName = isProductBOM ? CONFIG.SHEET_BOM_PRODUCT : p.sourceSheet;
    const sheetDetail = ss.getSheetByName(targetSheetName);
    if (!sheetDetail) return { success: false, message: "Thi?u Sheet: " + targetSheetName };

    const itemCode = cleanCode(p.itemCode);

    // 2. X?A D? LI?U BOM C? [cite: 1, 2025-12-21]
    const data = sheetDetail.getDataRange().getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (cleanCode(data[i][0]) === itemCode) {
        sheetDetail.deleteRow(i + 1);
      }
    }

    // 3. T?NH TO?N GI? V? CHU?N B? D? LI?U M?I [cite: 1, 2025-12-21]
    const sysData = getSystemData();
    const allMaster = sysData.masterData;
    let totalBatchCost = 0;
    const newRows = [];

    // T?o d?ng Header cho Sheet Chi ti?t
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

    // 4. GHI D? LI?U V? C?P NH?T CACHE T? ??NG [cite: 1, 2025-12-21]
    if (newRows.length > 0) {
      sheetDetail.getRange(sheetDetail.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      
      // C?p nh?t Cache ri?ng l? (Partial Update) - D?ng ??ng bi?n itemCode [cite: 1, 2025-12-21]
      try {
        if (itemCode) { 
          updateSingleBOMCache(itemCode); 
        }
      } catch (cacheErr) { 
        console.log("L?i Cache: " + cacheErr.toString()); 
      }
    }

    // 5. C?P NH?T GI? V?N V?O DANH M?C PRODUCT (N?U C?) [cite: 1, 2025-12-21]
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

    // 6. PH?N H?I K?T TH?C ?? T?T LOADING TR?N WEB [cite: 1, 1809, 2025-12-21]
    return { 
      success: true, 
      message: "?? l?u BOM m?n " + p.itemName + " th?nh c?ng!",
      newCost: finalUnitCost,
      newBatchPrice: totalBatchCost 
    };

  } catch (err) {
    return { success: false, message: "L?i h? th?ng: " + err.toString() };
  }
} // <--- ?? d?u ??ng k?t th?c h?m [cite: 1, 2025-12-21]

  function traceIngredientUsage(nvlCode) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      // [M?I] B? sung SHEET_BOM_PRODUCT v?o danh s?ch ?i t?m
      const sheets = [
        { name: CONFIG.SHEET_KIT_KITCHEN, team: 'KITCHEN' },
        { name: CONFIG.SHEET_KIT_PIZZA, team: 'PIZZA' },
        { name: CONFIG.SHEET_KIT_SERVICE, team: 'SERVICE' },
        { name: CONFIG.SHEET_BOM_PRODUCT, team: 'PRODUCT' } // <-- Th?m d?ng n?y
      ];
      
      const results = [];
      const target = String(nvlCode).trim();
      
      sheets.forEach(conf => {
          const sheet = ss.getSheetByName(conf.name);
          if (sheet) {
              const data = sheet.getDataRange().getValues();
              data.forEach(r => {
                  // C?t B (index 1) l? 'Child', C?t C (index 2) l? M? Con
                  if (String(r[1]).toLowerCase() === 'child' && String(r[2]).trim() === target) {
                      results.push({ 
                          parentCode: r[0], // M? Cha
                          team: conf.team,
                          qty: Number(r[4]) || 0, // SL
                          unit: String(r[5]) || '' // ?VT
                      });
                  }
              });
          }
      });

      // L?c tr?ng l?p (Unique)
      const unique = [];
      const map = new Map();
      for (const item of results) {
          // T?o kh?a unique l? M? cha + Team
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

  // --- BULK UPDATE TOOL (C?p nh?t c? c?t G v? H) ---
  /**
   * H?M 1: L?Y DANH S?CH M?N CHA ?ANG S? D?NG NVL (S?a l?i hi?n th? s? 0)
   * @param {string} ingredientCode - M? nguy?n li?u (Con)
   */
  function getParentItemsUsingIngredient(ingredientCode) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("BOM"); // ?? CH? ?: ??i t?n Sheet cho ??ng file c?a ch?u
    var data = sheet.getDataRange().getValues();
    
    var result = [];
    
    // Gi? ??nh c?u tr?c c?t BOM (Ch?u ??m l?i c?t trong file Excel nh? A=0, B=1...)
    // V? d?: A=M? Cha, B=T?n Cha, C=M? Con, D=T?n Con, E=??nh L??ng, F=?VT
    const COL_PARENT_CODE = 0; 
    const COL_PARENT_NAME = 1;
    const COL_CHILD_CODE = 2;
    const COL_QTY = 4; // ?? QUAN TR?NG: Ki?m tra l?i c?t E (??nh l??ng) c? ??ng l? index 4 kh?ng?
    const COL_UNIT = 5;

    for (var i = 1; i < data.length; i++) { // B? qua header
      var row = data[i];
      // So s?nh M? Con, chuy?n v? String ?? tr?nh l?i s?/ch?
      if (String(row[COL_CHILD_CODE]) === String(ingredientCode)) {
        result.push({
          parentCode: row[COL_PARENT_CODE],
          parentName: row[COL_PARENT_NAME],
          // S?A L?I ? ??Y: ??m b?o parse s?, n?u l?i th? v? 0
          qty: parseFloat(row[COL_QTY]) || 0, 
          unit: row[COL_UNIT]
        });
      }
    }
    
    return result;
  }

  /**
   * H?M 2: T?M V? THAY TH? NVL TRONG BOM (S?a l?i n?t kh?ng ch?y)
   * @param {string} oldCode - M? c? c?n thay
   * @param {string} newCode - M? m?i thay th? v?o
   */
  function replaceIngredientInBom(oldCode, newCode) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("BOM"); // ?? CH? ?: ??i t?n Sheet
    var range = sheet.getDataRange();
    var data = range.getValues();
    
    const COL_CHILD_CODE = 2; // C?t ch?a M? Nguy?n Li?u (Con) - Index 2 = C?t C
    var changeCount = 0;

    // Qu?t v? thay th? trong m?ng (Memory) cho nhanh
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][COL_CHILD_CODE]) === String(oldCode)) {
        data[i][COL_CHILD_CODE] = newCode; // Thay th? m?
        changeCount++;
      }
    }

    // Ghi ng??c l?i xu?ng Sheet (Ch? ghi n?u c? thay ??i)
    if (changeCount > 0) {
      range.setValues(data);
      return { success: true, message: "?? thay th? th?nh c?ng " + changeCount + " d?ng!" };
    } else {
      return { success: false, message: "Kh?ng t?m th?y m? c? " + oldCode + " trong BOM." };
    }
  }

  function importSalesHistory(payload) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName("Data Doanh Thu");
      if (!sheet) {
          sheet = ss.insertSheet("Data Doanh Thu");
          sheet.appendRow(["Ng?y Import", "Kho?ng Th?i Gian", "Chi Nh?nh", "M? SP", "T?n SP", "T?ng SL", "SL Delivery", "SL Dine-in", "SL Take-away", "Ng??i Nh?p"]);
      }
      
      const importDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      const period = payload.period || "";
      let userEmail = "";
      try { userEmail = Session.getActiveUser().getEmail(); } catch(e) { userEmail = "Unknown User"; }
      
      const rows = [];
      // H?m ?p ki?u an to?n
      const safeNum = (n) => typeof n === 'number' ? n : (Number(n) || 0);

      payload.items.forEach(item => {
          rows.push([
              importDate,
              period,
              item.store || "",
              String(item.code || ""), // ?p v? String ?? gi? s? 0 ??u
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
      
      return { success: true, message: "?? l?u th?nh c?ng " + rows.length + " d?ng!" };
    } catch (e) {
      return { success: false, message: "L?i Backend: " + e.toString() };
    }
  }

  /**
   * H?M M?I: Th?m nhanh danh s?ch Product m?i t? file Import Doanh Thu
   * T? ??ng map c?c c?t: Code, Name, Category, Class, Price
   */
  function quickAddMissingProducts(items) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PROD); // Danh m?c Product
    if (!sheet) return { success: false, message: "Kh?ng t?m th?y Sheet Danh m?c Product!" };

    try {
      const rowsToAdd = [];
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

      items.forEach(item => {
        // C?u tr?c c?t Sheet Product: 
        // A: Code | B: Name | C: Category | D: Class | E: Team | F: SAP Code | G: Cost | H: Status
        rowsToAdd.push([
          String(item.code),       // A: Code (D?ng l?m m? n?i b? lu?n)
          item.name,               // B: T?n
          item.category || "",     // C: Category (L?y t? file Excel)
          item.class || "",        // D: Class (L?y t? file Excel)
          "Service",               // E: Team (M?c ??nh Service v? b?n h?ng)
          String(item.code),       // F: SAP Code
          0,                       // G: Gi? v?n (T?m ?? 0)
          "Active"                 // H: Tr?ng th?i
        ]);
      });

      if (rowsToAdd.length > 0) {
        // T?m d?ng tr?ng ti?p theo
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      }

      return { success: true, message: `?? th?m th?nh c?ng ${rowsToAdd.length} m? m?i!` };
    } catch (e) {
      return { success: false, message: "L?i Backend: " + e.toString() };
    }
  }

  // [SOURCE: KAIZEN BRAIN SYSTEM]
// ==========================================================

// ?? THAY ID FILE TXT C?A CH?U V?O ??Y
const KAIZEN_CONFIG = {
  BRAIN_FILE_ID: "14K3qOvEtsLfmo_XJfV2yxbYJBknulgmL" 
};

/**
 * H?M 1: ??C TRI TH?C (D?ng ?? l?y Context n?m cho AI ??u bu?i)
 */
function getKaizenBrain() {
  try {
    const file = DriveApp.getFileById(KAIZEN_CONFIG.BRAIN_FILE_ID);
    return { 
      success: true, 
      content: file.getBlob().getDataAsString() 
    };
  } catch (e) {
    return { success: false, message: "L?i ??c n?o: " + e.toString() };
  }
}

// H?m n?y ch? d?ng ?? Ch?u ki?m tra Log th?i nh?
function debugReadBrain() {
  const data = getKaizenBrain();
  console.log("=== K?T QU? ??C N?O ===");
  console.log(data); 
  console.log("=======================");
}

/**
 * H?M 2: N?P TRI TH?C M?I (AI g?i h?m n?y qua User)
 * @param {string} newRuleContent - N?i dung quy t?c m?i
 */
function appendKaizenRule(newRuleContent) {
  try {
    const file = DriveApp.getFileById(KAIZEN_CONFIG.BRAIN_FILE_ID);
    const currentContent = file.getBlob().getDataAsString();
    
    // Timestamp ??nh d?ng VN
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    // T?o Block n?i dung m?i, c? ph?n c?ch r? r?ng
    const updateBlock = `\n\n# [UPDATE OTA: ${time}] ---------------------------\n${newRuleContent}`;
    
    // Ghi ??: N?i dung c? + Block m?i
    file.setContent(currentContent + updateBlock);
    
    return { success: true, message: `?? n?p tri th?c l?c ${time}!` };
  } catch (e) {
    return { success: false, message: "L?i n?p tri th?c: " + e.toString() };
  }
}

/**
 * H?M 3: T?I C?U TR?C N?O B? (D?ng khi Clean d?n d?p)
 * H?m n?y s? X?A S?CH c? v? GHI M?I to?n b?.
 */
function rewriteKaizenBrain(fullContent) {
  try {
    const file = DriveApp.getFileById(KAIZEN_CONFIG.BRAIN_FILE_ID);
    
    // Ghi ?? to?n b? n?i dung (setContent thay v? append)
    file.setContent(fullContent);
    
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    return { success: true, message: `?? t?i c?u tr?c n?o b? th?nh c?ng l?c ${time}!` };
  } catch (e) {
    return { success: false, message: "L?i ghi ?? n?o: " + e.toString() };
  }
}

function rebuildBOMCache(productID) {
  // 1. L?y d? li?u BOM g?c (D?ng c?y)
  let bomTree = getBOMTree(productID); 

  // 2. Bi?n ch?a k?t qu? ph?ng
  let flatList = [];

  // 3. H?m ?? quy ?? ??o s?u v? l?m ph?ng (CH? IT ?? ?? L?I)
  function flatten(node, currentYield, pathString) {
    
    // Ki?m tra an to?n: n?u node kh?ng c? children th? d?ng
    if (!node.children || node.children.length === 0) return;

    // Duy?t qua t?ng th?nh ph?n con
    node.children.forEach(child => {
      
      // T?nh Yield t?ch l?y
      let accumulatedYield = currentYield * (child.inputQty / child.outputQty);
      
      // --- [CH? IT FIX HI?N TH? T?I ??Y] ---
      
      // A. L?y ?VT (N?u h?m getBOMTree ch?a tr? v? unit, ch?u nh? ki?m tra l?i h?m ??)
      let unit = child.unit || ""; 
      
      // B. L?m tr?n s? l??ng cho g?n (3 s? l?), ??i d?u ch?m th?nh ph?y cho chu?n VN
      let qtyPretty = (Math.round(child.usage * 1000) / 1000).toString().replace('.', ',');

      // C. T?o chu?i hi?n th? b??c hi?n t?i: V? d? "(50 Gr) Ph? mai"
      let stepInfo = `[${qtyPretty} ${unit}] ${child.name}`;

      // D. N?i chu?i: D?ng d?u m?i t?n ??m "?" thay v? d?u ">" nh?n cho x?n
      let newPath = pathString + " ? " + stepInfo;
      
      // -------------------------------------

      if (child.type === 'RAW') {
        // ?I?M D?NG: N?u l? Raw, ghi v?o danh s?ch k?t qu?
        flatList.push({
          product_id: productID,
          raw_id: child.id,
          total_qty: child.usage * accumulatedYield, // CON S? V?NG
          path_log: newPath, // ???ng d?n ??p ?? ???c t?o ? tr?n
          updated: new Date()
        });
      } else if (child.type === 'SEMI') {
        // ?? QUY: N?u l? Semi, ??o ti?p v?i ???ng d?n m?i
        flatten(child, accumulatedYield, newPath);
      }
    });
  }
  
  // 4. B?t ??u ch?y
  // Chu?i kh?i ??u: T?n m?n ch?nh (V? d?: "Pizza H?i S?n")
  if (bomTree) {
      flatten(bomTree, 1.0, `(1) ${bomTree.name}`);
  }
  
  // 5. L?u flatList v?o Sheet 'DB_BOM_CACHE'
  // Ch? ?: H?m saveToSheet c?a ch?u c?n x? l? x?a d? li?u c? c?a productID n?y tr??c khi ghi m?i
  saveToSheet('DB_BOM_CACHE', flatList);
}

/**
 * ------------------------------------------------------------------------
 * [CORE] H?M T?I T?O CACHE BOM (?? FIX L?I 1.2 T?)
 * Ch?y h?m n?y t? menu "SYSTEM ADMIN" ?? l?m s?ch d? li?u.
 * ------------------------------------------------------------------------
 */
function regenerateAllBOMCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCache = ss.getSheetByName("DB_BOM_CACHE");
  const sheetProd = ss.getSheetByName("Danh m?c Product");
  
  if (!sheetCache || !sheetProd) {
    SpreadsheetApp.getUi().alert("? Thi?u Sheet Cache ho?c Danh m?c Product!");
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("? ?ang t?i b?n ?? c?ng th?c & T?n h?ng...", "System Admin");

  // 1. Load D? li?u
  const allRecipes = getRecipeMap(ss); 
  const nameMap = createNameMap(ss); // <--- [M?I] L?y t? ?i?n t?n
  
  const prodData = sheetProd.getDataRange().getValues();
  let cacheData = [];
  const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  SpreadsheetApp.getActiveSpreadsheet().toast("? ?ang t?nh to?n l? tr?nh (Trace Path)...", "System Admin");

  // 2. T?nh to?n
  for (let i = 1; i < prodData.length; i++) {
    let prodCode = String(prodData[i][0]).trim();
    if (!prodCode) continue;

    // Truy?n nameMap v?o h?m bung BOM
    let bomResult = explodeBOMForCache(prodCode, 1, allRecipes, nameMap); 
    
    for (let rawCode in bomResult) {
      let item = bomResult[rawCode];
      cacheData.push([
        prodCode,           
        rawCode,            
        item.qty,           
        item.path, // L?c n?y path ?? l? T?n -> T?n -> T?n         
        timeStr             
      ]);
    }
  }

  // 3. Ghi k?t qu?
  if (cacheData.length > 0) {
    sheetCache.getRange("A2:E").clearContent();
    sheetCache.getRange(2, 1, cacheData.length, 5).setValues(cacheData);
    SpreadsheetApp.getUi().alert(`? ?? c?p nh?t xong!\nKi?m tra c?t Trace_Path xem ?? hi?n T?n ch?a nh?.`);
  } else {
    SpreadsheetApp.getUi().alert("?? Kh?ng c? d? li?u ?? t?nh.");
  }
}

/**
 * [KAIZEN V3] H?M BUNG BOM AN TO?N (CH?NG V?NG L?P & TR?N B? NH?)
 * Update: Th?m c? ch? "Visited Stack" ?? ph?t hi?n v?ng l?p A -> B -> A
 */
function explodeBOMForCache(rootCode, demandQty, allRecipes, nameMap) {
  let results = {}; 

  const getName = (code) => nameMap[code] || code;
  // Format s?: 1,234.567 (B? b?t s? 0 th?a n?u c?n)
  const fmt = (num) => Number(num).toLocaleString('vi-VN', {maximumFractionDigits: 3});

  let rootName = getName(rootCode);
  
  // H?m ?? quy
  function traverse(currentCode, currentQty, history, visited = []) {
    if (visited.includes(currentCode)) return; // Ch?ng loop

    let recipe = allRecipes[currentCode];
    let currentName = getName(currentCode);

    // C?p nh?t l?ch s?
    let currentStep = { 
      name: currentName, 
      qty: currentQty,
      isSemi: (recipe && recipe.components && recipe.components.length > 0)
    };
    let newHistory = [...history, currentStep];

    // ?I?M CU?I (NVL ho?c Semi c?t)
    if (!currentStep.isSemi) {
      if (!results[currentCode]) {
        results[currentCode] = { qty: 0, branches: [] };
      }
      
      results[currentCode].qty += currentQty;
      
      // [FIX] T?O D?NG NH?NH (S? l??ng ??ng tr??c)
      let branchStr = newHistory.slice(1).map((step, index) => {
        let prefix = (index === 0) ? "   ?? " : " ? ";
        // KAIZEN: (SL) T?n
        return `${prefix}(${fmt(step.qty)}) ${step.name}`;
      }).join("");

      if (!results[currentCode].branches.includes(branchStr)) {
        results[currentCode].branches.push(branchStr);
      }
      return;
    }

    // BUNG TI?P
    let batchSize = recipe.batchOutput; 
    if (!batchSize || batchSize <= 0) batchSize = 1;
    let ratio = currentQty / batchSize;
    let newVisited = [...visited, currentCode];

    recipe.components.forEach(comp => {
      let childNeed = comp.qty * ratio; 
      traverse(comp.code, childNeed, newHistory, newVisited);
    });
  }

  // B?t ??u ch?y
  traverse(String(rootCode).trim(), demandQty, []);
  
  // T?NG H?P K?T QU?
  let finalResults = {};
  for (let code in results) {
    let item = results[code];
    
    // [FIX] D?ng ti?u ?? c?ng ??a s? l??ng l?n tr??c
    let treeView = `B?T ??U: (${fmt(demandQty)}) ${rootName}\n` + item.branches.join('\n');
    
    finalResults[code] = {
      qty: item.qty,
      path: treeView
    };
  }

  return finalResults;
}


// ==========================================
// KHU V?C 9: TRACEABILITY & SHADOW ENGINE
// ==========================================

/**
 * [V12.31 DUAL TRACE] H?M TRUY V?T TH?NG MINH 2 CHI?U
 * - Input: M? b?t k? (NVL ho?c Product)
 * - Logic: 
 * + N?u l? NVL -> T?m xem n? n?m trong M?n n?o? (Where Used)
 * + N?u l? Product -> T?m xem n? g?m nh?ng NVL n?o? (BOM Explosion)
 */
function getTraceDataFromCache(code) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = ss.getSheetByName(CONFIG.SHEET_BOM_CACHE);
    
    if (!sheetCache) return { success: false, message: "Ch?a c? Cache! H?y ch?y T?i t?o." };
    
    const data = sheetCache.getDataRange().getValues();
    // Cache Structure: [0]Product | [1]Raw | [2]Qty | [3]Path | [4]Time
    
    const target = String(code).trim();
    let results = [];
    let mode = ""; // ?? ghi log xem ?ang tra xu?i hay ng??c

    // 1. CHU?N B? MAP T?N (?? hi?n th? cho ??p)
    const masterMap = getMasterMap(ss);

    // 2. QU?T CACHE
    for (let i = 1; i < data.length; i++) {
      let prodCode = String(data[i][0]).trim();
      let rawCode = String(data[i][1]).trim();
      let qty = Number(data[i][2]) || 0;
      let path = String(data[i][3]);
      let time = data[i][4];

      // TR??NG H?P A: B?m v?o NVL (T?m M?n ?n s? d?ng)
      if (rawCode === target) {
        mode = "WHERE_USED";
        let prodInfo = masterMap[prodCode] || { name: prodCode };
        
        // L?y ?VT c?a c?i NVL ?ang soi (Target)
        let rawUnit = masterMap[rawCode] ? masterMap[rawCode].unit : ""; 

        results.push({
          code: prodCode,             
          name: prodInfo.name,        
          qty: qty,                   
          unit: rawUnit,              // <--- TH?M D?NG N?Y
          path: path,                 
          type: "D?ng trong"
        });
      }
      
      // TR??NG H?P B: B?m v?o Product (T?m th?nh ph?n NVL)
      else if (prodCode === target) {
        mode = "BOM_EXPLODE";
        let rawInfo = masterMap[rawCode] || { name: rawCode };
        
        // L?y ?VT c?a t?ng th?nh ph?n con
        let childUnit = rawInfo.unit || "";

        results.push({
          code: rawCode,              
          name: rawInfo.name,         
          qty: qty,                   
          unit: childUnit,            // <--- TH?M D?NG N?Y
          path: path,                 
          type: "Th?nh ph?n"
        });
      }
    }
    
    // S?p x?p: C?i n?o s? l??ng l?n l?n ??u
    results.sort((a, b) => b.qty - a.qty);

    if (results.length === 0) {
      return { success: false, message: `Kh?ng t?m th?y d? li?u Trace cho m? ${target} trong Cache.` };
    }

    return { success: true, data: results, mode: mode };
    
  } catch (e) {
    return { success: false, message: "L?i Trace: " + e.toString() };
  }
}

/**
 * [SHADOW ENGINE V12.29 - FULL AUDIT MODE]
 * T?nh to?n t?ng nhu c?u NVL d?a tr?n Doanh thu & Cache
 */
function runShadowAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSales = ss.getSheetByName("Data Doanh Thu");
  const sheetCache = ss.getSheetByName(CONFIG.SHEET_BOM_CACHE);
  let sheetResult = ss.getSheetByName("SHADOW_RESULT");

  if (!sheetSales || !sheetCache) return;

  // 1. T?O SHEET K?T QU?
  if (!sheetResult) {
    sheetResult = ss.insertSheet("SHADOW_RESULT");
    sheetResult.appendRow(["M? NVL", "T?n NVL (G?i ?)", "T?ng C?n (Shadow)", "?VT", "Chi ti?t truy v?t (Full Source)"]);
    sheetResult.getRange("A1:E1").setFontWeight("bold").setBackground("#d9ead3");
    sheetResult.setColumnWidth(2, 250);
    sheetResult.setColumnWidth(5, 400);
    sheetResult.setRowHeight(1, 40);
  } else {
    if(sheetResult.getLastRow() > 1) 
      sheetResult.getRange(2, 1, sheetResult.getLastRow()-1, 5).clearContent();
  }

  // 2. G?P DOANH THU
  const salesData = sheetSales.getDataRange().getValues();
  let salesMap = {}; 
  for (let i = 1; i < salesData.length; i++) {
    let code = String(salesData[i][3]).trim(); // Col D
    let qty = Number(salesData[i][5]) || 0;    // Col F
    if (code && qty > 0) salesMap[code] = (salesMap[code] || 0) + qty;
  }

  // 3. T?NH TO?N
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

  // 4. L?Y TH?NG TIN MASTER
  let masterInfo = getMasterMap(ss); 

  // 5. XU?T K?T QU?
  let outputRows = [];
  
  for (let raw in shadowResult) {
    let item = shadowResult[raw];
    let info = masterInfo[raw] || { name: "Unknown", unit: "" };
    
    item.sources.sort((a, b) => b.total - a.total);
    
    let traceInfo = item.sources.map(src => {
      return `[${src.prod}] (x${src.sales}) ? ${Math.round(src.total)}`;
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
    SpreadsheetApp.getUi().alert(`? ?? CH?Y AUDIT!\nT?ng m? NVL: ${outputRows.length}`);
  } else {
    SpreadsheetApp.getUi().alert("?? Kh?ng c? s? li?u t?nh to?n.");
  }
}

/**
 * [HELPER] MAP T?N & ?VT
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
      if (code && !map[code]) map[code] = { name: name, unit: "C?i" };
    }
  }
  return map;
}

/**
 * [KAIZEN V2] H?M ??C C?NG TH?C TO?N DI?N (BAO G?M C? PRODUCT & SEMI)
 * Update: ?? th?m ph?n ??c sheet 'BOM Product' ?? li?n k?t M?n ?n -> Semi
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

  // --- PH?N 1: ??C C?NG TH?C SEMI (3 Sheet B?p) ---
  const semiSheets = [CONFIG.SHEET_KIT_KITCHEN, CONFIG.SHEET_KIT_PIZZA, CONFIG.SHEET_KIT_SERVICE];
  
  semiSheets.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const data = s.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let r = data[i];
      let parentCode = clean(r[0]);         // C?t A: M? Cha
      let type = clean(r[1]).toLowerCase(); // C?t B: Lo?i
      let val = safeNum(r[4]);              // C?t E: SL/Yield

      if (!parentCode) continue;
      if (!recipes[parentCode]) recipes[parentCode] = { batchOutput: 1, components: [] };

      if (type === 'parent') {
        // L?y s?n l??ng m? n?u (Batch Output)
        recipes[parentCode].batchOutput = (val > 0) ? val : 1;
      } 
      else if (type === 'child') {
        let childCode = clean(r[2]); // C?t C: M? Con
        if (childCode) {
          recipes[parentCode].components.push({ code: childCode, qty: val });
        }
      }
    }
  });

  // --- PH?N 2: ??C C?NG TH?C PRODUCT (QUAN TR?NG: M?n -> Semi) ---
  // Ch? th?m ?o?n n?y ?? h? th?ng hi?u 1 Pizza g?m nh?ng g?
  const sheetBOM = ss.getSheetByName(CONFIG.SHEET_BOM_PRODUCT); // "BOM Product"
  if (sheetBOM) {
    const bomData = sheetBOM.getDataRange().getValues();
    // C?u tr?c file BOM Product c?a ch?u:
    // A: M? Parent (10000014) | B: Lo?i | C: M? Component (214) | E: SL (80)
    
    for (let i = 1; i < bomData.length; i++) {
      let r = bomData[i];
      let parentCode = clean(r[0]); // C?t A
      let childCode = clean(r[2]);  // C?t C
      let qty = safeNum(r[4]);      // C?t E

      if (!parentCode || !childCode) continue;

      // N?u ch?a c? trong danh s?ch c?ng th?c th? t?o m?i
      // V?i Product, Batch Output m?c ??nh lu?n l? 1 (1 c?i b?nh)
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
 * Kh?c ph?c tri?t ?? l?i ??t g?y d? li?u BOM
 */
  function buildPath(code, currentQty, currentLevel, currentPathNodes) {
    // a. Ch?n v?ng l?p v? t?n & Gi?i h?n ?? s?u (Max 10 t?ng)
    if (currentLevel > 10) {
       allPaths.push({ nodes: [...currentPathNodes, { qty: 0, name: "?? L?I: QU? NHI?U T?NG (LOOP)", level: currentLevel, type: "ERROR" }].reverse() });
       return;
    }

    const cleanCode = String(code).trim();
    const itemName = nameMap[cleanCode] || "M? " + cleanCode;
    // X?c ??nh lo?i h?ng t? Master Data (Ch?nh x?c h?n ch? d?a v?o vi?c c? BOM hay kh?ng)
    const masterType = typeMap.get(cleanCode) || "NVL"; 
    const hasRecipe = bomMap[cleanCode] ? true : false;
    
    // b. Format s? l??ng chu?n (Lu?n gi? 3 s? l? ?? kh?p v?i Frontend)
    const qtyStr = Number(currentQty).toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 3});

    const newNode = {
  // L?i #2: Chuy?n 'en-US' sang 'vi-VN' ?? d?ng d?u ch?m ph?n c?ch h?ng ngh?n
  qty: Number(currentQty).toLocaleString('vi-VN', {minimumFractionDigits: 0, maximumFractionDigits: 3}), 
  
  // L?i #3: ??i raw_qty th?nh qty_raw theo ??ng quy t?c Suffix "Dual State"
  qty_raw: Number(currentQty), 
  
  name: itemName,
  unit: item.unit || "Gr",
  level: currentLevel
};

    const newPath = [...currentPathNodes, newNode];

    // c. LOGIC QUY?T ??NH (CORE)
    if (hasRecipe) {
      // TR??NG H?P 1: C? c?ng th?c -> ??o ti?p
      bomMap[cleanCode].components.forEach(item => {
        let childQty = (currentQty * item.qty) / (bomMap[cleanCode].batchOutput || 1);
        buildPath(item.code, childQty, currentLevel + 1, newPath);
      });
    } else {
      // TR??NG H?P 2: Kh?ng c? c?ng th?c
      if (masterType === 'SEMI' || masterType === 'PROD') {
          // ?? R?I RO PH?T HI?N: L? Semi/Prod m? kh?ng c? BOM -> ??T G?Y!
          // Ghi nh?n d?ng l?i ?? Frontend hi?n th? m?u ??
          const errorNode = {
             qty: "MISSING",
             name: `?? C?NH B?O: ${itemName} (Ch?a khai b?o BOM)`,
             level: currentLevel + 1,
             type: "BROKEN" 
          };
          allPaths.push({ nodes: [...newPath, errorNode].reverse() });
      } else {
          // TR??NG H?P 3: L? NVL th?t s? -> ?i?m cu?i an to?n
          allPaths.push({ nodes: [...newPath].reverse() });
      }
    }
  }

/**
 * H?M 1: T?O T? ?I?N T?N (Name Mapping)
 * Gi?p ??i M? (100xxx) -> T?n m?n ?n/nguy?n li?u
 */
function createNameMap(ss) {
  let map = {};
  const clean = (val) => String(val || "").trim();

  // 1. Qu?t Danh m?c NVL
  const sheetNVL = ss.getSheetByName(CONFIG.SHEET_NVL);
  if (sheetNVL) {
    const data = sheetNVL.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let code = clean(data[i][1]);
      if (code) map[code] = clean(data[i][0]); 
    }
  }

  // 2. Qu?t Danh m?c Product
  const sheetProd = ss.getSheetByName(CONFIG.SHEET_PROD);
  if (sheetProd) {
    const data = sheetProd.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let code = clean(data[i][0]);
      if (code) map[code] = clean(data[i][1]);
    }
  }

  // 3. Qu?t 3 Sheet Semi (Kitchen, Pizza, Service)
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

/* [MODULE CACHE] C?P NH?T CACHE RI?NG L? (PARTIAL UPDATE) */
function updateSingleBOMCache(targetCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCache = ss.getSheetByName("DB_BOM_CACHE");
    if (!sheetCache) return;

    // 1. L?y c?u tr?c c?y m?i nh?t c?a m? h?ng n?y
    const newTreeData = getTraceDataFromCache(targetCode); // T?n d?ng h?m ph?n t?ch c?y c? s?n
    if (!newTreeData || !newTreeData.length) return;

    // 2. T?m v? tr? c? trong Sheet Cache ?? ghi ??
    const dataCache = sheetCache.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 0; i < dataCache.length; i++) {
      if (dataCache[i][0] === targetCode) {
        rowIndex = i + 1;
        break;
      }
    }

    const rowData = [targetCode, JSON.stringify(newTreeData), new Date()];

    // 3. Ghi ?? n?u ?? c?, ho?c th?m m?i n?u ch?a t?n t?i
    if (rowIndex > 0) {
      sheetCache.getRange(rowIndex, 1, 1, 3).setValues([rowData]);
    } else {
      sheetCache.appendRow(rowData);
    }
    console.log("?? c?p nh?t Cache cho m?: " + targetCode);
  } catch (e) {
    console.error("L?i UpdateSingleCache: " + e.toString());
  }
}

function bulkUpdateSemiCosts(){try{const ss=SpreadsheetApp.getActiveSpreadsheet(),sN=ss.getSheetByName(CONFIG.SHEET_NVL),sP=ss.getSheetByName(CONFIG.SHEET_PROD),sBp=ss.getSheetByName(CONFIG.SHEET_BOM_PRODUCT),costs=new Map();if(!sN||!sP||!sBp)return{success:false,message:"Thi?u Sheet Danh m?c/BOM!"};sN.getDataRange().getValues().forEach((r,i)=>{if(i>0)costs.set(String(r[1]).trim(),Number(r[3])||0)});const semiSheets=[CONFIG.SHEET_KIT_KITCHEN,CONFIG.SHEET_KIT_PIZZA,CONFIG.SHEET_KIT_SERVICE];semiSheets.forEach(name=>{const s=ss.getSheetByName(name);if(!s)return;const d=s.getDataRange().getValues();let pIdx=-1,pTot=0,pY=1;for(let i=0;i<d.length;i++){const type=String(d[i][1]).trim(),code=String(d[i][2]).trim();if(type==='Parent'){if(pIdx!==-1){d[pIdx][7]=Math.round(pTot);d[pIdx][6]=Number((pY>0?pTot/pY:0).toFixed(3));costs.set(String(d[pIdx][0]).trim(),d[pIdx][6])}pIdx=i;pTot=0;pY=Number(d[i][4])||1}else if(type==='Child'){const q=Number(d[i][4])||0,uP=costs.get(code)||0,lT=uP*q;d[i][6]=lT;pTot+=lT}}if(pIdx!==-1){d[pIdx][7]=Math.round(pTot);d[pIdx][6]=Number((pY>0?pTot/pY:0).toFixed(3));costs.set(String(d[pIdx][0]).trim(),d[pIdx][6])}s.getDataRange().setValues(d)});const bD=sBp.getDataRange().getValues(),pCosts=new Map();let pC="",pT=0;for(let i=1;i<bD.length;i++){const r=bD[i],type=String(r[1]).toLowerCase();if(type==='parent'){if(pC)pCosts.set(pC,pT);pC=String(r[0]).trim();pT=0}else if(type==='child'){pT+=(costs.get(String(r[2]).trim())||0)*(Number(r[4])||0)}}if(pC)pCosts.set(pC,pT);const pData=sP.getDataRange().getValues();for(let i=1;i<pData.length;i++){const c=String(pData[i][0]).trim();if(pCosts.has(c))pData[i][6]=Math.round(pCosts.get(c))}sP.getDataRange().setValues(pData);regenerateAllBOMCache();return{success:true,message:"? Domino th?nh c?ng: Gi? ?? c?p nh?t t? G?c ??n Ng?n!"}}catch(e){return{success:false,message:"L?i: "+e.toString()}}}

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

    // 1. Qu?t Sheet H?y NVL & SEMI 
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

    // 2. Qu?t Sheet H?y PRODUCT 
    const sPr = ss.getSheetByName(CONFIG.SHEET_SPOILAGE_PROD);
    if (sPr && sPr.getLastRow() > 1) {
      sPr.getRange(2, 1, sPr.getLastRow() - 1, 12).getValues().forEach(r => {
        const d = parseDateVN(r[0]);
        if (!d || (fDate && d < fDate) || (tDate && d > tDate)) return;
        const q = Number(r[4]), c = Number(r[11]) || 0;
        results.push({
          date: r[0] instanceof Date ? formatDate(r[0]) : r[0],
          code: String(r[2]), name: String(r[1]), qty: q,
          unit: 'C?i', note: String(r[5]), dept: String(r[6]),
          cost: c, amount: Math.round(q * c), type: 'PROD'
        });
      });
    }
    return { success: true, data: results.sort((a,b) => parseDateVN(b.date) - parseDateVN(a.date)) };
  } catch (e) { return { success: false, message: "L?i: " + e.toString() }; }
}
/**
 * XU?T EXCEL FORMAT SAP - PHI?N B?N HO?N CH?NH
 * - Progress tracking
 * - Data validation
 * - Format ItemCode chu?n (b? s? th?p ph?n)
 * - T?ch data l?i sang c?t ph?
 * - Freeze header
 */
function exportSpoilageToExcel(data, config) {
  try {
    Logger.log('? [BACKEND] exportSpoilageToExcel called');
    Logger.log('? [BACKEND] Data count: ' + data.length);
    
    const cache = CacheService.getScriptCache();
    cache.put('EXPORT_PROGRESS', '?ang ph?n t?ch d? li?u...', 120);
    
    // T?o Spreadsheet t?m
    const tempSS = SpreadsheetApp.create('TEMP_SPOILAGE_' + new Date().getTime());
    const tempSheet = tempSS.getSheets()[0];
    tempSheet.setName('Sheet1');
    
    // ? HEADER (4 c?t SAP + C?t ph?)
    const headers = ['ItemCode', 'ItemName', 'UomCode', 'Quantity'];
    const extraHeaders = ['OriginalCode', 'OriginalQty', 'OriginalUnit', 'Dept', 'Note', 'Date', 'Amount', 'ErrorReason'];
    tempSheet.appendRow([...headers, ...extraHeaders]);
    
    // ? X? L? DATA
    let validRows = [];
    let errorRows = [];
    let processed = 0;
    
    cache.put('EXPORT_PROGRESS', '?ang x? l? d? li?u... 0/' + data.length, 120);
    
    data.forEach(function(item, index) {
      processed++;
      
      // Update progress m?i 10 d?ng
      if (processed % 10 === 0) {
        cache.put('EXPORT_PROGRESS', '?ang x? l?... ' + processed + '/' + data.length, 120);
      }
      
      // [FIX 1] Format ItemCode - B? s? th?p ph?n
      let cleanCode = String(item.code || '').trim();
      cleanCode = cleanCode.replace(/\.0+$/, ''); // 1402194.0 ? 1402194
      cleanCode = cleanCode.replace(/,/g, ''); // B? d?u ph?y
      
      // [FIX 2] Validation
      let errorReason = '';
      let isValid = true;
      
      if (!cleanCode || cleanCode === '') {
        errorReason = 'Thi?u ItemCode';
        isValid = false;
      }
      
      if (!item.unit || String(item.unit).trim() === '') {
        errorReason += (errorReason ? '; ' : '') + 'Thi?u UomCode';
        isValid = false;
      }
      
      if (item.qty <= 0) {
        errorReason += (errorReason ? '; ' : '') + 'S? l??ng <= 0';
        isValid = false;
      }
      
      // T?o row data
      const rowData = {
        code: cleanCode,
        name: item.name,
        unit: String(item.unit || '').trim(),
        qty: Math.round(item.qty * 1000) / 1000,
        dept: item.dept || '',
        note: item.note || '',
        date: item.date || '',
        amount: Math.round(item.amount || 0),
        originalCode: item.code, // Gi? nguy?n code g?c
        originalQty: item.qty,
        originalUnit: item.unit,
        error: errorReason
      };
      
      if (isValid) {
        validRows.push(rowData);
      } else {
        errorRows.push(rowData);
      }
    });
    
    cache.put('EXPORT_PROGRESS', '?ang ghi d? li?u v?o file...', 120);
    
    // ? GHI DATA H?P L? (4 C?T ??U)
    validRows.forEach(function(row) {
      tempSheet.appendRow([
        row.code,        // ItemCode
        row.name,        // ItemName
        row.unit,        // UomCode
        row.qty,         // Quantity
        row.originalCode, // C?t ph?: Original Code
        row.originalQty,  // C?t ph?: Original Qty
        row.originalUnit, // C?t ph?: Original Unit
        row.dept,
        row.note,
        row.date,
        row.amount,
        '' // Error (tr?ng v? h?p l?)
      ]);
    });
    
    // ? GHI DATA L?I (?? tr?ng 4 c?t ??u, data ? c?t ph?)
    errorRows.forEach(function(row) {
      tempSheet.appendRow([
        '',              // ItemCode (TR?NG)
        '',              // ItemName (TR?NG)
        '',              // UomCode (TR?NG)
        '',              // Quantity (TR?NG)
        row.originalCode, // C?t ph?: Original Code
        row.originalQty,
        row.originalUnit,
        row.dept,
        row.note,
        row.date,
        row.amount,
        row.error // L? do l?i
      ]);
    });
    
    cache.put('EXPORT_PROGRESS', '?ang ??nh d?ng file...', 120);
    
    // ? FORMAT HEADER
    const headerRange = tempSheet.getRange(1, 1, 1, headers.length + extraHeaders.length);
    headerRange.setBackground('#4CAF50');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // [FIX 4] FREEZE HEADER
    tempSheet.setFrozenRows(1);
    
    // Format 4 c?t SAP (background xanh nh?t)
    const sapColsRange = tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, 4);
    sapColsRange.setBackground('#E8F5E9');
    
    // Format c?t l?i (background ?? nh?t)
    if (errorRows.length > 0) {
      const errorStartRow = validRows.length + 2; // +2 v? c? header
      const errorRange = tempSheet.getRange(errorStartRow, 1, errorRows.length, headers.length + extraHeaders.length);
      errorRange.setBackground('#FFEBEE');
    }
    
    // Format s?
    const qtyCol = tempSheet.getRange(2, 4, tempSheet.getLastRow() - 1, 1);
    qtyCol.setNumberFormat('#,##0.000');
    
    // Auto-resize
    for (let i = 1; i <= headers.length + extraHeaders.length; i++) {
      tempSheet.autoResizeColumn(i);
    }
    
    cache.put('EXPORT_PROGRESS', '?ang t?o file Excel...', 120);
    
    // ? XU?T FILE
    const dateStr = config.fromDate ? config.fromDate.replace(/-/g, '') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    const fileName = 'HUY_HANG_SAP_' + dateStr + '_' + config.unitMode + (config.explodeBOM ? '_BUNG_BOM' : '') + '.xlsx';
    
    const folderId = '1OAiOKp4HbLIXDKzlo4IkOrGFX2wjCvuM';
    const folder = DriveApp.getFolderById(folderId);
    
    const url = 'https://docs.google.com/feeds/download/spreadsheets/Export?key=' + tempSS.getId() + '&exportFormat=xlsx';
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    const blob = response.getBlob().setName(fileName);
    const file = folder.createFile(blob);
    
    // X?a Spreadsheet t?m
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
    
    // X?a progress
    cache.remove('EXPORT_PROGRESS');
    
    Logger.log('? [BACKEND] File created: ' + fileName);
    
    const summary = validRows.length + ' d?ng h?p l?' + (errorRows.length > 0 ? ', ' + errorRows.length + ' d?ng l?i (?? t?ch sang ph?i)' : '');
    
    return {
      success: true,
      url: file.getUrl(),
      message: '? ?? t?o ' + summary
    };
    
  } catch (e) {
    Logger.log('? [BACKEND] Error: ' + e.toString());
    const cache = CacheService.getScriptCache();
    cache.remove('EXPORT_PROGRESS');
    return { success: false, message: 'L?i: ' + e.toString() };
  }
}

/**
 * [M?I] H?M L?Y TI?N ??
 */
function getExportProgress() {
  const cache = CacheService.getScriptCache();
  return cache.get('EXPORT_PROGRESS') || '';
}

/**
 * [GIAI ?O?N 1 - FIX 2] T?NH BOM H?Y H?NG THEO S? L??NG
 */
function calculateSpoilageBOM(itemCode, spoiledQty, spoiledUnit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // L?y Recipe Map
    const recipeMap = getRecipeMap(ss);
    const masterMap = new Map();
    
    // Load Master Data
    const sysData = getSystemData();
    sysData.masterData.forEach(m => masterMap.set(String(m.code).trim(), m));
    
    // L?y BOM c?a m?n h?y
    const recipe = recipeMap[String(itemCode).trim()];
    if (!recipe || !recipe.components || recipe.components.length === 0) {
      return { success: false, message: 'M? n?y kh?ng c? BOM' };
    }
    
    // T?nh ratio
    const batchOutput = recipe.batchOutput || 1;
    const ratio = spoiledQty / batchOutput;
    
    // T?nh chi ti?t
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
    return { success: false, message: 'L?i: ' + e.toString() };
  }
}
}

