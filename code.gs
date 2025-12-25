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

// Last sync: 2025-12-25T04:08:50.282Z

// ============ getSystemData ============
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

// ============ saveSpoilageData ============
function saveSpoilageData(p){try{const ss=SpreadsheetApp.getActiveSpreadsheet(),sSp=ss.getSheetByName(CONFIG.SHEET_SPOILAGE),sSm=ss.getSheetByName(CONFIG.SHEET_SPOILAGE_SEMI),sPr=ss.getSheetByName(CONFIG.SHEET_SPOILAGE_PROD);if(!sSp||!sSm||!sPr)return{success:false,message:"Thi?u Sheet H?y!"};const sys=getSystemData(),mMap=new Map(),rMap=getRecipeMap(ss),nMap=createNameMap(ss);sys.masterData.forEach(m=>mMap.set(String(m.code).trim(),m));let d=p.date;if(d.includes("-")){let x=d.split('-');d=`${x[2]}/${x[1]}/${x[0]}`}const rSp=[],rSm=[],rPr=[];p.items.forEach(i=>{const c=String(i.code).trim(),mI=mMap.get(c),uP=mI?Number(mI.cost)||0:0,amt=Math.round(i.qty*uP);if(i.itemType==='PROD'){rPr.push([d,i.name,i.code,"",i.qty,i.note||"",p.dept,"","",mI?mI.category:"",mI?mI.class:"",uP])}else if(i.itemType==='SEMI'){const leaves=bomEngine(c,i.qty,rMap,mMap);const isBk=(leaves.length===1&&String(leaves[0].code).trim()===c);let wn=isBk?" | ?? CH?A BOM":"";rSp.push([d,c,i.name,"",i.unit,i.factor||1,"",i.unit,i.qty,"",(i.note||"H?y BTP")+wn,p.dept,"",uP,amt]);if(!isBk){const cons={};leaves.forEach(l=>{cons[l.code]=(cons[l.code]||0)+l.qty_raw});for(const[lC,lQ]of Object.entries(cons)){rSm.push([d,lC,"","","","",smartRound(lQ),"","","",`Bung t? ${i.qty} ${i.name}`,p.dept])}}}else{rSp.push([d,c,i.name,"",i.unit,i.factor||1,"",i.unit,i.qty,"",(i.note||""),p.dept,"",uP,amt])}});if(rSp.length>0)sSp.getRange(findNextEmptyRow(sSp),1,rSp.length,15).setValues(rSp);if(rSm.length>0)sSm.getRange(findNextEmptyRow(sSm),1,rSm.length,12).setValues(rSm);if(rPr.length>0)sPr.getRange(findNextEmptyRow(sPr),1,rPr.length,12).setValues(rPr);return{success:true,message:"? ?? l?u phi?u H?y th?nh c?ng & Kh?p c?t!"}}catch(e){return{success:false,message:"L?i: "+e.toString()}}}

// ============ saveTransferData ============
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

// ============ updateTransferFull ============
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

// ============ updateMasterData ============
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

// ============ deleteMasterData ============
function deleteMasterData(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName = (payload.type === 'PROD') ? CONFIG.SHEET_PROD : CONFIG.SHEET_NVL;
    const sheet = ss.getSheetByName(sheetName);
    if (payload.rowId) { sheet.deleteRow(payload.rowId); return { success: true, message: "?? x?a m?: " + payload.code }; }
    return { success: false, message: "Kh?ng t?m th?y ID!" };
  }

// ============ updateSystemData ============
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

// ============ getBOMDetail ============
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

// ============ saveBOM ============
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
}

// ============ traceIngredientUsage ============
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

