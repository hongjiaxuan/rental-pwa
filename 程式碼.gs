const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PHOTO_FOLDER_ID = '1BQG17fTrias0NQjaR4zJAk36p-zi4UfR'; // ★重要：填入資料夾ID★

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('套房管理系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ══════════════════════════════════════════════════════════════
// doPost — GitHub Pages / PWA 請求路由層
// 接收 JSON：{ fn: "functionName", payload: ... }
// 將呼叫路由到對應的 GAS 函數，回傳 JSON 結果
// ══════════════════════════════════════════════════════════════
function doPost(e) {
  // CORS headers（允許 GitHub Pages 域名呼叫）
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const body = JSON.parse(e.postData.contents);
    const fn   = String(body.fn   || '');
    const payload = body.payload;

    // 函數白名單（安全性：只允許呼叫已知函數）
    const ALLOWED = {
      getPendingRooms:              () => getPendingRooms(),
      getDashboardSummary:          () => getDashboardSummary(),
      getReconciliationData:        () => getReconciliationData(),
      getAllRoomsInfo:               () => getAllRoomsInfo(),
      getRentedRooms:               () => getRentedRooms(),
      getPendingRefunds:            () => getPendingRefunds(),
      getSettings:                  () => getSettings(),
      submitMeterReading:           () => submitMeterReading(payload),
      confirmPayment:               () => confirmPayment(payload),
      updateBillDetails:            () => updateBillDetails(payload),
      saveTenantInfo:               () => saveTenantInfo(payload),
      renewContract:                () => renewContract(payload),
      saveTaipowerBill:             () => saveTaipowerBill(payload),
      saveSettings:                 () => saveSettings(payload),
      getRoomCheckoutData:          () => getRoomCheckoutData(payload),
      processMoveOut:               () => processMoveOut(payload),
      getHistoricalBills:           () => getHistoricalBills(payload),
      confirmTpFinalRefundBackend:   () => confirmTpFinalRefundBackend(payload),
      confirmDepositRefund:         () => confirmDepositRefund(payload),
      generateTaipowerFinalBill:    () => generateTaipowerFinalBill(payload),
      distributePublicElectricity:  () => distributePublicElectricity(payload),
    };

    if (!ALLOWED[fn]) {
      output.setContent(JSON.stringify({ __gasError: '不允許的操作：' + fn }));
      return output;
    }

    const result = ALLOWED[fn]();
    output.setContent(JSON.stringify(result));

  } catch(err) {
    output.setContent(JSON.stringify({ __gasError: err.message }));
  }

  return output;
}

// === 1. 抄表模組 ===
function getPendingRooms() {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Rooms').getDataRange().getValues();
  let rooms = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === '出租中') {
      rooms.push({
        roomId: String(data[i][0]), tenant: String(data[i][1]), elecRate: Number(data[i][2]) || 6.5,
        waterFee: Number(data[i][3]) || 0, lastReading: Number(data[i][4]) || 0, rentAmount: Number(data[i][6]) || 0,
        deposit: Number(data[i][7]) || 0, endDate: data[i][8] ? Utilities.formatDate(new Date(data[i][8]), "GMT+8", "yyyy-MM-dd") : '未填寫',
        isPrepaid: String(data[i][9] || '').toUpperCase() === 'Y', balance: Number(data[i][10]) || 0
      });
    }
  } return rooms;
}

// ==========================================
// 一鍵續約引擎
// ==========================================
function renewContract(payload) {
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(5000)) throw new Error("系統忙碌");
  try {
    const rSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Rooms');
    const data = rSheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(String(data[i][0]) === payload.roomId) {
        const oldEndDate = data[i][8] ? Utilities.formatDate(new Date(data[i][8]), "GMT+8", "yyyy-MM-dd") : '未填寫';
        const oldNotes = String(data[i][11] || '');
        const newNote = `[系統] ${Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")} 續約完成 (原到期日: ${oldEndDate})\n`;

        rSheet.getRange(i+1, 7).setValue(payload.newRent);    // G欄: 租金
        rSheet.getRange(i+1, 9).setValue(payload.newEndDate); // I欄: 到期日
        rSheet.getRange(i+1, 12).setValue(newNote + oldNotes); // L欄: 修繕與備註紀錄 (將歷史封存)

        return {success: true, message: "續約成功！合約展延與租金已更新。"};
      }
    }
    throw new Error("找不到房號");
  } catch(e) { return {success: false, message: e.message}; }
  finally { lock.releaseLock(); }
}

// === ★ 核心升級：雙錢包智慧扣抵引擎 (含照片上傳修復) ★ ===
// === 修正後的抄表寫入邏輯 (程式碼.gs) ===
function submitMeterReading(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error("系統忙碌中");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rSheet = ss.getSheetByName('Rooms');
    const rData = rSheet.getDataRange().getValues();
    
    let tRoom = null, tIndex = -1;
    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][0]) === String(data.roomId)) { 
        tRoom = rData[i];
        tIndex = i + 1; break; 
      }
    }
    if (!tRoom) throw new Error("找不到房號資料");
    const now = new Date(); 
    const currentYM = Utilities.formatDate(now, "GMT+8", "yyyy-MM");
    
    const baseRent = Number(tRoom[6]) || 0;
    const baseWater = Number(tRoom[3]) || 0;
    const payCycle = String(tRoom[12] || '每月');
    const nextRentMonth = String(tRoom[13] || currentYM);
    const pendingPublicElec = Number(tRoom[14]) || 0; 
    const prepaidRentBal = Number(tRoom[10]) || 0; 
    const prevBillBal = Number(tRoom[15]) || 0;

    let multiplier = 1;
    if (payCycle === '半年') multiplier = 6;
    if (payCycle === '年收') multiplier = 12;

    let actualRent = 0, actualWater = 0;
    
    // 核心修正：不管是不是收租月，寫入資料庫的期間字串固定為 YYYY-MM
    let periodStr = currentYM; 

    if (currentYM >= nextRentMonth) {
        actualRent = baseRent * multiplier;
        actualWater = baseWater * multiplier;
        const nextDate = new Date(now.getFullYear(), now.getMonth() + multiplier, 1);
        const newNextRentMonth = Utilities.formatDate(nextDate, "GMT+8", "yyyy-MM");
        rSheet.getRange(tIndex, 14).setValue(newNextRentMonth);
    }

    let rentDeduct = 0;
    let newPrepaidRentBal = prepaidRentBal;
    if (actualRent > 0 && prepaidRentBal > 0) {
        rentDeduct = Math.min(actualRent, prepaidRentBal);
        newPrepaidRentBal -= rentDeduct;
    }

    const usage = Math.round((Number(data.currentReading) - Number(data.lastReading)) * 10) / 10;
    if (usage < 0) throw new Error("度數異常！");
    const elecFee = Math.round(usage * Number(data.elecRate));
    const subtotal = (actualRent - rentDeduct) + actualWater + elecFee + pendingPublicElec;

    let deductBillAmt = 0;
    let newBillBal = prevBillBal;
    let finalTotal = subtotal;

    if (newBillBal > 0) {
        if (newBillBal >= finalTotal) {
            deductBillAmt = finalTotal;
            newBillBal -= finalTotal;
            finalTotal = 0;
        } else {
            deductBillAmt = newBillBal;
            finalTotal -= newBillBal;
            newBillBal = 0;
        }
    } else if (newBillBal < 0) {
        finalTotal += Math.abs(newBillBal);
        deductBillAmt = newBillBal; 
        newBillBal = 0;
    }

    let billStatus = finalTotal === 0 ? '已收' : '未收';
    let payDate = finalTotal === 0 ? Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd") : '';

    let imageUrl = "";
    if (data.imageBase64) {
      try {
        const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
        const base64Data = data.imageBase64.split(',')[1] || data.imageBase64;
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', `${data.roomId}_${currentYM}_抄表.jpg`);
        imageUrl = folder.createFile(blob).getUrl();
      } catch (imgErr) {
        console.log("圖片儲存失敗: " + imgErr.message);
      }
    }

    const meterSheet = ss.getSheetByName('MeterReadings');
    if (meterSheet) {
      meterSheet.appendRow([
        "R" + now.getTime(), 
        Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"), 
        data.roomId, 
        periodStr, 
        data.currentReading, 
        imageUrl, 
        Session.getActiveUser().getEmail()
      ]);
    }

    const billSheet = ss.getSheetByName('Billings');
    if (billSheet) {
      billSheet.appendRow([
          "B" + now.getTime(), now, data.roomId, periodStr, 
          actualRent, elecFee, actualWater, finalTotal, 
          billStatus, payDate ? (payDate + " 00:00") : '', tRoom[1], data.lastReading, data.currentReading, 
          usage, data.elecRate, tRoom[8], finalTotal === 0 ? subtotal : '', payDate, deductBillAmt,
          pendingPublicElec, pendingPublicElec > 0 ? "代收公電費" : "", rentDeduct
      ]);
    }

    rSheet.getRange(tIndex, 5).setValue(data.currentReading);
    rSheet.getRange(tIndex, 11).setValue(newPrepaidRentBal);
    if (rSheet.getMaxColumns() < 16) rSheet.insertColumnsAfter(rSheet.getMaxColumns(), 16 - rSheet.getMaxColumns());
    rSheet.getRange(tIndex, 16).setValue(newBillBal);
    rSheet.getRange(tIndex, 15).clearContent(); 

    return { success: true, message: `${data.roomId} 抄表完成！` };
  } catch (e) { 
    return { success: false, message: e.message }; 
  } finally { 
    lock.releaseLock();
  }
}

function generateInitialBill(ss, p) {
  const bSheet = ss.getSheetByName('Billings');
  const rSheet = ss.getSheetByName('Rooms');
  const now = new Date();
  
  let multiplier = 1;
  if (p.payCycle === '半年') multiplier = 6;
  if (p.payCycle === '年收') multiplier = 12;

  const endDateObj = new Date(now.getFullYear(), now.getMonth() + multiplier, now.getDate());
  const periodStr = `${now.getMonth() + 1}/${now.getDate()} ~ ${endDateObj.getMonth() + 1}/${endDateObj.getDate()} (${p.payCycle})`;

  const initialRent = (Number(p.rent) || 0) * multiplier;
  const initialWater = (Number(p.waterFee) || 0) * multiplier;
  const deposit = Number(p.deposit) || 0; 
  const prepaidRentBal = Number(p.balance) || 0; 
  
  let rentDeduct = 0;
  let newPrepaidRentBal = prepaidRentBal;
  if (initialRent > 0 && prepaidRentBal > 0) {
      rentDeduct = Math.min(initialRent, prepaidRentBal);
      newPrepaidRentBal -= rentDeduct;
  }

  const subtotal = (initialRent - rentDeduct) + initialWater + deposit;
  const finalTotal = subtotal;

  bSheet.appendRow([
    "B" + now.getTime(), now, p.roomId, periodStr, 
    initialRent, 0, initialWater, finalTotal, 
    finalTotal === 0 ? '已收' : '未收', finalTotal === 0 ? (Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd") + " 00:00") : '', p.tenant, 
    p.lastReading || 0, p.lastReading || 0, 0, p.elecRate, p.endDate, 
    finalTotal === 0 ? subtotal : '', finalTotal === 0 ? Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd") : '', 0, 
    deposit, deposit > 0 ? "首期押金" : "", rentDeduct
  ]);
  
  const rData = rSheet.getDataRange().getValues();
  for (let i = 1; i < rData.length; i++) {
    if (String(rData[i][0]) === String(p.roomId)) {
      rSheet.getRange(i + 1, 11).setValue(newPrepaidRentBal); break;
    }
  }
}

function confirmPayment(payload) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) throw new Error("系統忙碌");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const bSheet = ss.getSheetByName('Billings');
    const bData = bSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][0]) === payload.billId) {
        const diff = Number(payload.actualAmount) - Number(bData[i][7]);
        bSheet.getRange(i+1, 9).setValue('已收');
        bSheet.getRange(i+1, 10).setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm"));
        bSheet.getRange(i+1, 17).setValue(payload.actualAmount);
        bSheet.getRange(i+1, 18).setValue(payload.payDate);

        if (diff !== 0) {
          const rSheet = ss.getSheetByName('Rooms');
          const rData = rSheet.getDataRange().getValues();
          if (rSheet.getMaxColumns() < 16) rSheet.insertColumnsAfter(rSheet.getMaxColumns(), 16 - rSheet.getMaxColumns());
          for (let r = 1; r < rData.length; r++) {
            if (String(rData[r][0]) === String(bData[i][2])) {
              // 💡 溢繳或少繳金額，寫入 P 欄 (帳款結餘)
              rSheet.getRange(r+1, 16).setValue((Number(rData[r][15]) || 0) + diff);
              break;
            }
          }
        }
        return { success: true, message: diff === 0 ? '入帳成功！' : `入帳成功！差額 $${diff} 已轉入帳款結餘。` };
      }
    } throw new Error("找不到帳單");
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function getReconciliationData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const bData = ss.getSheetByName('Billings').getDataRange().getValues();
    let bills = [];
    for (let i = bData.length - 1; i >= 1; i--) {
      if (String(bData[i][0]).trim() !== '' && String(bData[i][8]).trim() === '未收') {
        bills.push({
          billId: String(bData[i][0]), roomId: String(bData[i][2]), 
          yearMonth: (bData[i][3] instanceof Date) ? Utilities.formatDate(bData[i][3], "GMT+8", "yyyy-MM") : String(bData[i][3] || ''),
          rent: Number(bData[i][4])||0, elecFee: Number(bData[i][5])||0, waterFee: Number(bData[i][6])||0, totalAmount: Number(bData[i][7])||0,
          tenant: String(bData[i][10]), lastReading: Number(bData[i][11]), currentReading: Number(bData[i][12]), usage: Number(bData[i][13]), elecRate: Number(bData[i][14]), 
          endDate: (bData[i][15] instanceof Date) ? Utilities.formatDate(bData[i][15], "GMT+8", "yyyy-MM-dd") : String(bData[i][15] || ''), 
          prevBalance: Number(bData[i][18])||0, extraFee: Number(bData[i][19])||0, extraNote: String(bData[i][20]||''),
          rentDeduct: Number(bData[i][21])||0 // 讀取預繳租金扣抵額
        });
      }
    }
    const stData = ss.getSheetByName('Settlements').getDataRange().getValues();
    let deposits = [];
    for (let i = stData.length - 1; i >= 1; i--) {
      let status = String(stData[i][10]);
      if (status === '待退押金' || status === '待退台電款') {
        deposits.push({ 
          stId: String(stData[i][0]), roomId: String(stData[i][1]), tenant: String(stData[i][2]), moveOutDate: (stData[i][3] instanceof Date) ? Utilities.formatDate(stData[i][3], "GMT+8", "yyyy-MM-dd") : String(stData[i][3]), 
          refundAmount: Number(stData[i][9])||0, type: status, finalReading: Number(stData[i][4])||0, deposit: Number(stData[i][5])||0,
          finalElecFee: Number(stData[i][6])||0, tpRefund: Number(stData[i][7])||0, totalDeductions: Number(stData[i][8])||0, usage: Number(stData[i][11])||0
        });
      }
    }
    return { bills: bills, deposits: deposits };
  } catch (err) { throw new Error(err.message); }
}

function updateBillDetails(payload) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) throw new Error("系統忙碌中");
  try {
    const bSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Billings');
    const bData = bSheet.getDataRange().getValues();
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][0]) === payload.billId) {
        let lastReading = Number(bData[i][11]), elecRate = Number(bData[i][14]), prevBal = Number(bData[i][18]) || 0;
        let newCurRead = Number(payload.currentReading);
        let newUsage = Math.round((newCurRead - lastReading) * 10) / 10;
        if (newUsage < 0) throw new Error("度數不可小於上期！");
        let newElecFee = Math.round(newUsage * elecRate);
        
        let rent = Number(payload.rent) || 0, water = Number(payload.waterFee) || 0, extra = Number(payload.extraFee) || 0;
        let rentDeduct = Number(bData[i][21]) || 0;
        if (rentDeduct > rent) rentDeduct = rent; // 防呆：扣抵額不能大於修改後的租金
        
        let newTotal = (rent - rentDeduct) + newElecFee + water - prevBal + extra;
        if(newTotal < 0) newTotal = 0;

        bSheet.getRange(i+1, 5).setValue(rent);
        bSheet.getRange(i+1, 6).setValue(newElecFee);
        bSheet.getRange(i+1, 7).setValue(water);
        bSheet.getRange(i+1, 8).setValue(newTotal);
        bSheet.getRange(i+1, 13).setValue(newCurRead);
        bSheet.getRange(i+1, 14).setValue(newUsage);
        bSheet.getRange(i+1, 20).setValue(extra);
        bSheet.getRange(i+1, 21).setValue(payload.extraNote || '');
        bSheet.getRange(i+1, 22).setValue(rentDeduct);
        
        SpreadsheetApp.flush();
        return { success: true, message: '帳單修改成功！總金額已重算。' };
      }
    }
    throw new Error('找不到該帳單');
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function getAllRoomsInfo() {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Rooms').getDataRange().getValues();
  let rooms = [];
  for (let i = 1; i < data.length; i++) {
    rooms.push({
      roomId: String(data[i][0]), tenant: String(data[i][1]), elecRate: Number(data[i][2])||0, waterFee: Number(data[i][3])||0,
      status: String(data[i][5]), rent: Number(data[i][6])||0, deposit: Number(data[i][7])||0,
      endDate: data[i][8] ? Utilities.formatDate(new Date(data[i][8]), "GMT+8", "yyyy-MM-dd") : '未填寫',
      balance: Number(data[i][10]) || 0,     // K欄: 預繳租金
      notes: String(data[i][11] || ''),
      payCycle: String(data[i][12] || '每月'),
      billBalance: Number(data[i][15]) || 0  // P欄: 帳款結餘
    });
  } return rooms;
}

function saveTenantInfo(payload) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) throw new Error("系統忙碌");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rSheet = ss.getSheetByName('Rooms');
    const data = rSheet.getDataRange().getValues();
    if (rSheet.getMaxColumns() < 16) rSheet.insertColumnsAfter(rSheet.getMaxColumns(), 16 - rSheet.getMaxColumns());

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === payload.roomId) {
        const oldStatus = String(data[i][5]);
        let rowValues = data[i];
        rowValues[1] = payload.tenant; rowValues[2] = Number(payload.elecRate) || 0; rowValues[3] = Number(payload.waterFee) || 0;
        rowValues[5] = payload.status; rowValues[6] = Number(payload.rent) || 0; rowValues[7] = Number(payload.deposit) || 0;
        rowValues[8] = payload.endDate; rowValues[10] = Number(payload.balance) || 0; rowValues[11] = payload.notes || '';
        rowValues[12] = payload.payCycle || '每月';
        
        while(rowValues.length < 16) rowValues.push("");
        rowValues[15] = Number(payload.billBalance) || 0;

        let isNewTenant = (oldStatus === '空房' && payload.status === '出租中');
        if (isNewTenant) {
          const now = new Date(); let multiplier = 1;
          if (payload.payCycle === '半年') multiplier = 6; if (payload.payCycle === '年收') multiplier = 12;
          const nextDate = new Date(now.getFullYear(), now.getMonth() + multiplier, 1);
          rowValues[13] = Utilities.formatDate(nextDate, "GMT+8", "yyyy-MM");
        } else if (!rowValues[13]) {
          rowValues[13] = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM");
        }

        rSheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
        if (isNewTenant) generateInitialBill(ss, payload);
        SpreadsheetApp.flush();
        return { success: true, message: '設定與修繕紀錄已儲存！' };
      }
    } throw new Error('找不到房號');
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function getHistoricalBills(roomId) {
  const bData = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Billings').getDataRange().getValues();
  let history = [];
  for (let i = bData.length - 1; i >= 1; i--) {
    if (String(bData[i][2]) === roomId) {
      history.push({
        billId: String(bData[i][0]), yearMonth: String(bData[i][3]), totalAmount: Number(bData[i][7]) || 0,
        status: String(bData[i][8]), payDate: bData[i][17] ? Utilities.formatDate(new Date(bData[i][17]), "GMT+8", "yyyy-MM-dd") : '尚未繳款',
        roomId: roomId, tenant: String(bData[i][10]), endDate: String(bData[i][15]), lastReading: Number(bData[i][11]), 
        currentReading: Number(bData[i][12]), usage: Number(bData[i][13]), elecRate: Number(bData[i][14]), elecFee: Number(bData[i][5]), 
        waterFee: Number(bData[i][6]), rent: Number(bData[i][4]), prevBalance: Number(bData[i][18])||0, extraFee: Number(bData[i][19])||0, 
        extraNote: String(bData[i][20]||''), rentDeduct: Number(bData[i][21])||0
      });
    }
  } return history;
}
// 新增：處理台電退款結案，不再退回台電待結算清單
function confirmTpFinalRefundBackend(stId) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) throw new Error("系統忙碌");
  try {
    const stSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Settlements'); 
    const data = stSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === stId) { 
        stSheet.getRange(i + 1, 11).setValue('結案歸檔'); 
        return { success: true, message: '台電款項已確認退還，本單正式結案！' }; 
      }
    } 
    throw new Error('找不到該筆單據');
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}


function confirmDepositRefund(stId) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) throw new Error("系統忙碌");
  try {
    const stSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Settlements'); const data = stSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === stId) { stSheet.getRange(i + 1, 11).setValue('待台電結算'); return { success: true, message: '已確認退款！本單轉入台電待結算區。' }; }
    } throw new Error('找不到該筆單據');
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

// === 3. 台電登錄模組 ===
function saveTaipowerBill(data) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) throw new Error("系統忙碌中");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const tpSheet = ss.getSheetByName('TaipowerBills');
    const days = Math.ceil(Math.abs(new Date(data.endDate) - new Date(data.startDate)) / 86400000) + 1;
    let avgRate = (data.type === '私電') ? Number(data.totalAmount) / Number(data.totalKwh) : (Number(data.totalAmount) / 42) / days;
    tpSheet.appendRow(["TP" + new Date().getTime(), data.type, data.startDate, data.endDate, Number(data.totalAmount), Number(data.totalKwh) || 0, avgRate, new Date()]);
    return { success: true, message: `已儲存！單價: ${avgRate.toFixed(2)} 元` };
  } catch (err) { return { success: false, message: err.message }; } finally { lock.releaseLock(); }
}
// === 4. 退租模組 ===
function getRentedRooms() {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Rooms').getDataRange().getValues();
  let rooms = []; for (let i = 1; i < data.length; i++) { if (data[i][5] === '出租中') rooms.push({ roomId: String(data[i][0]), tenant: String(data[i][1]) }); } return rooms;
}

// ================= ★ 修改：按日攤提引擎 (修正天數缺漏判斷) ★ =================
function calculateProratedRefund(roomId, tenant, moveOutDateStr, finalUsageStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const billData = ss.getSheetByName('Billings').getDataRange().getValues();
  const tpData = ss.getSheetByName('TaipowerBills').getDataRange().getValues();

  let tpPrivate = [], tpPublic = [];
  for (let i = 1; i < tpData.length; i++) {
    let tpRec = { start: new Date(tpData[i][2]), end: new Date(tpData[i][3]), rate: Number(tpData[i][6]) };
    if (tpData[i][1] === '私電') tpPrivate.push(tpRec); else if (tpData[i][1] === '公電') tpPublic.push(tpRec);
  }

  let periods = [];
  let lastDate = null;
  // 1. 抓取歷史出帳紀錄
  for (let i = 1; i < billData.length; i++) {
    if (String(billData[i][2]) === roomId && String(billData[i][10]) === tenant) {
      let bDate = new Date(billData[i][1]);
      if (isNaN(bDate.getTime())) continue; 
      let sDate = lastDate ? new Date(lastDate.getTime() + 86400000) : new Date(bDate.getTime() - 29 * 86400000);
      periods.push({ start: sDate, end: bDate, kwh: Number(billData[i][13]) || 0, rate: Number(billData[i][14]) || 6.5 });
      lastDate = bDate;
    }
  }

  // 2. 加入退租當下畸零期 (修正：無歷史紀錄時自動往前推一個月)
  if (moveOutDateStr && finalUsageStr !== undefined) {
    let mDate = new Date(moveOutDateStr);
    if (!lastDate) {
      periods.push({ start: new Date(mDate.getTime() - 29 * 86400000), end: mDate, kwh: Number(finalUsageStr), rate: 6.5 });
    } else if (mDate > lastDate) {
      periods.push({ start: new Date(lastDate.getTime() + 86400000), end: mDate, kwh: Number(finalUsageStr), rate: 6.5 });
    }
  }

  let totalRefund = 0, missingDays = 0, calcDays = 0;
  // 3. 逐日攤提
  periods.forEach(p => {
    let sD = new Date(p.start.getFullYear(), p.start.getMonth(), p.start.getDate());
    let eD = new Date(p.end.getFullYear(), p.end.getMonth(), p.end.getDate());
    const daysInPeriod = Math.round((eD - sD) / 86400000) + 1;
    if(daysInPeriod <= 0) return;
    const dailyKwh = p.kwh / daysInPeriod;

    for (let d = new Date(sD); d <= eD; d.setDate(d.getDate() + 1)) {
      let cd = new Date(d); let privRate = null, pubRate = 0;
      for (let tp of tpPrivate) { 
        let tS = new Date(tp.start.getFullYear(), tp.start.getMonth(), tp.start.getDate()); let tE = new Date(tp.end.getFullYear(), tp.end.getMonth(), tp.end.getDate()); if (cd >= tS && cd <= tE) { privRate = tp.rate; break; } }
      for (let tp of tpPublic) { let tS = new Date(tp.start.getFullYear(), tp.start.getMonth(), tp.start.getDate()); let tE = new Date(tp.end.getFullYear(), tp.end.getMonth(), tp.end.getDate()); if (cd >= tS && cd <= tE) { pubRate = tp.rate; break; } }
      
      if (privRate === null) { missingDays++; } 
      else { totalRefund += (dailyKwh * (p.rate - privRate)) - pubRate; calcDays++; }
    }
  });
  return { refund: Math.round(totalRefund), missingDays: missingDays, calcDays: calcDays };
}

// 替換：退租資料讀取 (改叫引擎)
function getRoomCheckoutData(roomId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rData = ss.getSheetByName('Rooms').getDataRange().getValues();
    let roomInfo = {};
    for(let i=1; i<rData.length; i++) {
      if(String(rData[i][0]) === roomId) { roomInfo = { roomId: roomId, tenant: rData[i][1], elecRate: Number(rData[i][2])||6.5, lastReading: Number(rData[i][4])||0, rentAmount: Number(rData[i][6])||0, deposit: Number(rData[i][7])||0 }; break; }
    }

    // 將所有歷史帳單與台電費率「打包」傳給手機端，讓手機端能即時計算
    const bData = ss.getSheetByName('Billings').getDataRange().getValues();
    let history = []; let lastDate = null;
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][2]) === roomId && String(bData[i][10]) === roomInfo.tenant) {
        let bDate = new Date(bData[i][1]); if (isNaN(bDate.getTime())) continue;
        let sDate = lastDate ? new Date(lastDate.getTime() + 86400000) : new Date(bDate.getTime() - 29 * 86400000);
        history.push({ start: sDate.getTime(), end: bDate.getTime(), kwh: Number(bData[i][13])||0, rate: Number(bData[i][14])||6.5 });
        lastDate = bDate;
      }
    }

    const tpData = ss.getSheetByName('TaipowerBills').getDataRange().getValues();
    let tpPrivate = [], tpPublic = [];
    for (let i = 1; i < tpData.length; i++) {
      let tS = new Date(tpData[i][2]).getTime(), tE = new Date(tpData[i][3]).getTime();
      if (isNaN(tS) || isNaN(tE)) continue;
      let rec = { start: tS, end: tE, rate: Number(tpData[i][6]) };
      if (tpData[i][1] === '私電') tpPrivate.push(rec); else if (tpData[i][1] === '公電') tpPublic.push(rec);
    }

    return { room: roomInfo, history: history, tpPrivate: tpPrivate, tpPublic: tpPublic };
  } catch (e) { throw new Error(e.message); }
}

// 替換：退租寫入 (增加 L 欄 退租當期度數)
function processMoveOut(data) {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) throw new Error("系統忙碌中");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usage = Math.round((Number(data.finalReading) - Number(data.lastReading)) * 10) / 10;
    if (usage < 0) throw new Error("度數異常！");
    const finalElecFee = Math.round(usage * data.elecRate);
    const penaltyAmt = Number(data.penaltyAmt) || 0; const deductions = Number(data.deductions) || 0; const tpRefund = Number(data.tpRefund) || 0;
    const totalDeductions = penaltyAmt + deductions; const totalRefund = Number(data.deposit) - finalElecFee + tpRefund - totalDeductions;

    // 寫入 Settlements (L欄寫入使用度數)
    ss.getSheetByName('Settlements').appendRow([ "OUT" + new Date().getTime(), data.roomId, data.tenant, data.moveOutDate, data.finalReading, data.deposit, finalElecFee, tpRefund, totalDeductions, totalRefund, '待退押金', usage ]);
    
    const rSheet = ss.getSheetByName('Rooms'); const rData = rSheet.getDataRange().getValues();
    for (let i = 1; i < rData.length; i++) { if (String(rData[i][0]) === data.roomId) { rSheet.getRange(i+1, 2).clearContent(); rSheet.getRange(i+1, 5).setValue(data.finalReading); rSheet.getRange(i+1, 6).setValue('空房'); break; } }
    return { success: true, message: `結算單已產出！`, payload: { finalElecFee: finalElecFee, usage: usage, penaltyAmt: penaltyAmt, deductions: deductions, tpRefund: tpRefund, totalRefund: totalRefund } };
  } catch(e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}
// 替換：台電待結算清單 (全期計算差額)
function getPendingRefunds() {
  try {
    const stData = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Settlements').getDataRange().getValues();
    let pending = [];
    
    for (let i = 1; i < stData.length; i++) {
      // 必須是已經在對帳中心退完押金，轉入「待台電結算」的單據
      if (String(stData[i][10]) === '待台電結算') {
        let moveOutDate = stData[i][3];
        let moveOutDateStr = (moveOutDate instanceof Date) ? Utilities.formatDate(moveOutDate, "GMT+8", "yyyy-MM-dd") : String(moveOutDate);
        
        let stId = String(stData[i][0]);
        let roomId = String(stData[i][1]);
        let tenant = String(stData[i][2]);
        let finalUsage = String(stData[i][11] || 0); // L欄的最終使用度數
        let alreadyRefundedTp = Number(stData[i][7]) || 0; // 退租當下已經先退的台電金額 (H欄)

        // 呼叫後端的按日攤提引擎，重新以「最新」的台電帳單計算全期
        let calc = calculateProratedRefund(roomId, tenant, moveOutDateStr, finalUsage);
        
        // 剩餘應退差額 = 最新精算的總退還金額 - 退租當下已經先退給他的金額
        let remainingRefund = calc.refund - alreadyRefundedTp;

        pending.push({ 
          stId: stId, 
          roomId: roomId, 
          tenant: tenant, 
          moveOutDate: moveOutDateStr,
          missingDays: calc.missingDays, // 正確帶入缺漏天數
          remainingRefund: remainingRefund // 正確帶入應補退金額
        });
      }
    } 
    return pending;
  } catch (e) { 
    throw new Error(e.message);
  }
}
// === ★ 新增：產出最終單據並轉入對帳中心 ★ ===
function generateTaipowerFinalBill(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error("系統忙碌中");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const stSheet = ss.getSheetByName('Settlements');
    const bSheet = ss.getSheetByName('Billings');
    const data = stSheet.getDataRange().getValues();

    let found = false;
    // 1. 將舊的「待台電結算」標記為全數結清
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === payload.stId) {
        stSheet.getRange(i + 1, 11).setValue('全數結清');
        found = true;
        break;
      }
    }
    if (!found) throw new Error('找不到該筆退租單據');

    const amt = Number(payload.amt) || 0;
    const now = new Date();

    // 2. 判斷多退少補，產生對應的對帳紀錄
    if (amt > 0) {
      // 應退還給租客 -> 寫入 Settlements，狀態設為 '待退台電款' 讓對帳中心的綠色區塊抓取
      stSheet.appendRow(["TP_R" + now.getTime(), payload.roomId, payload.tenant, payload.moveOutDate, "", "", "", "", "", amt, '待退台電款', ""]);
    } else if (amt < 0) {
      // 應向租客補收 -> 寫入 Billings，狀態設為 '未收' 讓對帳中心的紅色區塊抓取
      const absAmt = Math.abs(amt);
      bSheet.appendRow(["TP_B" + now.getTime(), now, payload.roomId, "台電補收", 0, absAmt, 0, absAmt, '未收', '', payload.tenant, "", "", "", "", payload.moveOutDate, '', '', 0]);
    }

    return { success: true, message: '已轉入對帳中心！請至對帳中心完成結帳。' };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// ★ 總覽儀表板數據彙整
// ==========================================

function getDashboardSummary() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. 抓取房況與合約
    const rData = ss.getSheetByName('Rooms').getDataRange().getValues();
    let totalRooms = 0, rentedRooms = 0, expiringCount = 0, expiredCount = 0;
    const today = new Date();
    today.setHours(0,0,0,0);

    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][0]).trim() !== '') {
        totalRooms++; // 已經跳過標題列，所以這裡算出來的就是實際房數
        if (rData[i][5] === '出租中') {
          rentedRooms++;
          // 檢查合約是否到期
          if (rData[i][8] && rData[i][8] !== '未填寫') {
            const eDate = new Date(rData[i][8]);
            const diffDays = Math.ceil((eDate.getTime() - today.getTime()) / 86400000);
            if (diffDays < 0) expiredCount++;
            else if (diffDays <= 30) expiringCount++;
          }
        }
      }
    }
    
    // 2. 抓取應收帳款
    const bData = ss.getSheetByName('Billings').getDataRange().getValues();
    let totalUnpaid = 0, unpaidCount = 0;
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][0]).trim() !== '' && String(bData[i][8]) === '未收') {
        totalUnpaid += (Number(bData[i][7]) || 0);
        unpaidCount++;
      }
    }

    // 3. 抓取待退帳款 (押金/台電)
    const stData = ss.getSheetByName('Settlements').getDataRange().getValues();
    let totalRefunds = 0, refundCount = 0;
    for (let i = 1; i < stData.length; i++) {
      let status = String(stData[i][10]);
      if (status === '待退押金' || status === '待退台電款') {
        totalRefunds += (Number(stData[i][9]) || 0);
        refundCount++;
      }
    }

    return {
      success: true,
      data: {
        // ★ 修正：拿掉 -1，直接回傳 totalRooms 與正確的空房相減邏輯
        rooms: { 
          total: totalRooms, 
          rented: rentedRooms, 
          empty: totalRooms - rentedRooms, 
          expiring: expiringCount, 
          expired: expiredCount 
        },
        billing: { unpaidTotal: totalUnpaid, unpaidCount: unpaidCount },
        refunds: { refundTotal: totalRefunds, refundCount: refundCount }
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// getSettings — 讀取 Settings 工作表所有設定
// ══════════════════════════════════════════════════════════════
function getSettings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) { initSettingsSheet(); sheet = ss.getSheetByName('Settings'); }
    const data = sheet.getDataRange().getValues();
    const cfg = {};
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim();
      const val = String(data[i][1] || '').trim();
      if (key) cfg[key] = val;
    }
    return { success: true, data: cfg };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// saveSettings — 寫入 Settings 工作表
// ══════════════════════════════════════════════════════════════
function saveSettings(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) { initSettingsSheet(); sheet = ss.getSheetByName('Settings'); }
    const data = sheet.getDataRange().getValues();

    const updates = {
      bank_name:     payload.bank_name     || '',
      bank_account:  payload.bank_account  || '',
      bank_user:     payload.bank_user     || '',
      deadline_days: payload.deadline_days || '5',
      owner_name:    payload.owner_name    || '',
      system_note:   payload.system_note   || '',
    };

    // 更新已存在的 key
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim();
      if (updates.hasOwnProperty(key)) {
        sheet.getRange(i + 1, 2).setValue(updates[key]);
        delete updates[key];
      }
    }
    // 新增不存在的 key
    Object.entries(updates).forEach(([k, v]) => sheet.appendRow([k, v]));

    return { success: true, message: '設定已儲存' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// initSettingsSheet — 初始化 Settings 工作表預設值
// ══════════════════════════════════════════════════════════════
function initSettingsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.getRange(1, 1).setValue('key');
    sheet.getRange(1, 2).setValue('value');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 350);
  }
  const defaults = [
    ['bank_name',     '中國信託 (822) 鳳山分行'],
    ['bank_account',  '123-4567-89012'],
    ['bank_user',     '房東姓名或公司'],
    ['deadline_days', '5'],
    ['owner_name',    '房東'],
    ['system_note',   '如有疑問請聯繫房東'],
  ];
  const existing = sheet.getDataRange().getValues();
  const existingKeys = existing.slice(1).map(r => String(r[0]));
  defaults.forEach(([key, val]) => {
    if (!existingKeys.includes(key)) sheet.appendRow([key, val]);
  });
}

// ══════════════════════════════════════════════════════════════
// distributePublicElectricity — 公電分攤
// ══════════════════════════════════════════════════════════════
function distributePublicElectricity(payload) {
  // payload: { totalAmount, billingMonth }
  // 將公電費用按租戶數平均分攤到本月帳單
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const bSheet = ss.getSheetByName('Billings');
    const bData  = bSheet.getDataRange().getValues();
    const ym = payload.billingMonth; // yyyy-MM
    const amount = Number(payload.totalAmount) || 0;

    // 找本月帳單
    let rows = [];
    for (let i = 1; i < bData.length; i++) {
      const raw = bData[i][3];
      const rowYM = (raw instanceof Date)
        ? Utilities.formatDate(raw, 'GMT+8', 'yyyy-MM')
        : String(raw || '').slice(0, 7);
      if (rowYM === ym && String(bData[i][8]) !== '已收') rows.push(i + 1);
    }
    if (rows.length === 0) return { success: false, message: '找不到本月待收帳單' };

    const perRoom = Math.round(amount / rows.length);
    rows.forEach(rowIdx => {
      // T欄（第20欄）加扣款
      const cur = Number(bSheet.getRange(rowIdx, 20).getValue()) || 0;
      bSheet.getRange(rowIdx, 20).setValue(cur + perRoom);
      // U欄備註
      bSheet.getRange(rowIdx, 21).setValue('公電分攤');
      // 重算H欄（應收總計）= E+F+G+T - S（結餘）
      const rent  = Number(bSheet.getRange(rowIdx, 5).getValue())  || 0;
      const elec  = Number(bSheet.getRange(rowIdx, 6).getValue())  || 0;
      const water = Number(bSheet.getRange(rowIdx, 7).getValue())  || 0;
      const extra = cur + perRoom;
      const prev  = Number(bSheet.getRange(rowIdx, 19).getValue()) || 0;
      bSheet.getRange(rowIdx, 8).setValue(rent + elec + water + extra - prev);
    });

    return { success: true, message: `已將 $${amount} 公電費均攤至 ${rows.length} 戶（每戶 $${perRoom}）` };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
