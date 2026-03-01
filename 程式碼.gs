// ★★★ 重要：請將下方的 YOUR_SPREADSHEET_ID_HERE 換成你的試算表 ID ★★★
// 試算表 ID 在試算表網址中：https://docs.google.com/spreadsheets/d/【這裡就是ID】/edit
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
// ══════════════════════════════════════════════════════════════
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const body = JSON.parse(e.postData.contents);
    const fn   = String(body.fn   || '');
    const payload = body.payload;

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
      confirmTpFinalRefundBackend:  () => confirmTpFinalRefundBackend(payload),
      confirmDepositRefund:         () => confirmDepositRefund(payload),
      generateTaipowerFinalBill:    () => generateTaipowerFinalBill(payload),
      getRoomAnnualTpData:          () => getRoomAnnualTpData(payload),
      processAnnualTpReconciliation:() => processAnnualTpReconciliation(payload),
      getMonthlyReport:             () => getMonthlyReport(payload),
      exportMonthlyReportCSV:       () => exportMonthlyReportCSV(payload),
      getReportData:                () => getReportData(payload),
    };

    if (\!ALLOWED[fn]) {
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
// ROOMS 欄位 (新):
// A[0]房號 B[1]租客 C[2]狀態 D[3]合約起始日 E[4]合約到期日
// F[5]租金 G[6]押金 H[7]收款週期 I[8]電費單價 J[9]固定水費
// K[10]上期度數 L[11]預繳租金餘額 M[12]帳款結餘
// N[13]下次收租月 O[14]代收公電費 P[15]修繕備註
function getPendingRooms() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const roomData = ss.getSheetByName('Rooms').getDataRange().getValues();
  const meterData = ss.getSheetByName('MeterReadings').getDataRange().getValues();
  const now = new Date();
  const currentYM = Utilities.formatDate(now, "GMT+8", "yyyy-MM");

  // 建立本月已抄紀錄 Map：roomId → { lastReading, reading, usage, time, recordId }
  // MeterReadings 欄位: A[0]ID B[1]時間 C[2]房號 D[3]月份 E[4]上期度數 F[5]本期度數 G[6]用電度數 H[7]照片 I[8]操作者
  const doneMap = {};
  for (let i = 1; i < meterData.length; i++) {
    // ★ D欄若被 GAS 當成 Date 物件讀入，需格式化為 yyyy-MM
    const rawMYM = meterData[i][3];
    const mYM = (rawMYM instanceof Date)
      ? Utilities.formatDate(rawMYM, "GMT+8", "yyyy-MM")
      : String(rawMYM || '').trim();
    const mRoom = String(meterData[i][2] || '').trim(); // C: 房號
    if (mYM === currentYM && mRoom) {
      // 同房同月若多筆，取最新一筆
      doneMap[mRoom] = {
        recordId:    String(meterData[i][0]),
        time:        meterData[i][1] ? Utilities.formatDate(new Date(meterData[i][1]), "GMT+8", "yyyy-MM-dd HH:mm") : '',
        lastReading: Number(meterData[i][4]) || 0,  // E: 上期度數
        reading:     Number(meterData[i][5]) || 0,  // F: 本期度數
        usage:       Number(meterData[i][6]) || 0,  // G: 用電度數
        imageUrl:    String(meterData[i][7] || ''), // H: 照片
        yearMonth:   mYM
      };
    }
  }

  let pending = [];   // 未抄
  let completed = []; // 已抄

  for (let i = 1; i < roomData.length; i++) {
    if (roomData[i][2] !== '出租中') continue;
    const roomId = String(roomData[i][0]);
    const tenant = String(roomData[i][1]);
    const roomObj = {
      roomId:        roomId,
      tenant:        tenant,
      elecRate:      Number(roomData[i][8]) || 6.5,
      waterFee:      Number(roomData[i][9]) || 0,
      lastReading:   Number(roomData[i][10]) || 0,
      rentAmount:    Number(roomData[i][5]) || 0,
      deposit:       Number(roomData[i][6]) || 0,
      endDate:       roomData[i][4] ? Utilities.formatDate(new Date(roomData[i][4]), "GMT+8", "yyyy-MM-dd") : '未填寫',
      balance:       Number(roomData[i][11]) || 0,
      payCycle:      String(roomData[i][7] || '每月'),
      // N欄可能是 Date 物件，需格式化
      nextRentMonth: (roomData[i][13] instanceof Date)
        ? Utilities.formatDate(roomData[i][13], "GMT+8", "yyyy-MM")
        : String(roomData[i][13] || '')
    };

    if (doneMap[roomId]) {
      // 本月已抄：加入已完成清單
      completed.push(Object.assign({}, roomObj, {
        meterRecord: doneMap[roomId]
      }));
    } else {
      // 本月未抄：加入待抄清單
      pending.push(roomObj);
    }
  }

  return { pending: pending, completed: completed, yearMonth: currentYM };
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
        const oldEndDate = data[i][4] ? Utilities.formatDate(new Date(data[i][4]), "GMT+8", "yyyy-MM-dd") : '未填寫';
        const oldNotes = String(data[i][15] || '');
        const newNote = `[系統] ${Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")} 續約完成 (原到期日: ${oldEndDate})\n`;

        rSheet.getRange(i+1, 6).setValue(payload.newRent);    // F: 租金
        rSheet.getRange(i+1, 5).setValue(payload.newEndDate); // E: 合約到期日
        rSheet.getRange(i+1, 16).setValue(newNote + oldNotes); // P: 修繕備註
        // D欄: 續約起始日 = 原到期日的隔天
        const newStartDate = data[i][4]
          ? Utilities.formatDate(new Date(new Date(data[i][4]).getTime() + 86400000), "GMT+8", "yyyy-MM-dd")
          : Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
        rSheet.getRange(i+1, 4).setValue(newStartDate); // D: 合約起始日

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

    // ★ 防呆：檢查本月是否已抄過
    const meterSheet2 = ss.getSheetByName('MeterReadings');
    if (meterSheet2) {
      const mData = meterSheet2.getDataRange().getValues();
      for (let i = mData.length - 1; i >= 1; i--) {
        const rawMeterYM = mData[i][3];
        const meterYM = (rawMeterYM instanceof Date)
          ? Utilities.formatDate(rawMeterYM, "GMT+8", "yyyy-MM")
          : String(rawMeterYM || '').trim();
        if (String(mData[i][2]) === String(data.roomId) && meterYM === currentYM) {
          throw new Error(`${data.roomId} 本月 (${currentYM}) 已抄表完成，如需修正請至對帳中心修改帳單`);
        }
      }
    }
    
    const baseRent = Number(tRoom[5]) || 0;         // F: 租金
    const baseWater = Number(tRoom[9]) || 0;         // J: 固定水費
    const payCycle = String(tRoom[7] || '每月');     // H: 收款週期
    // ★ 修正：N欄若被 GAS 當成 Date 物件讀入，需格式化為 yyyy-MM
    let rawNextRentMonth = '';
    if (tRoom[13] instanceof Date) {
      rawNextRentMonth = Utilities.formatDate(tRoom[13], "GMT+8", "yyyy-MM");
    } else {
      rawNextRentMonth = String(tRoom[13] || '').trim();
    }
    // 格式驗證：必須符合 yyyy-MM，否則視為未初始化
    const nextRentMonth = /^\d{4}-\d{2}$/.test(rawNextRentMonth) ? rawNextRentMonth : '0000-00';
    const pendingPublicElec = Number(tRoom[14]) || 0; // O: 代收公電費
    const prepaidRentBal = Number(tRoom[11]) || 0;    // L: 預繳租金餘額
    const prevBillBal = Number(tRoom[12]) || 0;       // M: 帳款結餘

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
        rSheet.getRange(tIndex, 14).setNumberFormat("@").setValue(newNextRentMonth);
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
        // 帳款結餘為負數代表前期欠款，加入本期應收
        // deductBillAmt 保持 0（未抵用任何金額，而是補收欠款）
        finalTotal += Math.abs(newBillBal);
        deductBillAmt = 0;
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
      const newMeterRow = meterSheet.getLastRow() + 1;
      meterSheet.appendRow([
        "R" + now.getTime(),                                              // A: 紀錄ID
        Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"),       // B: 抄表時間
        data.roomId,                                                       // C: 房號
        periodStr,                                                         // D: 計費月份
        Number(data.lastReading),                                          // E: 上期度數
        Number(data.currentReading),                                       // F: 本期度數
        usage,                                                             // G: 用電度數
        imageUrl,                                                          // H: 電表照片
        Session.getActiveUser().getEmail()                                 // I: 操作者
      ]);
      // ★ 強制 D 欄為純文字格式，防止 GAS 將 yyyy-MM 自動轉成 Date 物件
      meterSheet.getRange(newMeterRow, 4).setNumberFormat("@");
    }

    const billSheet = ss.getSheetByName('Billings');
    if (billSheet) {
      const newBillRow = billSheet.getLastRow() + 1;
      billSheet.appendRow([
          "B" + now.getTime(), now, data.roomId, periodStr,
          actualRent, elecFee, actualWater, finalTotal,
          billStatus, payDate ? (payDate + " 00:00") : '', tRoom[1], data.lastReading, data.currentReading,
          usage, data.elecRate, tRoom[4], finalTotal === 0 ? subtotal : '', payDate, deductBillAmt,
          pendingPublicElec, pendingPublicElec > 0 ? "代收公電費" : "", rentDeduct
      ]);
      // ★ D欄強制純文字，防止 yyyy-MM 被解讀為 Date
      billSheet.getRange(newBillRow, 4).setNumberFormat("@");
    }

    rSheet.getRange(tIndex, 11).setValue(data.currentReading);  // K: 上期度數
    rSheet.getRange(tIndex, 12).setValue(newPrepaidRentBal);     // L: 預繳租金餘額
    if (rSheet.getMaxColumns() < 16) rSheet.insertColumnsAfter(rSheet.getMaxColumns(), 16 - rSheet.getMaxColumns());
    rSheet.getRange(tIndex, 13).setValue(newBillBal);            // M: 帳款結餘
    rSheet.getRange(tIndex, 15).clearContent();                  // O: 代收公電費清除

    // 帶回完整帳單物件，供前端立刻產繳費單（不依賴對帳中心快取）
    const billId = "B" + now.getTime();
    const endDateStr = tRoom[4] ? Utilities.formatDate(new Date(tRoom[4]), "GMT+8", "yyyy-MM-dd") : '未填寫';
    return {
      success: true,
      message: `${data.roomId} 抄表完成！`,
      bill: {
        billId:          billId,
        roomId:          data.roomId,
        tenant:          String(tRoom[1]),
        yearMonth:       periodStr,
        rent:            actualRent,
        elecFee:         elecFee,
        waterFee:        actualWater,
        totalAmount:     finalTotal,
        lastReading:     Number(data.lastReading),
        currentReading:  Number(data.currentReading),
        usage:           usage,
        elecRate:        Number(data.elecRate),
        endDate:         endDateStr,
        extraFee:        pendingPublicElec,
        extraNote:       pendingPublicElec > 0 ? '代收公電費' : '',
        rentDeduct:      rentDeduct,
        prevBalance:     deductBillAmt
      }
    };
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

  // 帳單月份以合約起始日為準，若無則用今天
  const startBase = p.startDate ? new Date(p.startDate) : now;
  const periodStr = Utilities.formatDate(startBase, "GMT+8", "yyyy-MM");

  const initialRent    = (Number(p.rent)     || 0) * multiplier;
  const initialWater   = (Number(p.waterFee) || 0) * multiplier;
  const deposit        = Number(p.deposit)   || 0;
  const prepaidRentBal = Number(p.balance)   || 0;

  // 預繳租金扣抵
  let rentDeduct = 0;
  let newPrepaidRentBal = prepaidRentBal;
  if (initialRent > 0 && prepaidRentBal > 0) {
    rentDeduct = Math.min(initialRent, prepaidRentBal);
    newPrepaidRentBal -= rentDeduct;
  }

  // 押金列入第一期帳單（T欄 extraFee），備註說明
  const finalTotal = (initialRent - rentDeduct) + initialWater + deposit;
  const cycleLabel = multiplier > 1 ? p.payCycle + '租金 $' + initialRent : '租金 $' + initialRent;
  const noteStr    = deposit > 0
    ? '入住押金 $' + deposit + '（' + cycleLabel + '，水費 $' + initialWater + '）'
    : cycleLabel + '，水費 $' + initialWater;

  const billId = "B" + now.getTime();
  const initBillRow = bSheet.getLastRow() + 1;
  bSheet.appendRow([
    billId,                                                                     // A: 帳單ID
    now,                                                                        // B: 建立時間
    p.roomId,                                                                   // C: 房號
    periodStr,                                                                  // D: 計費年月
    initialRent,                                                                // E: 租金
    0,                                                                          // F: 電費
    initialWater,                                                               // G: 水費
    finalTotal,                                                                 // H: 應收總額
    finalTotal === 0 ? '已收' : '未收',                                         // I: 狀態
    finalTotal === 0 ? Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm") : '', // J: 收款時間
    p.tenant,                                                                   // K: 租客
    Number(p.lastReading) || 0,                                                 // L: 上期度數
    Number(p.lastReading) || 0,                                                 // M: 本期度數
    0,                                                                          // N: 用電度數
    Number(p.elecRate) || 0,                                                    // O: 電費單價
    p.endDate || '',                                                            // P: 合約到期日
    finalTotal === 0 ? finalTotal : '',                                         // Q: 實收
    finalTotal === 0 ? Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd") : '',  // R: 收款日
    0,                                                                          // S: 帳款抵用
    deposit,                                                                    // T: 加扣款（押金）
    noteStr,                                                                    // U: 備註
    rentDeduct                                                                  // V: 預繳扣抵
  ]);
  bSheet.getRange(initBillRow, 4).setNumberFormat("@");

  // 更新 Rooms 預繳餘額
  const rData = rSheet.getDataRange().getValues();
  for (let i = 1; i < rData.length; i++) {
    if (String(rData[i][0]) === String(p.roomId)) {
      rSheet.getRange(i + 1, 12).setValue(newPrepaidRentBal);
      break;
    }
  }

  // ★ 回傳帳單物件，讓 saveTenantInfo 轉交前端立即產出繳費單
  return {
    billId:         billId,
    roomId:         p.roomId,
    tenant:         p.tenant,
    yearMonth:      periodStr,
    rent:           initialRent,
    elecFee:        0,
    waterFee:       initialWater,
    totalAmount:    finalTotal,
    lastReading:    Number(p.lastReading) || 0,
    currentReading: Number(p.lastReading) || 0,
    usage:          0,
    elecRate:       Number(p.elecRate) || 0,
    endDate:        p.endDate || '',
    extraFee:       deposit,
    extraNote:      deposit > 0 ? '入住押金' : '',
    rentDeduct:     rentDeduct
  };
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
              // 溢繳或少繳金額，寫入 M 欄 (帳款結餘, index12)
              rSheet.getRange(r+1, 13).setValue((Number(rData[r][12]) || 0) + diff);
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
        let lastReading = Number(bData[i][11]), elecRate = Number(bData[i][14]);
        // ★ 修正：deductBillAmt(S欄/index18)是「原已抵用額」，修改帳單時保持不變
        // prevBal 應從 Billings 本身的 deductBillAmt 讀出，用於還原應收小計
        const deductBillAmt = Number(bData[i][18]) || 0;
        let newCurRead = Number(payload.currentReading);
        let newUsage = Math.round((newCurRead - lastReading) * 10) / 10;
        if (newUsage < 0) throw new Error("度數不可小於上期！");
        let newElecFee = Math.round(newUsage * elecRate);
        
        let rent = Number(payload.rent) || 0, water = Number(payload.waterFee) || 0, extra = Number(payload.extraFee) || 0;
        let rentDeduct = Number(bData[i][21]) || 0;
        if (rentDeduct > rent) rentDeduct = rent; // 防呆：扣抵額不能大於修改後的租金
        
        // 總應收 = 重算小計 - 原帳款抵用額（保持不變）
        let newSubtotal = (rent - rentDeduct) + newElecFee + water + extra;
        let newTotal = Math.max(0, newSubtotal - deductBillAmt);

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
      roomId:      String(data[i][0]),
      tenant:      String(data[i][1]),
      status:      String(data[i][2]),           // C
      startDate:   data[i][3]  ? Utilities.formatDate(new Date(data[i][3]),  "GMT+8", "yyyy-MM-dd") : '未填寫', // D
      endDate:     data[i][4]  ? Utilities.formatDate(new Date(data[i][4]),  "GMT+8", "yyyy-MM-dd") : '未填寫', // E
      rent:        Number(data[i][5]) || 0,      // F
      deposit:     Number(data[i][6]) || 0,      // G
      payCycle:    String(data[i][7] || '每月'), // H
      elecRate:    Number(data[i][8]) || 0,      // I
      waterFee:    Number(data[i][9]) || 0,      // J
      lastReading: Number(data[i][10]) || 0,     // K
      balance:     Number(data[i][11]) || 0,     // L: 預繳租金餘額
      billBalance: Number(data[i][12]) || 0,     // M: 帳款結餘
      notes:       String(data[i][15] || ''),    // P: 修繕備註
      remark:      String(data[i][16] || '')     // Q: 租客備註（顯示於卡片）
      remark:      String(data[i][16] || '')     // Q: 租客備註
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
        const oldStatus = String(data[i][2]); // C: 狀態
        let rowValues = data[i];
        // 新欄位順序: A[0]房號 B[1]租客 C[2]狀態 D[3]起始日 E[4]到期日
        // F[5]租金 G[6]押金 H[7]收款週期 I[8]電費單價 J[9]固定水費
        // K[10]上期度數 L[11]預繳餘額 M[12]帳款結餘 N[13]下次收租月 O[14]代收公電 P[15]備註
        rowValues[1]  = payload.tenant;
        rowValues[2]  = payload.status;
        rowValues[3]  = payload.startDate || rowValues[3] || ''; // D: 合約起始日
        rowValues[4]  = payload.endDate;                         // E: 合約到期日
        rowValues[5]  = Number(payload.rent) || 0;               // F: 租金
        rowValues[6]  = Number(payload.deposit) || 0;            // G: 押金
        rowValues[7]  = payload.payCycle || '每月';              // H: 收款週期
        rowValues[8]  = Number(payload.elecRate) || 0;           // I: 電費單價
        rowValues[9]  = Number(payload.waterFee) || 0;           // J: 固定水費
        rowValues[11] = Number(payload.balance) || 0;            // L: 預繳租金餘額
        rowValues[15] = payload.notes  || '';  // P: 修繕備註
        while(rowValues.length < 17) rowValues.push("");
        rowValues[16] = payload.remark || '';  // Q: 租客備註（顯示於卡片）

        while(rowValues.length < 16) rowValues.push("");
        rowValues[12] = Number(payload.billBalance) || 0;        // M: 帳款結餘

        // ★ 空房判斷：空字串或 '空房' 都算空房
        const wasEmpty = (oldStatus === '空房' || oldStatus.trim() === '');
        let isNewTenant = (wasEmpty && payload.status === '出租中');
        if (isNewTenant) {
          // 下次收租月以「合約起始日」為基準，若無則用今天
          const startBase = payload.startDate ? new Date(payload.startDate) : new Date();
          let multiplier = 1;
          if (payload.payCycle === '半年') multiplier = 6;
          if (payload.payCycle === '年收') multiplier = 12;
          // 下次收租月 = 起始月 + 週期（第一期已產出，下次從起始月 + multiplier 起算）
          const nextDate = new Date(startBase.getFullYear(), startBase.getMonth() + multiplier, 1);
          rowValues[13] = Utilities.formatDate(nextDate, "GMT+8", "yyyy-MM"); // N: 下次收租月
          if (!rowValues[3]) rowValues[3] = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd"); // D: 起始日自動填今日
        } else if (!rowValues[13]) {
          rowValues[13] = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM");
        }

        rSheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
        // N欄(第14欄)強制文字格式，防止 yyyy-MM 被自動轉成 Date
        rSheet.getRange(i + 1, 14).setNumberFormat("@");

        // ★ 新租客：產出第一期帳單，並回傳帳單物件讓前端立即顯示繳費單
        let newBill = null;
        if (isNewTenant) newBill = generateInitialBill(ss, payload);

        SpreadsheetApp.flush();
        return {
          success: true,
          message: isNewTenant ? '租客設定完成，第一期帳單已產出！' : '設定與修繕紀錄已儲存！',
          isNewTenant: isNewTenant,
          bill: newBill   // 前端用來立即產繳費單
        };
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
        billId: String(bData[i][0]),
        yearMonth: (bData[i][3] instanceof Date)
          ? Utilities.formatDate(bData[i][3], "GMT+8", "yyyy-MM")
          : String(bData[i][3] || ''),
        totalAmount: Number(bData[i][7]) || 0,
        status: String(bData[i][8]), payDate: bData[i][17] ? Utilities.formatDate(new Date(bData[i][17]), "GMT+8", "yyyy-MM-dd") : '尚未繳款',
        roomId: roomId, tenant: String(bData[i][10]),
        endDate: bData[i][15] ? (bData[i][15] instanceof Date ? Utilities.formatDate(bData[i][15], "GMT+8", "yyyy-MM-dd") : String(bData[i][15]).slice(0,10)) : '未填寫',
        lastReading: Number(bData[i][11]), 
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
    let avgRate = (data.type === '私電') ? Number(data.totalAmount) / Number(data.totalKwh) : Number(data.totalAmount) / days;
    const totalDays = days; // 計費天數
    tpSheet.appendRow([
      "TP" + new Date().getTime(),  // A: 台電ID
      new Date(),                   // B: 登錄時間
      data.type,                    // C: 類型(私電/公電)
      data.startDate,               // D: 起始日
      data.endDate,                 // E: 結束日
      totalDays,                    // F: 計費天數
      Number(data.totalAmount),     // G: 總金額
      Number(data.totalKwh) || 0,  // H: 總度數(私電用)
      avgRate                       // I: 平均單價
    ]);
    return { success: true, message: `已儲存！單價: ${avgRate.toFixed(2)} 元` };
  } catch (err) { return { success: false, message: err.message }; } finally { lock.releaseLock(); }
}
// === 4. 退租模組 ===
function getRentedRooms() {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Rooms').getDataRange().getValues();
  let rooms = []; for (let i = 1; i < data.length; i++) { if (data[i][2] === '出租中') rooms.push({ roomId: String(data[i][0]), tenant: String(data[i][1]) }); } return rooms;
}

// ================= ★ 修改：按日攤提引擎 (修正天數缺漏判斷) ★ =================
function calculateProratedRefund(roomId, tenant, moveOutDateStr, finalUsageStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const billData = ss.getSheetByName('Billings').getDataRange().getValues();
  const tpData = ss.getSheetByName('TaipowerBills').getDataRange().getValues();

  let tpPrivate = [], tpPublic = [];
  for (let i = 1; i < tpData.length; i++) {
    // TaipowerBills 欄位: A[0]ID B[1]登錄時間 C[2]類型 D[3]起始日 E[4]結束日 F[5]天數 G[6]金額 H[7]度數 I[8]單價
    let tpRec = { start: new Date(tpData[i][3]), end: new Date(tpData[i][4]), rate: Number(tpData[i][8]) };
    if (tpData[i][2] === '私電') tpPrivate.push(tpRec); else if (tpData[i][2] === '公電') tpPublic.push(tpRec);
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
      if(String(rData[i][0]) === roomId) { roomInfo = { roomId: roomId, tenant: rData[i][1], elecRate: Number(rData[i][8])||6.5, lastReading: Number(rData[i][10])||0, rentAmount: Number(rData[i][5])||0, deposit: Number(rData[i][6])||0 }; break; }
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
      // TaipowerBills: C[2]類型 D[3]起始日 E[4]結束日 I[8]單價
      let tS = new Date(tpData[i][3]).getTime(), tE = new Date(tpData[i][4]).getTime();
      if (isNaN(tS) || isNaN(tE)) continue;
      let rec = { start: tS, end: tE, rate: Number(tpData[i][8]) };
      if (tpData[i][2] === '私電') tpPrivate.push(rec); else if (tpData[i][2] === '公電') tpPublic.push(rec);
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
    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][0]) === data.roomId) {
        rSheet.getRange(i+1, 2).clearContent();          // B: 租客姓名
        rSheet.getRange(i+1, 3).setValue('空房');          // C: 狀態
        rSheet.getRange(i+1, 4).clearContent();            // D: 合約起始日
        rSheet.getRange(i+1, 5).clearContent();            // E: 合約到期日
        rSheet.getRange(i+1, 8).clearContent();            // H: 收款週期
        rSheet.getRange(i+1, 11).setValue(data.finalReading); // K: 更新最終度數
        rSheet.getRange(i+1, 12).setValue(0);              // L: 預繳租金餘額歸零
        rSheet.getRange(i+1, 13).setValue(0);              // M: 帳款結餘歸零
        rSheet.getRange(i+1, 14).clearContent();           // N: 下次收租月
        break;
      }
    }
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
      bSheet.appendRow(["TP_B" + now.getTime(), now, payload.roomId, "台電補收",
        0, absAmt, 0, absAmt,      // E租金=0, F電費=補收額, G水費=0, H應收
        '未收', '',                 // I狀態, J收款時間
        payload.tenant,             // K租客
        "", "", "", "",             // L上期度, M本期度, N用電度, O電費單價
        payload.moveOutDate,        // P合約到期日(借用顯示退租日)
        '', '', 0,                  // Q實收, R收款日, S帳款抵用
        0, "台電補收", 0            // T代收公電, U備註, V預繳扣抵
      ]);
    }

    return { success: true, message: '已轉入對帳中心！請至對帳中心完成結帳。' };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


// ==========================================
// ★ 年度台電差額結算（續約用）
// ==========================================

// 1. 讀取該房間的帳單期間 + 台電費率供前端計算
function getRoomAnnualTpData(roomId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rData = ss.getSheetByName('Rooms').getDataRange().getValues();
    let roomInfo = null;
    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][0]) === roomId) {
        roomInfo = {
          roomId:    roomId,
          tenant:    String(rData[i][1]),
          startDate: rData[i][3] ? Utilities.formatDate(new Date(rData[i][3]), "GMT+8", "yyyy-MM-dd") : '',
          endDate:   rData[i][4] ? Utilities.formatDate(new Date(rData[i][4]), "GMT+8", "yyyy-MM-dd") : '',
          elecRate:  Number(rData[i][8]) || 6.5
        };
        break;
      }
    }
    if (!roomInfo) throw new Error('找不到房號');

    // 抓取本租約期間的所有帳單（以合約起始日篩選）
    const bData = ss.getSheetByName('Billings').getDataRange().getValues();
    const contractStart = roomInfo.startDate ? new Date(roomInfo.startDate).getTime() : 0;
    let history = [];
    let lastDate = null;
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][2]) !== roomId) continue;
      if (String(bData[i][10]) !== roomInfo.tenant) continue;
      let bDate = new Date(bData[i][1]);
      if (isNaN(bDate.getTime())) continue;
      // 只取本合約起始日之後的帳單
      if (bDate.getTime() < contractStart) continue;
      let sDate = lastDate
        ? new Date(lastDate.getTime() + 86400000)
        : new Date(bDate.getTime() - 29 * 86400000);
      history.push({
        start: sDate.getTime(),
        end:   bDate.getTime(),
        kwh:   Number(bData[i][13]) || 0,
        rate:  Number(bData[i][14]) || 6.5
      });
      lastDate = bDate;
    }

    // 台電費率
    const tpData = ss.getSheetByName('TaipowerBills').getDataRange().getValues();
    let tpPrivate = [], tpPublic = [];
    for (let i = 1; i < tpData.length; i++) {
      let tS = new Date(tpData[i][3]).getTime(), tE = new Date(tpData[i][4]).getTime();
      if (isNaN(tS) || isNaN(tE)) continue;
      let rec = { start: tS, end: tE, rate: Number(tpData[i][8]) };
      if (tpData[i][2] === '私電') tpPrivate.push(rec);
      else if (tpData[i][2] === '公電') tpPublic.push(rec);
    }

    return {
      room: roomInfo,
      history: history,
      lastBillDate: lastDate ? Utilities.formatDate(lastDate, "GMT+8", "yyyy-MM-dd") : '',
      tpPrivate: tpPrivate,
      tpPublic: tpPublic
    };
  } catch (e) { throw new Error(e.message); }
}

// 2. 寫入結算結果 + 同步執行續約
function processAnnualTpReconciliation(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) throw new Error("系統忙碌中");
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const now = new Date();
    const amt = Number(payload.tpDiff) || 0; // 正=退客戶, 負=補收
    const periodStr = (payload.startDate || '') + ' ~ ' + (payload.endDate || '');

    // A. 台電差額寫入
    if (amt > 0) {
      // 應退租客：寫 Settlements 狀態 '待退台電款'
      ss.getSheetByName('Settlements').appendRow([
        "AR_" + now.getTime(),   // A: ID (Annual Reconciliation)
        payload.roomId,          // B: 房號
        payload.tenant,          // C: 租客
        Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd"), // D: 結算日期
        '', payload.prevDeposit || '', '', '', '', // E-I
        amt,                     // J: 退還金額
        '待退台電款',             // K: 狀態
        '年度台電差額_' + periodStr // L: 備註
      ]);
    } else if (amt < 0) {
      // 應補收：寫 Billings 狀態 '未收'
      const absAmt = Math.abs(amt);
      const bSheet = ss.getSheetByName('Billings');
      const newRow = bSheet.getLastRow() + 1;
      bSheet.appendRow([
        "AR_B" + now.getTime(), now, payload.roomId, "年度台電補收",
        0, absAmt, 0, absAmt,
        '未收', '',
        payload.tenant,
        "", "", "", "",
        payload.endDate,
        '', '', 0,
        0, "年度台電差額補收 " + periodStr, 0
      ]);
      bSheet.getRange(newRow, 4).setNumberFormat("@");
    }
    // amt === 0 時不產生任何帳款，直接續約

    // B. 執行續約（同 renewContract 邏輯）
    const rSheet = ss.getSheetByName('Rooms');
    const rData = rSheet.getDataRange().getValues();
    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][0]) === payload.roomId) {
        const oldEndDate = rData[i][4]
          ? Utilities.formatDate(new Date(rData[i][4]), "GMT+8", "yyyy-MM-dd")
          : '未填寫';
        const oldNotes = String(rData[i][15] || '');
        const tpNote = amt !== 0
          ? `台電差額: ${amt > 0 ? '+' : ''}$${amt} | `
          : '';
        const newNote = `[系統] ${Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd")} 續約完成 (原到期日: ${oldEndDate} | ${tpNote}結算期: ${periodStr})
`;

        const newStartDate = rData[i][4]
          ? Utilities.formatDate(new Date(new Date(rData[i][4]).getTime() + 86400000), "GMT+8", "yyyy-MM-dd")
          : Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

        rSheet.getRange(i+1, 5).setValue(payload.newEndDate);         // E: 新到期日
        rSheet.getRange(i+1, 6).setValue(Number(payload.newRent)||0); // F: 新租金
        rSheet.getRange(i+1, 4).setValue(newStartDate);               // D: 起始日
        rSheet.getRange(i+1, 16).setValue(newNote + oldNotes);        // P: 備註
        break;
      }
    }

    let msg = '續約完成！';
    if (amt > 0) msg += ` 台電退還 $${amt} 已轉入對帳中心。`;
    else if (amt < 0) msg += ` 台電補收 $${Math.abs(amt)} 已轉入對帳中心。`;
    else msg += ' 台電差額為 $0，無需另行處理。';

    return { success: true, message: msg, tpAmt: amt };
  } catch (e) {
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
        if (rData[i][2] === '出租中') {
          rentedRooms++;
          // 檢查合約是否到期
          if (rData[i][4] && rData[i][4] !== '未填寫') {
            const eDate = new Date(rData[i][4]);
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


// ==========================================
// ★ 月份收款明細報表
// ==========================================

/**
 * 取得指定年份的月份收款明細
 * 回傳：每月 × 每房 的租金/電費/其他/押金/未收 明細
 */
function getMonthlyReport(year) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const bData = ss.getSheetByName('Billings').getDataRange().getValues();
    const rData = ss.getSheetByName('Rooms').getDataRange().getValues();
    const stData= ss.getSheetByName('Settlements').getDataRange().getValues();

    const targetYear = String(year || new Date().getFullYear());

    // ── 1. 建立房間清單（排序用） ──────────────────────────────────────
    const roomOrder = [];
    const roomMap   = {}; // roomId -> { tenant, deposit(來自 Rooms 表) }
    for (let i = 1; i < rData.length; i++) {
      const rid = String(rData[i][0] || '').trim();
      if (!rid) continue;
      roomOrder.push(rid);
      roomMap[rid] = { tenant: String(rData[i][1] || ''), rent: Number(rData[i][5]) || 0 };
    }

    // ── 2. 彙整所有帳單（按「房號___租客」分組，支援同房號不同租客） ──
    // monthData[ym][groupKey] = { rid, tenant, rent, elec, water, extra, extraNote, totalAmt, paid, unpaid }
    // groupKey = rid + '___' + tenant
    const monthData = {};
    for (let i = 1; i < bData.length; i++) {
      const rid    = String(bData[i][2]  || '').trim();
      const tenant = String(bData[i][10] || '').trim(); // K: 租客姓名
      if (!rid) continue;
      const rawYM = bData[i][3];
      const ym = (rawYM instanceof Date)
        ? Utilities.formatDate(rawYM, 'GMT+8', 'yyyy-MM')
        : String(rawYM || '');
      if (!ym.startsWith(targetYear)) continue;

      const groupKey = rid + '___' + tenant;
      if (!monthData[ym]) monthData[ym] = {};
      if (!monthData[ym][groupKey]) {
        monthData[ym][groupKey] = { rid, tenant, rent:0, elec:0, water:0, extra:0, extraNote:'', totalAmt:0, paid:0, unpaid:0 };
      }

      const row    = monthData[ym][groupKey];
      const status = String(bData[i][8]);
      const total  = Number(bData[i][7]) || 0;
      const actual = Number(bData[i][16]) || total; // Q欄：實收

      row.rent  += Number(bData[i][4])  || 0;  // E: 租金
      row.elec  += Number(bData[i][5])  || 0;  // F: 電費
      row.water += Number(bData[i][6])  || 0;  // G: 水費
      row.extra += Number(bData[i][19]) || 0;  // T: 加扣款
      if (bData[i][20]) row.extraNote = String(bData[i][20]);
      row.totalAmt += total;
      if (status === '已收') { row.paid += actual; }
      else { row.unpaid += total; }
    }

    // ── 3. 已收押金（Settlements 表，本年度已確認的押金） ──────────────
    // Settlements: A[0]ID B[1]房號 C[2]租客 D[3]日期 ... K[10]狀態
    const depositPaid = {}; // roomId -> 已收押金金額（累計）
    for (let i = 1; i < stData.length; i++) {
      const rid    = String(stData[i][1] || '').trim();
      const status = String(stData[i][10] || '');
      // 押金入帳（非「待退押金」狀態的 Settlements 才算已收）
      // 改從 Billings 的 extraFee 欄判斷較準確，這裡只記錄退租時的已退押金
      // 實際已收押金已包含在帳單的 extra 欄位（入住時的第一期帳單）
    }

    // ── 4. 整理月份清單（含有資料的月份，排序） ───────────────────────
    const months = Object.keys(monthData).sort();

    // ── 5. 計算每月合計（按 groupKey 彙整） ──────────────────────────
    const monthSummary = {};
    months.forEach(ym => {
      let sumRent=0, sumElec=0, sumWater=0, sumExtra=0, sumPaid=0, sumUnpaid=0;
      Object.values(monthData[ym]).forEach(r => {
        sumRent  += r.rent;
        sumElec  += r.elec;
        sumWater += r.water;
        sumExtra += r.extra;
        sumPaid  += r.paid;
        sumUnpaid+= r.unpaid;
      });
      monthSummary[ym] = { sumRent, sumElec, sumWater, sumExtra, sumPaid, sumUnpaid };
    });

    return {
      success: true,
      data: {
        year: targetYear,
        roomOrder,
        roomMap,
        months,
        monthData,
        monthSummary
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 匯出指定年份月份報表為 CSV（UTF-8 BOM，Excel 可直接開啟）
 */
function exportMonthlyReportCSV(year) {
  try {
    const res = getMonthlyReport(year);
    if (!res.success) return { success: false, message: res.message };
    const d = res.data;

    const rows = [];
    // 標題行
    rows.push(['月份', '房號', '租客', '租金', '電費', '水費', '其他收款', '其他說明', '應收合計', '已收金額', '未收金額', '收款狀態']);

    d.months.forEach(ym => {
      const roomData = d.monthData[ym];
      // 依 rid 排序（同房號不同租客各一行）
      const groupKeys = Object.keys(roomData).sort((a, b) => {
        const ia = d.roomOrder.indexOf(a.split('___')[0]);
        const ib = d.roomOrder.indexOf(b.split('___')[0]);
        return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
      });
      groupKeys.forEach(gk => {
        const r = roomData[gk];
        rows.push([
          ym,
          r.rid,
          r.tenant,
          r.rent,
          r.elec,
          r.water,
          r.extra,
          r.extraNote,
          r.totalAmt,
          r.paid,
          r.unpaid,
          r.unpaid > 0 ? '未收' : '已收'
        ]);
      });
      // 月份小計行
      const s = d.monthSummary[ym];
      rows.push([ym + ' 小計', '', '', s.sumRent, s.sumElec, s.sumWater, s.sumExtra, '', s.sumPaid + s.sumUnpaid, s.sumPaid, s.sumUnpaid, '']);
      rows.push([]); // 空行分隔
    });

    // 年度合計
    let totRent=0, totElec=0, totWater=0, totExtra=0, totPaid=0, totUnpaid=0;
    d.months.forEach(ym => {
      const s = d.monthSummary[ym];
      totRent+=s.sumRent; totElec+=s.sumElec; totWater+=s.sumWater;
      totExtra+=s.sumExtra; totPaid+=s.sumPaid; totUnpaid+=s.sumUnpaid;
    });
    rows.push([year + ' 年度合計', '', '', totRent, totElec, totWater, totExtra, '', totPaid+totUnpaid, totPaid, totUnpaid, '']);

    // 轉 CSV
    const csv = '\uFEFF' + rows.map(r => r.map(v => '"' + String(v || '').replace(/"/g, '""') + '"').join(',')).join('\n');
    return { success: true, csv: csv, filename: year + '_月份收款明細.csv' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// ★ Settings 工作表管理
// ==========================================

/**
 * 初始化 Settings 工作表（若不存在則建立，若存在則補齊缺漏的 key）
 * 手動執行一次，或在 doGet 時自動呼叫
 */
function initSettingsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.getRange('A1:B1').setValues([['Key', 'Value']]);
    sheet.getRange('A1:B1').setFontWeight('bold');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 350);
  }

  const defaults = [
    ['bank_name',     '中國信託 (822) 鳳山分行'],
    ['bank_account',  '123-4567-89012'],
    ['bank_user',     '房東姓名或公司'],
    ['deadline_days', '5'],
    ['owner_name',    '房東'],
    ['system_note',   '如有疑問請聯繫房東']
  ];

  const existing = sheet.getDataRange().getValues();
  const existingKeys = existing.slice(1).map(r => String(r[0]));

  defaults.forEach(([key, val]) => {
    if (!existingKeys.includes(key)) {
      sheet.appendRow([key, val]);
    }
  });

  // 格式化：A 欄淡灰底
  const lastRow = sheet.getLastRow();
  sheet.getRange(2, 1, lastRow - 1, 1).setBackground('#f3f4f6');
  SpreadsheetApp.flush();
  return { success: true, message: 'Settings 工作表初始化完成！' };
}

/**
 * 讀取所有設定，回傳 { bank_name, bank_account, ... } 物件
 */
function getSettings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      // 第一次使用時自動初始化
      initSettingsSheet();
      return getSettings();
    }
    const rows = sheet.getDataRange().getValues().slice(1); // 跳過標題
    const cfg = {};
    rows.forEach(r => {
      if (String(r[0]).trim()) cfg[String(r[0]).trim()] = String(r[1] || '');
    });
    return { success: true, data: cfg };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 儲存設定（前端傳來 key-value 物件）
 */
function saveSettings(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, message: '系統忙碌' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) { initSettingsSheet(); sheet = ss.getSheetByName('Settings'); }

    const rows = sheet.getDataRange().getValues();
    Object.entries(payload).forEach(([key, val]) => {
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === key) {
          sheet.getRange(i + 1, 2).setValue(val);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([key, val]);
    });
    SpreadsheetApp.flush();
    return { success: true, message: '設定已儲存！' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally { lock.releaseLock(); }
}

// ==========================================
// ★ 報表與數據分析
// ==========================================

function getReportData() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const bData = ss.getSheetByName('Billings').getDataRange().getValues();
    const rData = ss.getSheetByName('Rooms').getDataRange().getValues();
    const tpData= ss.getSheetByName('TaipowerBills').getDataRange().getValues();

    const now       = new Date();
    const thisYear  = now.getFullYear();
    const lastYear  = thisYear - 1;

    // ── 1. 月收入趨勢（近 12 個月，已收帳單） ──────────────────────────
    const monthlyIncome = {}; // key: yyyy-MM, value: 已收金額
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][8]) !== '已收') continue;
      // 收款日期用 R 欄(17)，若無則用建立時間 B 欄(1)
      const rawDate = bData[i][17] || bData[i][1];
      if (!rawDate) continue;
      const d  = new Date(rawDate);
      const ym = Utilities.formatDate(d, 'GMT+8', 'yyyy-MM');
      monthlyIncome[ym] = (monthlyIncome[ym] || 0) + (Number(bData[i][16]) || Number(bData[i][7]) || 0);
    }

    // 產生近 12 個月 key 清單（含 0）
    const last12 = [];
    for (let m = 11; m >= 0; m--) {
      const d  = new Date(now.getFullYear(), now.getMonth() - m, 1);
      const ym = Utilities.formatDate(d, 'GMT+8', 'yyyy-MM');
      last12.push({ ym, income: Math.round(monthlyIncome[ym] || 0) });
    }

    // ── 2. 每房獲利（本年度已收租金 - 台電補貼差額，按房號） ─────────────
    // 本年度已收帳單按房號彙總
    const roomIncome   = {}; // roomId -> 已收金額
    const roomRent     = {}; // roomId -> 租金部分
    const roomElec     = {}; // roomId -> 電費補貼（台電差額結算）
    for (let i = 1; i < bData.length; i++) {
      const rid = String(bData[i][2] || '').trim();
      if (!rid) continue;
      // 計費年月欄（D欄=index3）
      const rawYM = bData[i][3];
      const ym = (rawYM instanceof Date)
        ? Utilities.formatDate(rawYM, 'GMT+8', 'yyyy-MM')
        : String(rawYM || '');
      if (!ym.startsWith(String(thisYear))) continue; // 只算本年度
      if (String(bData[i][8]) !== '已收') continue;
      const actual = Number(bData[i][16]) || Number(bData[i][7]) || 0;
      roomIncome[rid] = (roomIncome[rid] || 0) + actual;
      roomRent[rid]   = (roomRent[rid]   || 0) + (Number(bData[i][4]) || 0);
      roomElec[rid]   = (roomElec[rid]   || 0) + (Number(bData[i][5]) || 0);
    }

    // 取 Rooms 表房號清單（含空房，收入為 0）
    const roomList = [];
    for (let i = 1; i < rData.length; i++) {
      const rid    = String(rData[i][0] || '').trim();
      if (!rid) continue;
      const status = String(rData[i][2] || '');
      roomList.push({
        roomId:  rid,
        tenant:  String(rData[i][1] || '-'),
        status:  status,
        rent:    Number(rData[i][5]) || 0,
        income:  Math.round(roomIncome[rid] || 0),
        rentSum: Math.round(roomRent[rid]   || 0),
        elecSum: Math.round(roomElec[rid]   || 0)
      });
    }
    roomList.sort((a, b) => b.income - a.income);

    // ── 3. 台電費用趨勢（近 12 個月私電費率，按帳單區間） ────────────────
    // TaipowerBills: A[0]ID B[1]時間 C[2]類型 D[3]起日 E[4]迄日 F[5]天數 G[6]金額 H[7]度數 I[8]費率 J[9]每日每戶公電
    const tpTrend = [];
    for (let i = 1; i < tpData.length; i++) {
      if (String(tpData[i][2]) !== '私電') continue;
      const rawEnd = tpData[i][4];
      if (!rawEnd) continue;
      const endDate = (rawEnd instanceof Date) ? rawEnd : new Date(rawEnd);
      const ym      = Utilities.formatDate(endDate, 'GMT+8', 'yyyy-MM');
      const rate    = Number(tpData[i][8]) || 0; // I: 費率
      const amount  = Number(tpData[i][6]) || 0; // G: 金額
      const kwh     = Number(tpData[i][7]) || 0; // H: 度數
      tpTrend.push({ ym, rate: Math.round(rate * 100) / 100, amount, kwh });
    }
    // 只取近 12 個月，按時間排序
    tpTrend.sort((a, b) => a.ym.localeCompare(b.ym));
    const tpTrend12 = tpTrend.filter(t => t.ym >= last12[0].ym);

    // ── 4. 年度摘要 ──────────────────────────────────────────────────────
    let yearTotal = 0, yearCount = 0;
    for (let i = 1; i < bData.length; i++) {
      const rawYM = bData[i][3];
      const ym = (rawYM instanceof Date)
        ? Utilities.formatDate(rawYM, 'GMT+8', 'yyyy-MM')
        : String(rawYM || '');
      if (!ym.startsWith(String(thisYear))) continue;
      if (String(bData[i][8]) !== '已收') continue;
      yearTotal += Number(bData[i][16]) || Number(bData[i][7]) || 0;
      yearCount++;
    }

    let lastYearTotal = 0;
    for (let i = 1; i < bData.length; i++) {
      const rawYM = bData[i][3];
      const ym = (rawYM instanceof Date)
        ? Utilities.formatDate(rawYM, 'GMT+8', 'yyyy-MM')
        : String(rawYM || '');
      if (!ym.startsWith(String(lastYear))) continue;
      if (String(bData[i][8]) !== '已收') continue;
      lastYearTotal += Number(bData[i][16]) || Number(bData[i][7]) || 0;
    }

    return {
      success: true,
      data: {
        thisYear,
        lastYear,
        yearTotal:     Math.round(yearTotal),
        lastYearTotal: Math.round(lastYearTotal),
        yearCount,
        last12,
        roomList,
        tpTrend12
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}


// Line Notify 功能已移除（2024年 Line 已停止 Notify 服務）


