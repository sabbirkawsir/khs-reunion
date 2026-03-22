// ============================================================
// কুলানন্দপুর উচ্চ বিদ্যালয় ঈদ পুনর্মিলনী ২০২৬
// Google Apps Script — সম্পূর্ণ ব্যাকেন্ড
// ============================================================
// ব্যবহারের নিয়ম:
// ১. এই কোডটি Google Apps Script এ পেস্ট করুন (script.google.com)
// ২. SHEET_ID তে আপনার Google Sheet এর ID বসান
// ৩. Deploy > New Deployment > Web App হিসেবে পাবলিশ করুন
// ৪. "Execute as: Me" এবং "Who has access: Anyone" সিলেক্ট করুন
// ৫. পাওয়া URL টি সব HTML ফাইলের SCRIPT_URL এ বসান
// ============================================================

const SHEET_ID = '1eloL9zLLugXRSeZqD2JxvYKBwX_Urj_dLhkUDGcoiF4'; // আপনার Sheet ID
const REG_SHEET = 'নিবন্ধন';       // Registration data sheet name
const EXP_SHEET = 'ব্যয়';          // Expense data sheet name

// ==================== GET Handler ====================
function doGet(e) {
  const action = e.parameter.action || 'list';
  let result;
  
  try {
    if (action === 'list') {
      result = getRegistrations();
    } else if (action === 'find') {
      result = findRegistration(e.parameter.q || '');
    } else if (action === 'summary') {
      result = getSummary();
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== POST Handler ====================
function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (data.action === 'expense') {
      result = addExpense(data);
    } else {
      result = saveRegistration(data);
    }
  } catch (err) {
    result = { error: err.message };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== Registration Functions ====================
function saveRegistration(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(REG_SHEET);
  
  // শীট না থাকলে তৈরি করুন
  if (!sheet) {
    sheet = ss.insertSheet(REG_SHEET);
    const headers = [
      'নিবন্ধন আইডি', 'নাম', 'মোবাইল', 'ইমেইল', 'ক্যাটাগরি',
      'পেশা/শ্রেণী', 'ব্যাচ', 'ঠিকানা', 'প্রাপ্তবয়স্ক সদস্য',
      'শিশু সদস্য', 'পেমেন্ট পদ্ধতি', 'ট্রানজেকশন আইডি',
      'মোট ফী', 'নোট', 'নিবন্ধনের সময়'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold')
      .setBackground('#0d4023').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  
  const row = [
    data.regId || '',
    data.name || '',
    data.phone || '',
    data.email || '',
    data.category || '',
    data.profession || '',
    data.batch || '',
    data.address || '',
    data.adults || 0,
    data.children || 0,
    data.paymentMethod || '',
    data.txnId || '',
    data.totalFee || 0,
    data.note || '',
    data.timestamp || new Date().toLocaleString('bn-BD')
  ];
  
  sheet.appendRow(row);
  
  // Auto-resize columns
  try { sheet.autoResizeColumns(1, row.length); } catch(e) {}
  
  return { success: true, regId: data.regId };
}

function getRegistrations() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(REG_SHEET);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { data: [] };
  }
  
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();
  
  const data = rows.map(r => ({
    regId: r[0],
    name: r[1],
    phone: r[2],
    email: r[3],
    category: r[4],
    profession: r[5],
    batch: r[6],
    address: r[7],
    adults: r[8],
    children: r[9],
    paymentMethod: r[10],
    txnId: r[11],
    totalFee: r[12],
    note: r[13],
    timestamp: r[14]
  })).filter(r => r.name); // Empty rows skip
  
  return { data };
}

function findRegistration(q) {
  const all = getRegistrations();
  const found = all.data.find(d => 
    d.regId === q || d.phone === q
  );
  return found ? { data: found } : { data: null };
}

// ==================== Expense Functions ====================
function addExpense(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(EXP_SHEET);
  
  if (!sheet) {
    sheet = ss.insertSheet(EXP_SHEET);
    const headers = ['ব্যয়ের খাত', 'পরিমাণ (টাকা)', 'নোট', 'তারিখ'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold')
      .setBackground('#c0392b').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  
  sheet.appendRow([data.name, data.amt, data.note || '', data.date || new Date().toLocaleDateString('bn-BD')]);
  return { success: true };
}

function getExpenses() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(EXP_SHEET);
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(r => r[0])
    .map(r => ({ name: r[0], amt: r[1], note: r[2], date: r[3] }));
}

// ==================== Summary Function ====================
function getSummary() {
  const regs = getRegistrations().data;
  const exps = getExpenses();
  
  const jobRegs = regs.filter(r => r.category === 'চাকুরীজীবী');
  const stuRegs = regs.filter(r => r.category !== 'চাকুরীজীবী');
  const totalAdults = regs.reduce((s, r) => s + (parseInt(r.adults) || 0), 0);
  
  const jobAmt = jobRegs.length * 1000;
  const stuAmt = stuRegs.length * 500;
  const adultsAmt = totalAdults * 1000;
  
  return {
    income: {
      job: jobRegs.length,
      stu: stuRegs.length,
      jobAmt,
      stuAmt,
      adultsCnt: totalAdults,
      adultsAmt,
      total: jobAmt + stuAmt + adultsAmt
    },
    expenses: exps,
    totalRegistrations: regs.length
  };
}

// ==================== Setup Function ====================
// প্রথমবার রান করুন — শীট সেটআপ করতে
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Summary sheet
  let summary = ss.getSheetByName('সারসংক্ষেপ');
  if (!summary) summary = ss.insertSheet('সারসংক্ষেপ');
  
  summary.clearContents();
  summary.getRange('A1').setValue('কুলানন্দপুর উচ্চ বিদ্যালয় — ঈদ পুনর্মিলনী ২০২৬');
  summary.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#0d4023');
  
  Logger.log('Setup সম্পন্ন হয়েছে! এখন Deploy করুন।');
  SpreadsheetApp.getUi().alert('✅ Setup সফল হয়েছে! এখন Deploy > New Deployment করুন।');
}
