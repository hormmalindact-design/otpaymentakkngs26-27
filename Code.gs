
/********************************************************
 * NETLIFY API GATEWAY
 * Netlify Function -> Apps Script doPost -> Google Sheets
 * Set Script Property API_GATEWAY_SECRET to the same
 * secret configured as Netlify environment variable.
 ********************************************************/
function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var req = JSON.parse(raw);
    var props = PropertiesService.getScriptProperties();
    var expectedSecret = String(props.getProperty('API_GATEWAY_SECRET') || '');
    var suppliedSecret = String(req.gatewaySecret || '');

    if (!expectedSecret || !suppliedSecret || suppliedSecret !== expectedSecret) {
      return jsonApiResponse_({ success: false, error: 'មិនមានសិទ្ធិចូលប្រើ API ទេ។' });
    }

    var action = String(req.action || '').trim();
    var payload = req.payload || {};
    var result;

    switch (action) {
      case 'getTeachersList':
        result = getTeachersList();
        break;
      case 'getPaymentData':
        result = getPaymentData();
        break;
      case 'chatWithAI':
        result = chatWithAI(payload.messages, payload.selectedModel);
        break;
      default:
        return jsonApiResponse_({ success: false, error: 'មុខងារ API មិនត្រឹមត្រូវ។' });
    }

    return jsonApiResponse_(result);
  } catch (err) {
    console.error('doPost API error: ' + (err && err.stack ? err.stack : err));
    return jsonApiResponse_({
      success: false,
      error: err && err.message ? err.message : 'មានបញ្ហាក្នុងការដំណើរការ API។'
    });
  }
}

function jsonApiResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// កែសម្រួលមុខងារ doGet ដើម្បីផ្តាច់ Route រវាង Admin និង Mobile View របស់គ្រូ
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  let tpl;
  
  if (page === 'mobileview') {
    // បើ URL មាន ?page=mobileview ឱ្យវាបើកឯកសារ MobileView.html ដោយមិនចាំបាច់ Login
    tpl = HtmlService.createTemplateFromFile('MobileView');
    return tpl
      .evaluate()
      .setTitle('📱 របាយការណ៍គ្រូបង្រៀន')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0');
  }
  
  // បើគ្មាន ?page ទេ ឱ្យវាបើកឯកសារ WebDashboard.html (មានផ្ទាំង Login សម្រាប់ Admin)
  tpl = HtmlService.createTemplateFromFile('WebDashboard');
  return tpl
    .evaluate()
    .setTitle('ប្រព័ន្ធគ្រប់គ្រងការបង់ប្រាក់សាលារៀន - Admin')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0');
}

/** បង្កើតសម័យ Admin សម្រាប់ AI endpoint ប៉ុណ្ណោះ */
function createAdminAiSession_(username) {
  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('AI_ADMIN_' + token, String(username), 21600);
  return token;
}
function requireAdminAiSession_(token) {
  var t = String(token || '').trim();
  if (!t || t.length < 30) throw new Error('សូមចូល Admin ម្តងទៀត ដើម្បីផ្ទៀងផ្ទាត់សិទ្ធិ។');
  var username = CacheService.getScriptCache().get('AI_ADMIN_' + t);
  if (!username) throw new Error('សម័យ Admin ផុតកំណត់។ សូម Logout ហើយ Login ម្តងទៀត។');
  return username;
}
function logoutAdminAiSession(token) {
  var t = String(token || '').trim();
  if (t) CacheService.getScriptCache().remove('AI_ADMIN_' + t);
  return true;
}

// មុខងារផ្ទៀងផ្ទាត់ការចូលប្រើប្រាស់របស់ Admin
function verifyAdminLogin(username, password) {
  var u = String(username || '').trim();
  var p = String(password || '').trim();
  if (!u || !p) {
    return { success: false, message: 'សូមបញ្ចូលឈ្មោះអ្នកប្រើប្រាស់ និងលេខសម្ងាត់!' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var adminSheet = ss.getSheetByName('Admin');

  // ប្រសិនបើមាន Sheet ឈ្មោះ 'Admin' (ជួរឈរ A: Username, ជួរឈរ B: Password, ជួរឈរ C: Display Name)
  if (adminSheet && adminSheet.getLastRow() >= 2) {
    var data = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, 3).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      var sheetUser = String(data[i][0] || '').trim();
      var sheetPass = String(data[i][1] || '').trim();
      var displayName = String(data[i][2] || sheetUser).trim();
      if (sheetUser !== '' && sheetUser === u && sheetPass === p) {
        return { success: true, adminName: displayName, aiToken: createAdminAiSession_(u) };
      }
    }
    return { success: false, message: 'ឈ្មោះអ្នកប្រើប្រាស់ ឬលេខសម្ងាត់មិនត្រឹមត្រូវ!' };
  }

  // គណនី Admin លំនាំដើម (Default) ប្រសិនបើមិនទាន់បង្កើត Sheet 'Admin'
  // អ្នកអាចប្តូរ Username និង Password នៅទីនេះបាន
  if (u === 'admin' && p === '123456') {
    return { success: true, adminName: 'អ្នកគ្រប់គ្រងប្រព័ន្ធ (Admin)', aiToken: createAdminAiSession_(u) };
  }

  return { success: false, message: 'ឈ្មោះអ្នកប្រើប្រាស់ ឬលេខសម្ងាត់មិនត្រឹមត្រូវ!' };
}

// មុខងារជំនួយសម្រាប់បម្លែងភេទឱ្យទៅជាស្តង់ដារ (Male / Female)
function normalizeGender_(g) {
  var s = String(g || '').trim().toLowerCase();
  if (['female', 'f', 'ស្រី', 'ស្ត្រី', 'ភេទស្រី'].indexOf(s) !== -1) return 'Female';
  if (['male', 'm', 'ប្រុស', 'បុរស', 'ភេទប្រុស'].indexOf(s) !== -1) return 'Male';
  return String(g || '').trim();
}

function getOrCreatePaymentSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Payment');
  if (!sheet) {
    sheet = ss.insertSheet('Payment');
    sheet.getRange(1, 1, 1, 15).setValues([[
      'ID','payment_id','student_id','student_name','gender','class','teacher_name',
      'amount','teacher80','school20','daily_amount','paid_date','report_date','days','Payment Method'
    ]]);
    sheet.getRange(1, 1, 1, 15).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function dateToYmd_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var d = new Date(s);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getNextStudentId() {
  var sheet = getOrCreatePaymentSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '001';
  var ids = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
  var maxId = 0;
  ids.forEach(function(r) { var n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (!isNaN(n) && n > maxId) maxId = n; });
  return String(maxId + 1).padStart(3, '0');
}

function checkStudentIdExists(studentId, currentRowIndex) {
  var sheet = getOrCreatePaymentSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var ids = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    var rowNumber = i + 2;
    if (currentRowIndex && rowNumber === Number(currentRowIndex)) continue;
    if (String(ids[i][0]).trim() === String(studentId).trim()) return true;
  }
  return false;
}

function getTeachersList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var teachers = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    var status = String(data[i][3] || '').trim();
    if (name && (status === 'Active' || !status)) {
      teachers.push({
        name: name,
        gender: normalizeGender_(data[i][2]),
        telegram: data[i][4] || ''
      });
    }
  }
  return teachers;
}

function getPaymentData() {
  var sheet = getOrCreatePaymentSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var teacherGenderMap = {};
  var teachers = getTeachersList();
  teachers.forEach(function(t) {
    teacherGenderMap[t.name] = t.gender;
  });

  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var result = [];
  data.forEach(function(r, index) {
    if (!r[3]) return;
    var tName = String(r[6] || '').trim();
    result.push({
      rowIndex: index + 2,
      payment_id: r[1] || '',
      student_id: String(r[2] || '').trim(),
      student_name: String(r[3] || '').trim(),
      gender: normalizeGender_(r[4]),
      class: String(r[5] || '').trim(),
      teacher_name: tName,
      teacher_gender: teacherGenderMap[tName] || '',
      amount: Number(r[7]) || 0,
      teacher80: Number(r[8]) || 0,
      school20: Number(r[9]) || 0,
      paid_date: dateToYmd_(r[11]),
      payment_method: r[14] || 'Cash'
    });
  });
  return result;
}

function savePaymentRecord(form) {
  if (!form || !String(form.student_id || '').trim() || !String(form.student_name || '').trim()) {
    return { status: 'error', message: 'សូមបំពេញលេខកូដ និងឈ្មោះសិស្ស។' };
  }
  var sheet = getOrCreatePaymentSheet();
  var row = form.rowIndex ? Number(form.rowIndex) : 0;
  if (row && (row < 2 || row > sheet.getLastRow())) return { status: 'error', message: 'មិនអាចរកឃើញជួរទិន្នន័យដែលត្រូវកែប្រែទេ។' };
  if (checkStudentIdExists(form.student_id, row)) return { status: 'error', message: 'លេខកូដសិស្ស ' + form.student_id + ' មានរួចហើយក្នុងប្រព័ន្ធ!' };
  var amount = Number(form.amount) || 0;
  if (amount < 0) return { status: 'error', message: 'ចំនួនប្រាក់មិនអាចតិចជាងសូន្យបានទេ។' };
  var paidDate = dateToYmd_(form.paid_date);
  if (!paidDate) return { status: 'error', message: 'សូមជ្រើសរើសកាលបរិច្ឆេទបង់ប្រាក់។' };
  
  var cleanGender = normalizeGender_(form.gender);
  var teacher80 = amount * 0.8, school20 = amount * 0.2, days = 30;
  
  if (row) {
    sheet.getRange(row, 3, 1, 10).setValues([[
      String(form.student_id).trim(), String(form.student_name).trim(), cleanGender, String(form.class || '').trim(),
      String(form.teacher_name || '').trim(), amount, teacher80, school20, amount / days, paidDate
    ]]);
    sheet.getRange(row, 15).setValue(form.payment_method || 'Cash');
    sheet.getRange(row, 12).setNumberFormat('@');
    sheet.getRange(row, 12).setValue(paidDate);
    return { status: 'success', message: 'កែប្រែទិន្នន័យបានជោគជ័យ!' };
  }
  var paymentId = 'PAY-' + new Date().getTime();
  sheet.appendRow([
    '', paymentId, String(form.student_id).trim(), String(form.student_name).trim(), cleanGender,
    String(form.class || '').trim(), String(form.teacher_name || '').trim(), amount, teacher80,
    school20, amount / days, paidDate, paidDate, days, form.payment_method || 'Cash'
  ]);
  return { status: 'success', message: 'រក្សាទុកទិន្នន័យបានជោគជ័យ!' };
}

function deletePaymentRecord(rowIndex) {
  var sheet = getOrCreatePaymentSheet();
  var row = Number(rowIndex);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('ជួរទិន្នន័យមិនត្រឹមត្រូវ។');
  sheet.deleteRow(row);
  return 'លុបទិន្នន័យបានជោគជ័យ!';
}

function backupMonthlyData(monthStr) {
  if (!/^\d{4}-\d{2}$/.test(String(monthStr))) return 'ទម្រង់ខែមិនត្រឹមត្រូវ។';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var paymentSheet = getOrCreatePaymentSheet();
  var archiveName = 'Archive_' + monthStr;
  var archiveSheet = ss.getSheetByName(archiveName);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveName);
    var headers = paymentSheet.getRange(1, 1, 1, 15).getValues();
    archiveSheet.getRange(1, 1, 1, 15).setValues(headers);
    archiveSheet.getRange(1, 1, 1, 15).setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
    archiveSheet.setFrozenRows(1);
  }
  var lastRow = paymentSheet.getLastRow();
  if (lastRow < 2) return 'គ្មានទិន្នន័យសម្រាប់ខែ ' + monthStr + ' ទេ';
  var data = paymentSheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var rows = data.filter(function(r) { return dateToYmd_(r[11]).slice(0, 7) === monthStr; });
  if (!rows.length) return 'គ្មានទិន្នន័យសម្រាប់ខែ ' + monthStr + ' ទេ';
  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rows.length, 15).setValues(rows);
  return 'ការ Backup ទិន្នន័យខែ ' + monthStr + ' បានជោគជ័យ! (' + rows.length + ' ជួរ)';
}


/********************************************************
 * KHMER AI CHATBOT - MWAPI INTEGRATION
 * បន្ថែមមុខងារថ្មី ដោយមិនប៉ះពាល់ Payment System
 ********************************************************/

/**
 * ទាញយកការកំណត់ API ពី Script Properties
 */
function getMWAPIConfig_() {
  var props = PropertiesService.getScriptProperties();

  var apiKey = props.getProperty('MWAPI_KEY');
  var baseUrl = props.getProperty('MWAPI_BASE_URL')
    || 'https://api.mwapi.dev';
  var defaultModel = props.getProperty('MWAPI_MODEL')
    || 'claude-sonnet-4-6';

  if (!apiKey) {
    throw new Error(
      'រកមិនឃើញ MWAPI_KEY ទេ។ សូមកំណត់ API Key ក្នុង Script Properties។'
    );
  }

  return {
    apiKey: apiKey.trim(),
    baseUrl: baseUrl.replace(/\/+$/, ''),
    model: defaultModel.trim()
  };
}


/**
 * មុខងារសន្ទនាជាមួយ AI
 * ហៅពី Index.html តាម google.script.run.chatWithAI()
 */
function chatWithAI(messages, selectedModel) {
  try {

    var config = getMWAPIConfig_();

    // ពិនិត្យទិន្នន័យសន្ទនា
    if (!Array.isArray(messages) || messages.length === 0) {
      return {
        success: false,
        error: 'សូមបញ្ចូលសំណួររបស់អ្នកជាមុនសិន។'
      };
    }

    // កំណត់ Model
    var model = String(selectedModel || config.model).trim();

    // ការពារកុំឱ្យបញ្ចូល Model ទទេ ឬអក្សរខុសទម្រង់
    if (!/^[a-zA-Z0-9._-]+$/.test(model)) {
      return {
        success: false,
        error: 'ឈ្មោះ AI Model មិនត្រឹមត្រូវទេ។'
      };
    }

    // រៀបចំប្រវត្តិសន្ទនា
    var cleanMessages = messages.slice(-20).map(function(item) {

      var role = String(item.role || '');
      var content = String(item.content || '');

      if (role !== 'user' && role !== 'assistant') {
        throw new Error('តួនាទីសន្ទនាមិនត្រឹមត្រូវ។');
      }

      if (!content.trim()) {
        throw new Error('សារសន្ទនាមិនអាចទទេបានទេ។');
      }

      return {
        role: role,
        content: content.substring(0, 12000)
      };
    });

    // URL របស់ MWAPI
    var apiUrl = config.baseUrl + '/v1/messages';

    // រៀបចំ Request
    var payload = {
      model: model,
      max_tokens: 1400,
      temperature: 0.3,

      system:
        'អ្នកគឺជា Khmer AI Assistant ដែលមានជំនាញខ្ពស់។ ' +
        'សូមឆ្លើយតបជាភាសាខ្មែរដែលច្បាស់លាស់ សុភាពរាបសា និងងាយយល់។ ' +
        'ប្រសិនបើអ្នកប្រើសួរជាភាសាផ្សេង អ្នកអាចឆ្លើយតាមភាសាដែលគេសួរ។ ' +
        'សូមផ្តល់ចម្លើយមានរចនាសម្ព័ន្ធ និងឧទាហរណ៍នៅពេលចាំបាច់។',

      messages: cleanMessages
    };

    var options = {
      method: 'post',

      contentType: 'application/json',

      headers: {
        'x-api-key': config.apiKey,
        'anthropic-version': '2023-06-01'
      },

      payload: JSON.stringify(payload),

      muteHttpExceptions: true
    };

    // ហៅទៅ MWAPI
    var response = UrlFetchApp.fetch(apiUrl, options);

    var statusCode = response.getResponseCode();
    var responseText = response.getContentText();

    var result;

    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      throw new Error(
        'MWAPI បានបញ្ជូនទិន្នន័យមិនមែន JSON។ HTTP ' +
        statusCode
      );
    }

    // ពិនិត្យ HTTP Error
    if (statusCode < 200 || statusCode >= 300) {

      var errorMessage =
        result.error && result.error.message
          ? result.error.message
          : 'HTTP ' + statusCode;

      console.error('MWAPI Error: ' + responseText);

      return {
        success: false,
        error: 'MWAPI Error: ' + errorMessage
      };
    }

    // ដកចម្លើយពី Content
    var answer = '';

    if (result.content && Array.isArray(result.content)) {
      answer = result.content
        .filter(function(block) {
          return block.type === 'text';
        })
        .map(function(block) {
          return block.text || '';
        })
        .join('\n');
    }

    if (!answer) {
      return {
        success: false,
        error: 'AI មិនបានបញ្ជូនចម្លើយមកវិញទេ។'
      };
    }

    return {
      success: true,
      answer: answer,
      model: result.model || model
    };

  } catch (error) {

    console.error('Chatbot Error: ' + error.message);

    return {
      success: false,
      error: error.message || 'មានបញ្ហាក្នុងការភ្ជាប់ AI។'
    };
  }
}



/********************************************************
 * SCHOOL AI ASSISTANT - PHASE 1 (READ-ONLY + ADMIN AUTH)
 ********************************************************/
function aiGetSchoolSummary(adminToken, monthStr) {
  requireAdminAiSession_(adminToken);
  var month = String(monthStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'));
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('ទម្រង់ខែមិនត្រឹមត្រូវ។');
  var payments = getPaymentData(), teachers = getTeachersList();
  var filtered = payments.filter(function(r) { return String(r.paid_date || '').slice(0, 7) === month; });
  var total=0, teacher80=0, school20=0, cash=0, qr=0, studentSet={}, teacherMap={};
  filtered.forEach(function(r) {
    var amount=Number(r.amount)||0; total+=amount; teacher80+=Number(r.teacher80)||0; school20+=Number(r.school20)||0;
    if (String(r.payment_method||'').toUpperCase()==='QR') qr+=amount; else cash+=amount;
    studentSet[String(r.student_id||'')]=true;
    var name=String(r.teacher_name||'មិនទាន់កំណត់');
    if(!teacherMap[name]) teacherMap[name]={teacher_name:name,students:{},total:0,teacher80:0,school20:0};
    teacherMap[name].students[String(r.student_id||'')]=true; teacherMap[name].total+=amount;
    teacherMap[name].teacher80+=Number(r.teacher80)||0; teacherMap[name].school20+=Number(r.school20)||0;
  });
  var byTeacher=Object.keys(teacherMap).map(function(name){var x=teacherMap[name];return {teacher_name:x.teacher_name,student_count:Object.keys(x.students).length,total:x.total,teacher80:x.teacher80,school20:x.school20};}).sort(function(a,b){return b.total-a.total;});
  return {month:month,payment_records:filtered.length,unique_students:Object.keys(studentSet).length,active_teachers:teachers.length,total_amount:total,teacher80:teacher80,school20:school20,cash_amount:cash,qr_amount:qr,by_teacher:byTeacher};
}

function aiAskSchoolData(adminToken, question, monthStr, selectedModel) {
  requireAdminAiSession_(adminToken);
  var q=String(question||'').trim();
  if(!q) return {success:false,error:'សូមបញ្ចូលសំណួររបស់អ្នក។'};
  if(q.length>8000) return {success:false,error:'សំណួរវែងពេក (អតិបរមា 8,000 តួអក្សរ)។'};
  var r=aiGetSchoolSummary(adminToken,monthStr);
  // គណនាភាគរយនៅ Server ដើម្បីកុំឱ្យ AI គណនាដោយស្មាន
  var verifiedTeacher80 = Math.round((Number(r.total_amount) || 0) * 0.80);
  var verifiedSchool20 = (Number(r.total_amount) || 0) - verifiedTeacher80;
  // កំណត់ចំនួនគ្រូក្នុងបរិបទ ដើម្បីកុំឱ្យ request ធំពេក និងយឺត
  var teacherRows = (r.by_teacher || []).slice(0, 30);
  var context=[
    'ទិន្នន័យសាលាពី Google Sheets សម្រាប់ខែ '+r.month,
    'ចំនួនកំណត់ត្រាបង់ប្រាក់: '+r.payment_records,
    'ចំនួនសិស្សមិនស្ទួន: '+r.unique_students,
    'ចំនួនគ្រូសកម្ម: '+r.active_teachers,
    'ប្រាក់សរុប 100%: '+r.total_amount+' រៀល',
    'ចំណែកគ្រូ 80% (គណនាផ្ទៀងផ្ទាត់ពី Server): '+verifiedTeacher80+' រៀល',
    'ចំណែកសាលា 20% (គណនាផ្ទៀងផ្ទាត់ពី Server): '+verifiedSchool20+' រៀល',
    'Cash: '+r.cash_amount+' រៀល',
    'QR: '+r.qr_amount+' រៀល',
    'ទិន្នន័យតាមគ្រូ (បង្ហាញអតិបរមា 30 គ្រូ): '+JSON.stringify(teacherRows)
  ].join('\n');
  console.log('aiAskSchoolData: month='+r.month+', questionLength='+q.length+', teacherRows='+teacherRows.length);
  return chatWithAI([{role:'user',content:
    'សំណួរ៖ '+q+'\\n\\n'+context+
    '\\n\\nសេចក្តីណែនាំ៖ ប្រើតែទិន្នន័យ Google Sheets ខាងលើ។ តម្លៃចំណែកគ្រូ និងសាលាដែលបានផ្តល់ គឺជាលទ្ធផលគណនាផ្ទៀងផ្ទាត់ពី Server រួចហើយ។ ហាមគណនាឡើងវិញ ឬប្តូរតម្លៃទាំងនេះ។ បង្ហាញចំនួនជាលេខមានសញ្ញាក្បៀស និងឯកតា រៀល។ កុំប្រើ Markdown table ឬសញ្ញា | លុះត្រាតែអ្នកប្រើស្នើ។ បើទិន្នន័យមិនគ្រប់ សូមប្រាប់ថាមិនមានទិន្នន័យ។ ឆ្លើយជាភាសាខ្មែរ។'
  }],selectedModel);
}

function testAdminAiSession(adminToken) {
  return {success:true,message:'សិទ្ធិ Admin ត្រឹមត្រូវ។',username:requireAdminAiSession_(adminToken)};
}

/**
 * សាកល្បងភ្ជាប់ MWAPI
 * ជ្រើសរើស testMWAPI ហើយចុច Run
 */
function testMWAPI() {

  var result = chatWithAI(
    [{
      role: 'user',
      content: 'សួស្តី! សូមណែនាំខ្លួនជាភាសាខ្មែរ។'
    }],
    ''
  );

  console.log(JSON.stringify(result, null, 2));

  return result;
}


/********************************************************
 * AI SMART DASHBOARD - READ ONLY / ADMIN TOKEN REQUIRED
 * មុខងារថ្មីនេះអាន Sheet Payment និង Teachers ប៉ុណ្ណោះ
 ********************************************************/
function aiGetDashboardData(adminToken, monthStr) {
  requireAdminAiSession_(adminToken);
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var month = String(monthStr || today.slice(0, 7));
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('ទម្រង់ខែមិនត្រឹមត្រូវ។');

  var payments = getPaymentData();
  var monthParts = month.split('-');
  var previousDate = new Date(Number(monthParts[0]), Number(monthParts[1]) - 2, 1);
  var previousMonth = Utilities.formatDate(previousDate, tz, 'yyyy-MM');
  var monthRows = payments.filter(function(r) { return String(r.paid_date || '').slice(0, 7) === month; });
  var previousMonthRows = payments.filter(function(r) { return String(r.paid_date || '').slice(0, 7) === previousMonth; });
  var todayRows = payments.filter(function(r) { return String(r.paid_date || '') === today; });
  function aggregate(rows) {
    var out = {records: rows.length, total: 0, teacher80: 0, school20: 0, cash: 0, qr: 0, students: {}, teachers: {}};
    rows.forEach(function(r) {
      var amount = Number(r.amount) || 0;
      out.total += amount;
      out.teacher80 += Number(r.teacher80) || 0;
      out.school20 += Number(r.school20) || 0;
      if (String(r.payment_method || '').trim().toUpperCase() === 'QR') out.qr += amount;
      else out.cash += amount;
      if (r.student_id) out.students[String(r.student_id)] = true;
      var name = String(r.teacher_name || 'មិនទាន់កំណត់');
      if (!out.teachers[name]) out.teachers[name] = {name:name, total:0, records:0, students:{}};
      out.teachers[name].total += amount;
      out.teachers[name].records++;
      if (r.student_id) out.teachers[name].students[String(r.student_id)] = true;
    });
    out.uniqueStudents = Object.keys(out.students).length;
    out.byTeacher = Object.keys(out.teachers).map(function(k) {
      var t = out.teachers[k];
      return {name:t.name,total:t.total,records:t.records,students:Object.keys(t.students).length};
    }).sort(function(a,b){return b.total-a.total;});
    delete out.students; delete out.teachers;
    return out;
  }

  var monthly = aggregate(monthRows);
  var previousMonthly = aggregate(previousMonthRows);
  var daily = aggregate(todayRows);
  var byDayMap = {};
  monthRows.forEach(function(r) {
    var day = String(r.paid_date || '').slice(8, 10);
    if (!day) return;
    if (!byDayMap[day]) byDayMap[day] = {day:day,total:0,records:0};
    byDayMap[day].total += Number(r.amount) || 0;
    byDayMap[day].records++;
  });
  var byDay = Object.keys(byDayMap).sort().map(function(k){return byDayMap[k];});
  return {
    success:true, generatedAt:Utilities.formatDate(new Date(),tz,'yyyy-MM-dd HH:mm:ss'),
    todayDate:today, selectedMonth:month, previousMonthLabel:previousMonth,
    today:daily, month:monthly, previousMonth:previousMonthly, byDay:byDay,
    topTeachers:monthly.byTeacher.slice(0,10),
    activeTeacherCount:getTeachersList().length
  };
}


/********************************************************
 * AI SCHOOL INSIGHTS - READ ONLY / ADMIN TOKEN REQUIRED
 * វិភាគតែទិន្នន័យសរុប មិនអានឈ្មោះ ឬ Student ID របស់សិស្ស
 ********************************************************/
function aiGenerateSchoolInsights(adminToken, monthStr, selectedModel) {
  requireAdminAiSession_(adminToken);

  var dashboard = aiGetDashboardData(adminToken, monthStr);
  if (!dashboard || dashboard.success !== true) {
    return {success:false, error:'មិនអាចទាញទិន្នន័យសម្រាប់វិភាគបានទេ។'};
  }

  var current = dashboard.month || {};
  var previous = dashboard.previousMonth || {};
  var currentTotal = Number(current.total) || 0;
  var previousTotal = Number(previous.total) || 0;
  var change = currentTotal - previousTotal;
  var changePct = previousTotal > 0 ? (change / previousTotal * 100) : null;

  var daily = (dashboard.byDay || []).map(function(x) {
    return {day:x.day, amount:Number(x.total)||0, records:Number(x.records)||0};
  });
  var teacherTop = (dashboard.topTeachers || []).slice(0, 10).map(function(x) {
    return {teacher:x.name, amount:Number(x.total)||0, payment_records:Number(x.records)||0};
  });

  var dailyAmounts = daily.map(function(x){return x.amount;}).filter(function(x){return x>0;});
  var dailyAverage = dailyAmounts.length
    ? dailyAmounts.reduce(function(a,b){return a+b;},0) / dailyAmounts.length : 0;
  var highestDay = daily.length ? daily.reduce(function(a,b){return b.amount>a.amount?b:a;}) : null;
  var lowestDay = dailyAmounts.length
    ? daily.filter(function(x){return x.amount>0;}).reduce(function(a,b){return b.amount<a.amount?b:a;})
    : null;

  var facts = {
    selected_month:dashboard.selectedMonth,
    previous_month:dashboard.previousMonthLabel,
    current_month_total_riel:currentTotal,
    previous_month_total_riel:previousTotal,
    change_riel:change,
    change_percent:changePct,
    current_month_payment_records:Number(current.records)||0,
    previous_month_payment_records:Number(previous.records)||0,
    unique_students_in_selected_month:Number(current.uniqueStudents)||0,
    active_teachers:Number(dashboard.activeTeacherCount)||0,
    teacher_share_80_riel:Number(current.teacher80)||0,
    school_share_20_riel:Number(current.school20)||0,
    cash_riel:Number(current.cash)||0,
    qr_riel:Number(current.qr)||0,
    daily_average_on_days_with_payments_riel:Math.round(dailyAverage),
    highest_revenue_day:highestDay,
    lowest_nonzero_revenue_day:lowestDay,
    daily_revenue:daily,
    top_10_teachers_by_revenue:teacherTop
  };

  var prompt = [
    'អ្នកជំនាញវិភាគទិន្នន័យហិរញ្ញវត្ថុសាលារៀន។ សូមវិភាគទិន្នន័យសរុបខាងក្រោមជាភាសាខ្មែរ។',
    'ត្រូវបែងចែកចម្លើយជា៖ ១) សង្ខេបស្ថានភាព ២) និន្នាការធៀបខែមុន ៣) ចំណុចគួរឱ្យចាប់អារម្មណ៍ ៤) អនុសាសន៍អនុវត្តបាន ៣ ចំណុច។',
    'ប្រើតែតម្លៃដែលមានក្នុង JSON។ កុំបង្កើតហេតុផល ឬទិន្នន័យដែលមិនមាន។ ប្រសិនបើខែនេះមិនទាន់ចប់ ឬទិន្នន័យមិនពេញលេញ សូមបញ្ជាក់ថាការប្រៀបធៀបអាចមិនស្មើគ្នាតាមចំនួនថ្ងៃ។',
    'កុំសន្និដ្ឋានថាគ្រូណាម្នាក់ល្អ ឬអន់ ដោយផ្អែកលើចំណូលតែប៉ុណ្ណោះ។ កុំបង្ហាញព័ត៌មានសម្គាល់សិស្ស។',
    'បញ្ជាក់ច្បាស់ថា AI ផ្តល់ជាការវិភាគជំនួយ មិនមែនជាការសម្រេចចិត្តហិរញ្ញវត្ថុដោយស្វ័យប្រវត្តិ។',
    'ចំនួនទឹកប្រាក់ទាំងអស់គិតជារៀល។ កុំប្រើ Markdown table។',
    'ទិន្នន័យ JSON៖ ' + JSON.stringify(facts)
  ].join('\n');

  var ai = chatWithAI([{role:'user', content:prompt}], selectedModel);
  if (!ai || ai.success !== true) {
    return {success:false, error:ai && ai.error ? ai.error : 'MWAPI មិនអាចបង្កើតការវិភាគបានទេ។'};
  }

  return {
    success:true,
    answer:ai.answer,
    model:ai.model || '',
    month:dashboard.selectedMonth,
    generatedAt:Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
}


/********************************************************
 * AI SMART DASHBOARD - MONTHLY REPORT ARCHIVE
 * Creates/uses a separate archive sheet; never edits Payment.
 ********************************************************/
function aiSaveMonthlyReport(adminToken, monthStr, reportText, summaryJson) {
  var adminUser = requireAdminAiSession_(adminToken);
  var month = String(monthStr || '').trim();
  var report = String(reportText || '').trim();
  var summary = String(summaryJson || '').trim();
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('ទម្រង់ខែមិនត្រឹមត្រូវ។');
  if (!report) throw new Error('មិនមានលទ្ធផល AI សម្រាប់រក្សាទុកទេ។');
  if (report.length > 40000) throw new Error('របាយការណ៍វែងពេកសម្រាប់រក្សាទុក។');
  if (summary.length > 20000) throw new Error('ទិន្នន័យសង្ខេបវែងពេក។');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('AI_Report_Archive');
  if (!sheet) {
    sheet = ss.insertSheet('AI_Report_Archive');
    sheet.getRange(1, 1, 1, 5).setValues([['Saved At', 'Month', 'Admin', 'AI Report', 'Dashboard Summary JSON']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 170); sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 150); sheet.setColumnWidth(4, 600); sheet.setColumnWidth(5, 400);
    sheet.getRange('D:E').setWrap(true);
  }
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 5).setValues([[new Date(), month, adminUser, report, summary]]);
  sheet.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  return { success: true, row: row, sheet: 'AI_Report_Archive', month: month };
}
