function sendEmails() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet1 = ss.getSheetByName('ADPYM');
var sheet2 = ss.getSheetByName('DPYM');
var sheet3 = ss.getSheetByName('導師list');
var materialsSheet = ss.getSheetByName('materials');

// 使用者提供的測試 Email
// var toList = ['karry561@gmail.com','karry561@gmail.com'];
// var ccList = ['karry561@gmail.com','karry561@gmail.com','karry561@gmail.com'];
var toList = ['dorachan@hkma.org.hk', 'hannahsit@hk-ma.org.hk']; // Fixed to list for email1 and email3，HKMA兩位收信人 <hannahshit>
var ccList = ['fcjim@hongyip.com', 'kllam1@hongyip.com','heleeerle@hongyip.com']; // Fixed cc list for various emails
processSheet(sheet1, 'HKMA 物業管理高級文憑', false, ss, sheet3, materialsSheet, toList, ccList);
processSheet(sheet2, 'HKMA 物業管理文憑', true, ss, sheet3, materialsSheet, toList, ccList);

Logger.log('Remaining daily email quota: ' + MailApp.getRemainingDailyQuota()); // Log quota
}

function parseSheetCellToDate(cell, displayValue, timeZone) {
if (displayValue) {
var m = String(displayValue).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
if (m) {
return new Date(
Number(m[3]),
Number(m[2]) - 1,
Number(m[1]),
12, 0, 0, 0
);
}
}

if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
  var ymd = Utilities.formatDate(cell, timeZone, 'yyyy-MM-dd').split('-');
  return new Date(
  Number(ymd[0]),
  Number(ymd[1]) - 1,
  Number(ymd[2]),
  12, 0, 0, 0
  );
  }

return null;
}

// 輔助函式：簡單格式化日期 (被 processCourse 用於 Logger.log)
function formatDate(date) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // 輸出時，確保只輸出 YYYY-MM-DD 格式，不顯示時間
    return Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}




// ===============================================
// Function to process each sheet (Sheet1 or Sheet2), parse courses and chapters
// ===============================================
function processSheet(sheet, programTitle, isDiploma, ss, sheet3, materialsSheet, toList, ccList) {
    
    // 限制讀取範圍，避免讀取到表單底部無關或錯誤的日期數據
    var lastRow = sheet.getLastRow();
    // 數據安全上限為 200 行，請根據您的實際情況調整！
    var maxDataRow = Math.min(lastRow, 200); 
    var lastCol = sheet.getLastColumn();
    
    var data = sheet.getRange(1, 1, maxDataRow, lastCol).getValues(); 
    var displayData = sheet.getRange(1, 1, maxDataRow, lastCol).getDisplayValues();
    var fontWeights = sheet.getRange(1, 1, maxDataRow, lastCol).getFontWeights(); 
    var backgrounds = sheet.getRange(1, 1, maxDataRow, lastCol).getBackgrounds(); 

    var timeZone = ss.getSpreadsheetTimeZone(); 
    var currentCourse = null;

    for (var row = 1; row < data.length; row++) {
        var isBold = fontWeights[row][0] === 'bold';
        var aVal = data[row][0];
        var isGrey = backgrounds[row][0] !== '#ffffff';

        if (isBold && aVal) {
            if (currentCourse) {
                processCourse(currentCourse, sheet, programTitle, isDiploma, sheet3, materialsSheet, toList, ccList);
            }
            if (!isGrey) {
                currentCourse = {
                    name: aVal.replace(/\s*\d+$/, ''),
                    sessions: [],
                    tutors: new Set(),
                    statusRow: row + 1,
                    confirmRow: null,
                    lastRow: null,
                    firstDate: null // 初始化為 null
                };
            } else {
                currentCourse = null; 
            }
        } else if (currentCourse && aVal) {
            var chapter = aVal;
            var tutorCol = isDiploma ? 2 : 1; 
            var tutor = data[row][tutorCol];
            var startCol = isDiploma ? 3 : 2; 
            var numClasses = isDiploma ? 5 : 8; 
            var displayDates = [];
            var dateColors = [];

            for (var col = startCol; col < startCol + numClasses; col++) {
                var cellValue = data[row][col];
                var displayValue = displayData[row][col].trim();
                var color = backgrounds[row][col];
                var date = null;
                var subTutor = null;

                var subTutorMatch = displayValue.match(/\(([^)]+)\)/);
                if (subTutorMatch) {
                    subTutor = subTutorMatch[1];
                }

                // *** 使用修正後的日期解析輔助函式 ***
                var parsed = parseSheetCellToDate(cellValue, displayValue, timeZone);

                if (parsed) {
                    date = parsed;
                    Logger.log(                  //查看parseDate和utf
                      'firstDate update | course=' + currentCourse.name +
                      ' - row=' + (row + 1) +
                      ' - col=' + (col + 1) +
                      ' - raw=' + cellValue + 
                      ' /n - display=' + displayValue +
                      ' /n- parsed=' + Utilities.formatDate(parsed, timeZone, 'yyyy-MM-dd')
                      );
                    // 更新 course.firstDate (取最小日期)
                    if (!currentCourse.firstDate || date.getTime() < currentCourse.firstDate.getTime()) {
                        // 這裡使用 new Date(date.getTime()) 創建副本，確保日期物件是全新的
                        currentCourse.firstDate = new Date(date.getTime());
                    }
                }

                var effTutor = subTutor || data[row][tutorCol];
                if (effTutor) currentCourse.tutors.add(effTutor);

                displayDates.push(displayValue);
                dateColors.push(color);
            }
            
            currentCourse.sessions.push({
                chapter: chapter,
                tutor: tutor || '',
                dates: displayDates,
                colors: dateColors
            });
            
            if (currentCourse.confirmRow === null) {
                currentCourse.confirmRow = row + 1;
            }
            currentCourse.lastRow = row + 1;
        }
    }
    
    if (currentCourse) {
        processCourse(currentCourse, sheet, programTitle, isDiploma, sheet3, materialsSheet, toList, ccList);
    }
}


// ===============================================
// Function to process each course and send emails based on conditions
// ===============================================
function processCourse(course, sheet, programTitle, isDiploma, sheet3, materialsSheet, toList, ccList) {
    // 檢查 firstDate 是否已設定 (防止 undefined 錯誤)
    if (!course.firstDate || !course.lastRow) {
        Logger.log('Skipping course ' + course.name + ': First date is undefined or no chapters found.');
        return; 
    }

    // ==========================================
    //  欄位設定 (確保 DPYM 和 ADPYM 的欄位索引正確)
    // ==========================================
    var statusCols, hkmaConfirmCol;

    if (isDiploma) {
        // Sheet 2 (DPYM): Email 1-4: L(12), M(13), N(14), O(15). HKMA Confirm K(11)
        statusCols = [12, 13, 14, 15]; 
        hkmaConfirmCol = 11; 
    } else {
        // Sheet 1 (ADPYM): Email 1-4: N(14), O(15), P(16), Q(17). HKMA Confirm M(13)
        statusCols = [14, 15, 16, 17]; 
        hkmaConfirmCol = 13; 
    }
    // ==========================================

    var tutorMap = getTutorMap(sheet3);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timeZone = ss.getSpreadsheetTimeZone();
    
    // ** 1. 正常運行模式 (使用當前日期) **
    var currentDateStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    var currentParts = currentDateStr.split('-');
    var currentDate = new Date(
    Number(currentParts[0]),
    Number(currentParts[1]) - 1,
    Number(currentParts[2]),
    12, 0, 0, 0
    );
    
    // ** 2. 測試模式 (使用固定日期) - 請根據需要切換 **
    // var currentDate = new Date("2026-01-12T00:00:00+08:00"); 
    // ----------------------------------------------------
    
    var statusRow = course.lastRow;

    Logger.log('Processing course ' + course.name + ' (Sheet: ' + (isDiploma ? 'DPYM' : 'ADPYM') + ') at row ' + statusRow);

    // 1. Check HKMA Confirmation 
    var rawConfirm = sheet.getRange(statusRow, hkmaConfirmCol).getValue();
    var hkmaConfirm = String(rawConfirm || '').trim();
    
    if (hkmaConfirm !== '√' && hkmaConfirm !== '✔') {
        Logger.log('Skipping ' + course.name + ': HKMA Confirm is "' + hkmaConfirm + '" in column ' + hkmaConfirmCol + '.');
        return;
    }

    var propKey = sheet.getName() + '_' + course.name + '_N_set_date';
    var props = PropertiesService.getScriptProperties();

    // 2. Email 1 (Confirm) - statusCols[0]
    var col1 = statusCols[0];
    if (sheet.getRange(statusRow, col1).getValue() === '') {
        sendScheduleEmail(course, programTitle, isDiploma, tutorMap, 1, toList, ccList); 
        sheet.getRange(statusRow, col1).setValue('✔'); 
        props.setProperty(propKey, currentDate.getTime().toString()); 
    }

    // Get N set date
    var nSetTimestamp = props.getProperty(propKey);
    var nSetDate = nSetTimestamp ? new Date(parseInt(nSetTimestamp)) : null;
    if (nSetDate) nSetDate.setHours(0, 0, 0, 0);

    // 3. Email 2 (Materials to Tutors) - statusCols[1]
    var col2 = statusCols[1];
    if (sheet.getRange(statusRow, col2).getValue() === '' &&
        sheet.getRange(statusRow, col1).getValue() !== '') { 
        sendMaterialsEmail(course, programTitle, isDiploma, tutorMap, materialsSheet, 2, toList, ccList); 
        sheet.getRange(statusRow, col2).setValue('✔');
          Logger.log("Email 2 is sending...")
    }

    // 4. Email 3 (Materials to HKMA) - statusCols[2]
    var col3 = statusCols[2];
    if (sheet.getRange(statusRow, col3).getValue() === '' && sheet.getRange(statusRow, col2).getValue() !== '') {
            sendMaterialsEmail(course, programTitle, isDiploma, tutorMap, materialsSheet, 3, toList, ccList); 
            sheet.getRange(statusRow, col3).setValue('✔');
            Logger.log("Email 3 is sending...")

    }

    // 5. Email 4 (Reminder) - statusCols[3]
    var col4 = statusCols[3];
    var status1 = sheet.getRange(statusRow, col1).getValue();
    var status2 = sheet.getRange(statusRow, col2).getValue();
    var status4 = sheet.getRange(statusRow, col4).getValue();

    if (status4 === '' && status1 !== '' && status2 !== '') {
        
        var reminderDate = new Date(course.firstDate.getTime()); // 確保使用副本
        reminderDate.setDate(reminderDate.getDate() - 3);
        // 因為 course.firstDate 已經被 parseSheetCellToDate 修正為 T12:00:00，這裡只需確保 reminderDate 的時間也設為 T00:00:00 進行比較
        reminderDate.setHours(0, 0, 0, 0); 

        var firstClassDate = new Date(course.firstDate.getTime());
        firstClassDate.setHours(0, 0, 0, 0);

        Logger.log('🔍 Checking Email 4 for ' + course.name + ':');
        Logger.log(' - Current Date: ' + formatDate(currentDate));
        Logger.log(' - First Class Date: ' + formatDate(firstClassDate));
        Logger.log(' - Reminder Date (First-3): ' + formatDate(reminderDate));

        // 條件檢查：今日達到或超過提醒日，且在開課日之前
        if (currentDate.getTime() >= reminderDate.getTime() && currentDate.getTime() < firstClassDate.getTime()){
            Logger.log('✅ Date condition MET! Sending Email 4...');
            sendScheduleEmail(course, programTitle, isDiploma, tutorMap, 4, toList, ccList); 
            sheet.getRange(statusRow, col4).setValue('✔');
        }
        else if(currentDate.getTime() > reminderDate.getTime()){
          Logger.log('The date is too late');
        }else if (currentDate.getTime() < reminderDate.getTime()){
          Logger.log('The date is too early');
        }
    } else {
        if (status4 !== '') Logger.log('Skipping Email 4: Already sent (Status in column ' + col4 + ').');
        if (status1 === '') Logger.log('Skipping Email 4: Email 1 not sent yet (column ' + col1 + ' empty).');
        if (status2 === '') Logger.log('Skipping Email 4: Email 2 not sent yet (column ' + col2 + ' empty).');
    }
}

// Function to get tutor map from sheet3...
function getTutorMap(sheet3) {
var data = sheet3.getRange('A2:E30').getValues(); // Specific range for tutor data
var map = {};
for (var r = 0; r < data.length; r++) {
var name = data[r][0];
var email = data[r][4];
if (name && email) { // Only add if name and email present, skip non-tutor rows
map[name] = {
title: data[r][1],
tel: data[r][2],
email: email
};
}
}
return map;
}

// Function to send schedule emails... (omitted for brevity, keep the original content)
function sendScheduleEmail(course, programTitle, isDiploma, tutorMap, emailType, toList, ccList) {
var tutors = Array.from(course.tutors);
var tutorRecipients = tutors.map(t => ({ name: t + (tutorMap[t] ? ' ' + tutorMap[t].title : ''), email: tutorMap[t] ? tutorMap[t].email : '' })).filter(r => r.email);

var recipients = tutorRecipients;
var ccRecipients = [];

if (emailType === 1) {
recipients = recipients.concat(toList.map(email => ({ name: '', email: email })));
ccRecipients = ccList.map(email => ({ name: '', email: email })); // all 3
} else if (emailType === 4) {
ccRecipients = ccList.slice(0, 2).map(email => ({ name: '', email: email })); // cc[0], cc[1]
}

var subject = (emailType === 1) ? '【確認授課】' + programTitle + '-' + course.name : '【提醒授課】' + programTitle + '-' + course.name;

var greeting = '致 ' + tutors.join('、') + ':';

var body = greeting + '<br><br>' +
'感謝閣下答應到教授課程,詳情如下:<br><br>' +
'課程名稱：' + programTitle + ' - ' + course.name + '<br>' +
'日期: 見下表<br>' +
'時間: 19:00-22:00<br>' +
'地點: 九龍尖沙咀麼地道 75 號 南洋中心第二座 3 樓<br>' +
'聯絡電話: 2574-9346 (香港管理專業協會)<br>' +
'導師聯絡:<br>';

tutors.forEach(t => {
if (tutorMap[t]) {
body += '• ' + t + ' （' + tutorMap[t].tel + '）<br>';
}
});

body += '<br>' + buildScheduleTable(programTitle, course.name, course.sessions, programTitle + ' - ' + course.name) +
'<p>如有任何查詢，請致電： 2523-9313 與李晧而小姐(Miss. Erle Lee)或2523-9363與林敬樂先生(Mr. Pius Lam)聯絡，謝謝!</p><br>';

if (isDiploma) {
body += '';
}

body += '<br><br>Best regards,<br>Erle Lee<br>Assistant Training Officer<br>HR - Training & Development<br>Hong Yip Service Company Ltd<br>Tel:2523 9313<br>Fax: 2523 9707';

// body += '<br><br><img src="https://drive.google.com/uc?export=view&id=1Z77cSNbWbu1DEZoXmdsmvPL8hQup4ih_" alt="Company Logo" width="200" height="auto">';

// Extract emails as comma-separated strings
var toEmails = recipients.map(r => r.email).filter(e => e).join(',');
var ccEmails = ccRecipients.map(r => r.email).filter(e => e).join(',');

// Send via Gmail
var obj = {
to: toEmails,
cc: ccEmails,
subject: subject,
htmlBody: body
};
try {
MailApp.sendEmail(obj);
Logger.log('Email sent: ' + subject + ' for course ' + course.name);
} catch (e) {
Logger.log('Error sending email: ' + subject + ' for course ' + course.name + '. Error: ' + e.message);
}
}


// Function to send materials emails... (omitted for brevity, keep the original content)
function sendMaterialsEmail(course, programTitle, isDiploma, tutorMap, materialsSheet, emailType, toList, ccList) {
  var tutors = Array.from(course.tutors);
  var tutorRecipients = tutors.map(t => ({ name: t + (tutorMap[t] ? ' ' + tutorMap[t].title : ''), email: tutorMap[t] ? tutorMap[t].email : '' })).filter(r => r.email);

  var recipients = (emailType === 2) ? tutorRecipients : toList.map(email => ({ name: '', email: email }));
  var ccRecipients = [];
  if (emailType === 2) {
    ccRecipients = [{ name: '', email: ccList[1] }];
  } else if (emailType === 3) {
    ccRecipients = [{ name: '', email: ccList[1] }, { name: '', email: ccList[2] }];
  }

  var subject = '【教材】' + programTitle + '-' + course.name;

  var greeting = (emailType === 2) ? '致 ' + tutors.join('、') + ':' : '致 各位';

  var body = greeting + '<br><br>' +
    '教學材料已上傳到以下雲端連結，閣下請自行下載<br><br>';

  var materialsData = materialsSheet.getDataRange().getValues();
  var filledData = [];
  var currCourse = '';
  
  // **Email 2 修正 A: 確保 DPYM 結構（章節為空，但連結有值）的行不會被排除**
  for (var r = 0; r < materialsData.length; r++) {
    if (materialsData[r][0] && materialsData[r][0].trim() !== '') {
      currCourse = materialsData[r][0].trim();
    }
    var chap = materialsData[r][1] ? String(materialsData[r][1]).trim() : ''; 
    
    // 新邏輯：只要有課程名稱，且連結欄位有值，就納入 filledData
    if (currCourse && (chap || materialsData[r][2] || materialsData[r][3])) { 
      filledData.push([currCourse, chap, materialsData[r][2], materialsData[r][3]]);
    }
  }

  var links = [];
  var courseLinkFound = false;

  course.sessions.forEach(session => {
    for (var i = 0; i < filledData.length; i++) {
      var materialsCourseName = filledData[i][0];
      var materialsChapterName = filledData[i][1];

      if (materialsCourseName === course.name) {
        var isMatch = false;

        // 情境 1 (DPYM): Chapter 為空 (單連結)，且尚未被添加
        if (materialsChapterName === '' && !courseLinkFound) {
          isMatch = true;
          courseLinkFound = true; 
        } 
        // 情境 2 (ADPYM): Chapter 不為空，必須精確匹配 session.chapter
        else if (materialsChapterName !== '' && materialsChapterName === session.chapter) {
          isMatch = true; 
        }

        if (isMatch) {
          var linkCol = (emailType === 2) ? 2 : 3;
          var link = filledData[i][linkCol];
          if (link) {
            var displayText = materialsChapterName || course.name;
            links.push(displayText + ':<a href="' + link + '">' + link + '</a><br>');
            
            // 對於 DPYM 結構，找到一次連結後即可跳出內層循環
            if (materialsChapterName === '') {
              break;
            }
          }
        }
      }
    }
  });

  body += links.join('');
  body += '<br>如有任何查詢，請致電： 2523-9313 與李晧而小姐(Miss. Erle Lee)或2523-9363與林敬樂先生(Mr. Pius Lam)聯絡，謝謝!<br>';
  if (isDiploma) {
    body += '';
  }

  body += '<br>Best regards,<br>Erle Lee<br>Assistant Training Officer<br>HR - Training & Development<br>Hong Yip Service Company Ltd<br>Tel: 2523 9313<br>Fax: 2523 9707';

  body += '<br><br><img src="https://www.hongyip.com/themes/custom/hongyip_theme/assets/img/logo-m.png" alt="Hong Yip Service Company Ltd" width="200" height="auto">';

  // Extract emails as comma-separated strings
  var toEmails = recipients.map(r => r.email).filter(e => e).join(',');
  var ccEmails = ccRecipients.map(r => r.email).filter(e => e).join(',');

  // Send via Gmail
  var obj = {
    to: toEmails,
    cc: ccEmails,
    subject: subject,
    htmlBody: body
  };
  try {
    MailApp.sendEmail(obj);
    Logger.log('Email sent: ' + subject + ' for course ' + course.name);
  } catch (e) {
    Logger.log('Error sending email: ' + subject + ' for course ' + course.name + '. Error: ' + e.message);
  }
}


// Function to build HTML table for schedule...
function buildScheduleTable(programTitle, courseName, sessions, title) {
// 修正：使用 Array.isArray 檢查 sessions
if (!sessions || !Array.isArray(sessions) || sessions.length === 0) {
Logger.log('No sessions for table, skipping.');
return ''; 
}

// Determine max non-empty columns
var maxCols = 0;
sessions.forEach(session => {
for (var i = session.dates.length - 1; i >= 0; i--) {
if (session.dates[i].trim() !== '') {
maxCols = Math.max(maxCols, i + 1);
break;
}
}
});

if (maxCols === 0) return ''; // No dates, skip table

var html = '<table border="1"><tr><th colspan="' + (maxCols + 2) + '">' + programTitle + '-' + courseName + '</th></tr><tr><td>課程</td><td>導師</td>';
for (var col = 1; col <= maxCols; col++) {
html += '<td>課堂' + col + '</td>';
}
html += '</tr>';

sessions.forEach(session => {
html += '<tr><td>' + (session.chapter || '') + '</td><td>' + (session.tutor || '') + '</td>';
for (var i = 0; i < maxCols; i++) {
var dateStr = (session.dates[i] || '').trim();
var color = session.colors[i] !== '#ffffff' ? 'background-color: ' + session.colors[i] + ';' : '';
html += '<td style="' + color + '">' + dateStr + '</td>';
}
html += '</tr>';
});

html += '</table>';
return html;
}
