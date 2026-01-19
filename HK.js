function hk_inlinemail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Churn list (main)_HK');
  var docId = "1-pEuHGFhF0Bw5ampVmL6bjDKVvsZBweb5yG3fdsAkTE";  // template
  
  var aVals = sheet.getRange("A:A").getValues();
  var lastRow = 0;
  for (var i = aVals.length - 1; i >= 0; i--) {
    if (aVals[i][0] != "") { lastRow = i + 1; break; }
  }

  for (var i = lastRow; i >= 3; i--) {
    var readyToSent = sheet.getRange(i, 21).getValue(); 
    var timestamp = sheet.getRange(i, 22).getValue();  

    if (readyToSent === true && (timestamp === "" || timestamp === null)) {
      
      var rowData = sheet.getRange(i, 1, 1, 20).getValues()[0];
      var resto = rowData[1];      
      var am_name = rowData[2];    
      var service = rowData[3];   
      var recipient = rowData[11]; 
      var actual_term_date = sheet.getRange(i, 13).getDisplayValue(); 
      var system_closure_date = sheet.getRange(i, 14).getDisplayValue(); 
      var form_link = sheet.getRange(i, 20).getDisplayValue(); 

      // --- AM 簽名檔 ---
      var amInfo = getAmInfo(ss, am_name);
      var info_name = amInfo.fullname;
      var info_tel = amInfo.tel;
      var info_email = amInfo.email;

      // --- 判斷是否有訂金項目 ---
      var hasDeposit = (rowData[4] !== "" || rowData[5] !== "");
      
      // --- 選擇範本標籤 ---
      var startTag, endTag;
      if (service === "Add-ons") {
        startTag = "[[AddonOnly_TEMPLATE]]";
        endTag = "[[END_AddonOnly]]";
        mailSubject = "inline | " + resto + " - Termination Notice 終止通知 (Add-on Feature 附加功能)";  // 加值服務解約標題
      } else if (hasDeposit) {
        startTag = "[[TMS+Deposit_TEMPLATE]]";
        endTag = "[[END_TMS+Deposit]]";
        mailSubject = "inline | " + resto + " - Termination Notice 終止通知";  // TMS 解約標題
      } else {
        startTag = "[[TMS_TEMPLATE]]";
        endTag = "[[END_TMS]]";
        mailSubject = "inline | " + resto + " - Termination Notice 終止通知";  // TMS 解約標題
      }

      // --- 從文件擷取範本 ---
      var template = extractTemplate(docId, startTag, endTag);
      
      // --- 變數替換對照表 ---
      var reps = {
        "{{CSM/BD}}": am_name,
        "{{System_Closure_Date}}": system_closure_date,
        "{{Actual_Termination_Date}}": actual_term_date,
        "{{Link}}": form_link,
        "{{Info_Email}}": info_email,
        "{{Info_Name}}": info_name,
        "{{Info_Tel}}": info_tel
      };

      // --- 處理 AddonOnly 的加值服務名稱 ---
      if (service === "Add-ons") {
        var addons = getAddonDetails(i, sheet);
        reps["{{Addon_Name_EN}}"] = addons.en;
        reps["{{Addon_Name_CH}}"] = addons.ch;
        
        // 如果沒有訂金，移除報表說明段落（注意！這邊需與模板內文一模一樣！）
        if (!hasDeposit) {
          var depTextEn = "If you have been using the deposit feature and need to export the deposit report, please complete it via our backhouse - myinline.com before the termination date. For detailed steps, please refer to: https://help.inline.app/en/articles/5660673.";
          var depTextZh = "若您曾使用訂金功能並需匯出相關報表，請務必在功能關閉日前透過我們的後台 - myinline.com 完成。 詳細步驟請參閱：https://help.inline.app/zh-TW/articles/5660673。";
          template = template.replace(new RegExp(depTextEn, "g"), "").replace(new RegExp(depTextZh, "g"), "");
        }
      }

      // --- 執行變數替換 ---
      for (var key in reps) {
        var re = new RegExp(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
        template = template.replace(re, reps[key]);
      }

      // --- 處理粗體格式 (把 **文字** 轉換成 <b>文字</b>) ---
      template = template.replace(/\*\*([\s\S]*?)\*\*/g, "<b>$1</b>");

      // --- 換行符號轉 HTML ---
      var finalHtmlBody = template.replace(/\n/g, "<br>");

      // --- 寄信 ---
      MailApp.sendEmail({
        to: recipient,
        name: 'Inline Customer Success Team',
        cc: info_email,
        subject: mailSubject,
        htmlBody: "<div style='font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5;'>" + finalHtmlBody + "</div>"
      });

      sheet.getRange(i, 22).setValue(new Date());
      console.log("✅ 成功寄出：" + resto);
    }
  }
}

// 擷取標籤內文字
function extractTemplate(docId, startTag, endTag) {
  var doc = DocumentApp.openById(docId);
  var text = doc.getBody().getText();
  var regex = new RegExp(escapeRegExp(startTag) + "([\\s\\S]*?)" + escapeRegExp(endTag));
  var match = text.match(regex);
  return match ? match[1].trim() : "Template Not Found";
}

// 動態搜尋 information 工作表
function getAmInfo(ss, name) {
  var infoSheet = ss.getSheetByName('information');
  var data = infoSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === name.toString().trim()) {
      return {
        fullname: data[i][1],  // B 欄
        tel: data[i][2],       // C 欄
        email: data[i][3]      // D 欄
      };
    }
  }
  return { fullname: name, tel: "", email: "" };
}

function getAddonDetails(row, sheet) {
  var cols = [
    {n:'Bank Transfer Deposit', c:'銀行轉帳預付訂金', col:5},
    {n:'Credit Card Deposit', c:'信用卡訂金', col:6},
    {n:'Call Transfer', c:'電話轉接', col:7},
    {n:'Survey', c:'問卷調查', col:8},
    {n:'Voucher', c:'優惠券', col:9},
    {n:'Meituan/Dianping Integration', c:'美團/大眾點評串接', col:10},
    {n:'WhatsApp AI CS Chatbot', c:'WhatsApp AI 客服小幫手', col:11}
  ];
  var en = [], ch = [];
  cols.forEach(function(d){
    if(sheet.getRange(row, d.col).getValue() !== ""){
      en.push(d.n); ch.push(d.c);
    }
  });
  return {
    en: en.join('／'),
    ch: ch.join('／')
  };
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}