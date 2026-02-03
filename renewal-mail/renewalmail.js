function sendRenewalEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("testing");
  const csmSheet = ss.getSheetByName("「info」的副本");
  const lastRow = mainSheet.getLastRow();

  if (lastRow < 3) return;

  const csmData = csmSheet.getDataRange().getValues();
  const csmMap = {};
  for (let i = 0; i < csmData.length; i++) {
    csmMap[csmData[i][0].toString().trim()] = {
      fullname: csmData[i][1],
      title: csmData[i][2],
      mobile: csmData[i][3],
      email: csmData[i][4]
    };
  }

  const range = mainSheet.getRange(1, 1, lastRow, 26);
  const data = range.getValues();

  // Plan Type 中文對照表
  const planTypeMap = {
    "Standard": "標準",
    "Light": "輕量版",
    "Premium": "尊尚",
    "Email-only": "基本",
    "Unbundled": "Unbundled"
  };

  for (let i = 2; i < lastRow; i++) {
    const row = data[i];
    const isChecked = row[23]; // X 欄
    const sentStatus = row[24]; // Z 欄

    if (isChecked !== true || (sentStatus !== "" && sentStatus !== null)) continue;

    const restaurantName = row[1]; // B
    const csmName = row[2].toString().trim(); // C
    const oldStartDate = formatDate(row[3]); // D
    const currentEndDate = formatDate(row[4]); // E
    const planType = row[5]; // F
    const isDDA = row[6]; // G
    const newStartDate = formatDate(row[7]); // H
    const billingCycle = row[9]; // J
    const rsvSets = row[10]; // K
    const rsvPrice = row[11]; // L

    // 新增 Overage / Notification 費用欄位 (M, N, O, P)
    const m_overage = row[12]; // M
    const n_whatsapp = row[13]; // N
    const o_sms = row[14]; // O
    const p_call = row[15]; // P

    const recipientEmail = row[22]; // W (現在 ID 歸檔後的 Email 位在 V)

    if (!recipientEmail) continue;

    // --- 方案費用描述---
    let overageTextEN = "";
    let overageTextCH = "";

    let partsEN = [];
    let partsCH = [];
    const lowerPlan = planType.toLowerCase();

    // 1. 處理 M 欄 (Standard 系列)
    if (lowerPlan.includes("standard") || lowerPlan.includes("light") || lowerPlan.includes("premium")) {
      if (m_overage && m_overage !== 0 && m_overage !== "") {
        partsEN.push(`overage fee is HK$${m_overage}/credit`);
        partsCH.push(`超額費為 HK$${m_overage}/組`);
      }
    }

    // 2. 處理 N, O, P 欄 (通知費系列)
    let notifItemsEN = [];
    let notifItemsCH = [];

    // N, O 與方案類型掛鉤
    if (lowerPlan.includes("unbundled")) {
      if (n_whatsapp && n_whatsapp !== 0 && n_whatsapp !== "") {
        notifItemsEN.push(`HK$${n_whatsapp}/WhatsApp`);
        notifItemsCH.push(`HK$${n_whatsapp}/WhatsApp`);
      }
      if (o_sms && o_sms !== 0 && o_sms !== "") {
        notifItemsEN.push(`HK$${o_sms}/SMS`);
        notifItemsCH.push(`HK$${o_sms}/短訊`);
      }
    }

    // P 欄獨立判斷
    if (p_call && p_call !== 0 && p_call !== "") {
      notifItemsEN.push(`HK$${p_call} /queuing call`);
      notifItemsCH.push(`HK$${p_call}/候位來電`);
    }

    // 3. 如果有通知費項目，統一加上開頭
    if (notifItemsEN.length > 0) {
      // 英文組合：A, B & C
      const enList = notifItemsEN.join(", ").replace(/, ([^,]*)$/, " & $1");
      partsEN.push(`notifications fee is ${enList}`);

      // 中文組合：A、B、C
      const chList = notifItemsCH.join("、").replace(/、([^、]*)$/, "，及 $1");
      partsCH.push(`通知費為 ${chList}`);
    }

    // 4. 合併
    if (partsEN.length > 0) {
      overageTextEN = `, ${partsEN.join(" & ")}`;
      overageTextCH = `，${partsCH.join("，")}`;
    }

    // --- 加值功能邏輯 (現在從 Q 欄/索引 16 開始，往後推 6 欄) ---
    const addonDetails = [
      { en: "Bank Transfer Deposit", ch: "銀行轉帳預付訂金", preEN: "HK$", postEN: " per transaction", preCH: "每筆交易 HK$", postCH: "" },
      { en: "Credit Card Deposit", ch: "信用卡訂金", preEN: "HK$", postEN: "<i> (transaction fees apply)</i>", preCH: "HK$", postCH: "<i>（需支付交易手續費）</i>" },
      { en: "Call Transfer", ch: "電話轉接功能", preEN: "HK$", postEN: "", preCH: "HK$", postCH: "" },
      { en: "Meituan/Dianping Integration", ch: "美團/大眾點評串接", preEN: "HK$", postEN: " per seated cover", preCH: "HK$", postCH: "/入座人數" },
      { en: "Survey / Thank-you message", ch: "問卷調查功能／自動感謝信息", preEN: "HK$", postEN: "", preCH: "HK$", postCH: "" },
      { en: "WhatsApp AI CS Chatbot", ch: "WhatsApp AI 客服小幫手", preEN: "HK$", postEN: "", preCH: "HK$", postCH: "" }
    ];

    let addonContentEN = "";
    let addonContentCH = "";
    let hasAddonData = false;

    for (let k = 0; k < 6; k++) {
      let val = row[16 + k]; // 從 Q 欄開始
      
      if (val !== "" && val !== null && val !== 0) {
        hasAddonData = true;
        addonContentCH += `✓ ${addonDetails[k].ch} - ${addonDetails[k].preCH}${val}${addonDetails[k].postCH}<br>`;
        addonContentEN += `✓ ${addonDetails[k].en} - ${addonDetails[k].preEN}${val}${addonDetails[k].postEN}<br>`;
      } else {
        addonContentCH += `□ ${addonDetails[k].ch}<br>`;
        addonContentEN += `□ ${addonDetails[k].en}<br>`;
      }
    }

    const addonSectionEN = hasAddonData ? `<p>Add-on Features:<br>${addonContentEN}</p>` : "";
    const addonSectionCH = hasAddonData ? `<p>加值功能：<br>${addonContentCH}</p>` : "";

    const csmInfo = csmMap[csmName] || { fullname: csmName, title: "Customer Success Manager", mobile: "", email: "" };
    const planCH = planTypeMap[planType] || planType;

    const subject = `inline | ${restaurantName}  Automatic Renewal Notice 自動續約通知`;

    let ddaSectionEN = isDDA === true ? `<p>To make payments easier and more convenient, we are now providing <b>Direct Debit Authorisation (DDA)</b>. By enrolling in DDA, subscription fees will be transferred automatically on a recurring basis. If you are interested, please reply to this email and we will contact you with further details.</p>` : "";
    let ddaSectionCH = isDDA === true ? `<p>為了簡化付款程序，輕鬆處理帳單，我們現在提供 <b>直接付款授權（DDA）</b> 申請服務。通過申請 DDA，訂閱費用將定期自動轉帳。若您對此感興趣，請回覆此郵件，我們將與您聯繫並提供進一步詳情。</p>` : "";

    const htmlBody = `
      <div style="font-family: 'Times New Roman', Times, serif; line-height: 1.6; color: #333; font-size: 15px;">
        <p>To Whom It May Concern,</p>
        <div style="text-align: center; margin: 20px 0;">
          <h2 style="font-weight: bold; font-size: 1.25em;">inline Apps Automatic Renewal Notice</h2>
        </div>
        <p>inline and <b>${restaurantName}</b> entered into the agreement for the provision of restaurant software services. The current agreement term is effective from ${oldStartDate} and ends on ${currentEndDate}.</p>
        <p><b>[${billingCycle === 'monthly' ? 'Monthly' : 'Annual'} ${planType} Plan]</b><br>
        Subscription Details:<br>
        HK$${rsvPrice} including ${rsvSets} reservation/queuing credits${overageTextEN}</p>
        
        ${addonSectionEN}

        <p>As agreed in the agreement, the subscription shall be <b>automatically renewed for one (1) additional year</b> at the end of each term. The new term is effective from <b>${newStartDate}</b>, unless either party provides written notice of its desire not to automatically renew the agreement at least thirty (30) days prior to the end of the current term.</p>
        <p>Except the agreement term stated in this notice, all of the terms and conditions of the agreement remain unchanged and in full force and effect. <b>Should you require any additional information or would like to discuss this matter further, please reply to this email and we will contact you shortly. If we do not receive any response within 7 days, the agreement will proceed with automatic renewal as stated above.</b></p>
        ${ddaSectionEN}
        <p>We sincerely appreciate your continued trust in inline. We look forward to continuing our partnership and supporting your business.</p>

        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
        
        <div style="text-align: center; margin: 20px 0;">
          <h2 style="font-weight: bold; font-size: 1.25em;">inline 軟體自動續約通知</h2>
        </div>
        <p>inline 與 <b>${restaurantName}</b> 就提供軟體系統簽定之系統訂閱合約書（下稱「合約」），現行訂閱期於 ${oldStartDate} 開始，並將於 ${currentEndDate} 屆滿。</p>
        <p><b>[${billingCycle === 'monthly' ? '月繳' : '年繳'} ${planCH}方案]</b><br>
        訂閱詳情：<br>
        HK$${rsvPrice} 包含 ${rsvSets && rsvSets.toString().toLowerCase() === "unlimited" ? "無限" : rsvSets} 組訂候位${overageTextCH}</p>

        ${addonSectionCH}

        <p>按合約之條款，訂閱期屆滿後<b>將自動延長一年</b>，其後亦同。新訂閱期將由 <b>${newStartDate}</b> 起生效，為期一年，除任一方於訂閱期屆滿之 30 日前以書面通知終止合約。</p>
        <p>除合約訂閱期限自動延長一年外，其餘本通知未盡之事項，以原合約條款之規定辦理。<b>若有任何疑問，或希望進一步討論此事宜，請回覆此郵件，我們將儘快與您聯繫。如我們未在 7 日內接獲您的通知，合約將依上述約定自動續約。</b></p>
        ${ddaSectionCH}
        <p>感謝您一直以來對 inline 的信任。我們期待與您繼續合作，並持續為您的業務發展提供支持。</p>
        <br>
        <p>Best Regards,</p>
        <p><strong>${csmInfo.fullname}</strong><br>
        <b>${csmInfo.title}</b><br>
        <b>Mobile: +${csmInfo.mobile}</b><br>
        <b>Email: ${csmInfo.email}</b><br>
        <b>46/F Lee Garden One, 33 Hysan Avenue, Causeway Bay, Hong Kong</b></p>
      </div>
    `;

    GmailApp.sendEmail(recipientEmail, subject, "", {
      htmlBody: htmlBody,
      name: "inline Customer Success",
      //cc: csmInfo.email
    });

    // 寄送完畢後，回填狀態、日期
    mainSheet.getRange(i + 1, 26).setValue(new Date());
    mainSheet.getRange(i + 1, 25).setValue("sent");
  }
}

function formatDate(date) {
  if (!(date instanceof Date)) return date;
  return Utilities.formatDate(date, "GMT+8", "yyyy-MM-dd");
}