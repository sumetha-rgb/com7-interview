const FETCH_SEND_EMAIL_URL = "https://api-send-mail-com7.vercel.app/api/mail/send-mail"

/**
 * ส่งอีเมล: พยายามส่งผ่าน Vercel API (UrlFetchApp) เป็นหลักเหมือนเดิม
 * ถ้าล้มเหลว (เช่น UrlFetchApp quota หมดรายวัน, network error, หรือ
 * Vercel ตอบ error) จะ fallback ไปส่งผ่าน Gmail (MailApp.sendEmail)
 * แทนโดยอัตโนมัติ เพราะ MailApp ใช้ quota คนละก้อนกับ UrlFetchApp
 * (ผูกกับบัญชี Google ที่รันสคริปต์ ไม่ใช่ project) จึงไม่โดนผลกระทบ
 * พร้อมกันตอน UrlFetchApp โควต้าเต็ม
 *
 * Return shape เหมือนเดิมทุกอย่าง { success, message, statusCode } เพิ่ม
 * แค่ field "channel" (ไม่บังคับใช้) บอกว่าส่งผ่านช่องทางไหนสำเร็จ
 */
function sendEmail(data) {
  if (!data) {
    return {
      success: false,
      message: "No data provided",
    };
  }

  // 1) ช่องทางหลัก: Vercel API ผ่าน UrlFetchApp (พฤติกรรมเดิมทุกอย่าง)
  try {
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(data),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(FETCH_SEND_EMAIL_URL, options);

    const statusCode = response.getResponseCode();
    const bodyText = response.getContentText();

    let body;
    try {
      body = JSON.parse(bodyText);
    } catch (e) {
      body = { message: bodyText };
    }

    if (statusCode >= 200 && statusCode < 300) {
      return {
        success: true,
        message: body.message || "ส่งอีเมลสำเร็จ",
        statusCode,
        channel: "vercel",
      };
    }

    // Vercel ตอบกลับมาแต่ไม่ใช่ 2xx (เช่น 500 จากฝั่งเขาเอง) → ลอง fallback
    console.warn(
      `sendEmail: Vercel API responded ${statusCode}: ${bodyText}. Falling back to Gmail.`,
    );
  } catch (e) {
    // เข้ากรณีนี้บ่อยที่สุดตอน UrlFetchApp quota หมดรายวัน
    // ("Service invoked too many times for one day: urlfetch")
    console.warn(
      `sendEmail: UrlFetchApp failed (${e.message}). Falling back to Gmail.`,
    );
  }

  // 2) ช่องทางสำรอง: ส่งตรงผ่าน Gmail (ไม่ใช้ UrlFetchApp เลย)
  return sendEmailViaGmailFallback(data);
}

/**
 * ส่งอีเมลสำรองผ่าน MailApp (Gmail ของบัญชีที่รันสคริปต์)
 * หมายเหตุ: เทมเพลตที่ใช้ตรงนี้เป็นแบบเรียบง่าย ไม่ได้ตรงดีไซน์กับเทมเพลต
 * จริงฝั่ง Vercel เป๊ะ 100% (ไม่มีสิทธิ์เข้าถึงโค้ดฝั่งนั้น) แต่ครบข้อมูล
 * สำคัญที่ผู้สมัครต้องรู้ ใช้เฉพาะตอนช่องทางหลักใช้งานไม่ได้เท่านั้น
 */
function sendEmailViaGmailFallback(data) {
  try {
    if (!data.to) {
      return {
        success: false,
        message: "ไม่พบอีเมลผู้รับ (fallback)",
        channel: "gmail-fallback",
      };
    }

    // เช็ค quota ของ Gmail ก่อน กัน throw ซ้ำถ้าโควต้าฝั่งนี้หมดไปด้วย
    // (Gmail ปกติ 100 ฉบับ/วัน, Google Workspace 1,500 ฉบับ/วัน)
    let remainingQuota = null;
    try {
      remainingQuota = MailApp.getRemainingDailyQuota();
    } catch (e) {
      console.warn("sendEmail fallback: cannot read MailApp quota: " + e.message);
    }

    if (remainingQuota !== null && remainingQuota <= 0) {
      console.error("sendEmail fallback: MailApp daily quota exhausted too.");
      return {
        success: false,
        message: "โควต้าอีเมลสำรอง (Gmail) หมดสำหรับวันนี้เช่นกัน",
        channel: "gmail-fallback",
      };
    }

    const { subject, htmlBody } = buildFallbackEmailContent(data);

    MailApp.sendEmail({
      to: data.to,
      subject: subject,
      htmlBody: htmlBody,
      name: "COM7 Recruitment",
    });

    console.log(`sendEmail fallback: sent via Gmail to ${data.to}`);

    return {
      success: true,
      message: "ส่งอีเมลสำเร็จผ่านช่องทางสำรอง (Gmail)",
      statusCode: 200,
      channel: "gmail-fallback",
    };
  } catch (e) {
    console.error("sendEmail fallback (Gmail) failed: ", e);
    return {
      success: false,
      message: "ส่งอีเมลไม่สำเร็จทั้งช่องทางหลักและช่องทางสำรอง: " + e.message,
      channel: "gmail-fallback",
    };
  }
}

/**
 * สร้างหัวข้อ/เนื้อหาอีเมลสำหรับกรณี fallback เท่านั้น
 * ครอบคลุม pattern payload ที่ใช้เรียก sendEmail() อยู่จริงในระบบ:
 * - type === "CALL_LINK": แจ้งลิงก์สัมภาษณ์ (ใช้เทมเพลตจริงเดียวกับ Vercel)
 * - ไม่มี type และไม่มี eventName: แจ้งเปลี่ยนสถานะ/ยกเลิกการสมัคร
 * - กรณีอื่นๆ: ยืนยัน/แจ้งเตือนนัดหมายสัมภาษณ์ทั่วไป (นัดใหม่, เลื่อนนัด)
 */
function buildFallbackEmailContent(data) {
  const fullname = data.fullname || "ผู้สมัคร";
  const eventName = data.eventName || "COM7";

  if (data.type === "CALL_LINK") {
    // เทมเพลตนี้เป็นของจริงที่ใช้ฝั่ง Vercel (คัดลอกมาทั้งดุ้น ไม่ใช่เวอร์ชัน
    // เรียบง่ายแบบเดิม) เนื้อหาเป็น static ทั้งหมด ไม่มีการแทนที่ชื่อผู้รับ
    // หรือรายละเอียด event เพราะต้นฉบับออกแบบมาแบบนี้อยู่แล้ว
    // ลิงก์อ้างอิงจาก WEBSITE_URL (ประกาศไว้ที่ services/canlendar.service.js
    // และใช้ร่วมกันทั้งโปรเจกต์ผ่าน global scope ของ Apps Script) แทนการ
    // hardcode URL ซ้ำอีกที่ — หมายเหตุ: URL ในเทมเพลตต้นฉบับที่ได้รับมา
    // (...AKfycbyH.../exec) ไม่ตรงกับ WEBSITE_URL ที่ใช้อยู่จริงในระบบตอนนี้
    // (...AKfycbzy.../exec) จึงยึดตาม WEBSITE_URL ปัจจุบันเพื่อให้ลิงก์ใช้งาน
    // ได้จริง ถ้าต้องการ URL อื่นเฉพาะอีเมลนี้ แจ้งได้เพื่อแก้ไข
    const subject = "COM7 ประกาศแก้ไขลิงก์สัมภาษณ์งาน";
    const websiteUrl =
      typeof WEBSITE_URL !== "undefined" && WEBSITE_URL
        ? WEBSITE_URL
        : data.interviewLink || "#";
    const htmlBody = `
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      body {
        font-family: "Sarabun", Arial, sans-serif;
        color: #333333;
        line-height: 1.6;
        margin: 0;
        padding: 0;
      }
      .container {
        max-width: 1200px;
        margin: 0 auto;
        overflow: hidden;
        background-color: #ffffff;
      }
      .img-fluid {
        width: 100%;
        height: auto;
        display: block;
      }
      .content-container {
        padding: 2rem;
      }
      .section-box {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 20px;
        border-left: 5px solid #7cbd1e;
      }
      .info-list p {
        margin: 5px 0;
      }
      .btn-link {
        color: #007bff;
        word-break: break-all;
        text-decoration: none;
      }
      .steps {
        margin: 0;
        padding: 0;
      }
      .footer-text {
        margin-top: 20px;
        font-size: 0.85em;
        color: #666;
        border-top: 1px solid #eee;
        padding-top: 15px;
      }
      h3 {
        color: #000;
        margin-top: 0;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="content-container">
        <h1 style="color: #0ea60e">COM7 ประกาศแก้ไขลิงก์สัมภาษณ์งาน</h1>
        <p style="margin: 0; text-indent: 25px">
          บริษัท คอมเซเว่น จำกัด (มหาชน)
          ขออภัยในความไม่สะดวกในการเข้าร่วมการสัมภาษณ์งาน
        </p>
        <p>ท่านสามารถเข้าสัมภาษณ์ได้ที่ลิงก์ด้านล่างนี้</p>

        <p style="margin-bottom: 0"><strong>💻 ลิงค์สัมภาษณ์งาน</strong></p>
        <p style="margin: 0; margin-bottom: 10px">
          <a
            href="${websiteUrl}"
            class="btn-link"
            >${websiteUrl}</a
          >
        </p>

        <p>รายละเอียดดังนี้</p>
        <ul class="steps" style="list-style-type: none">
          <li>1️⃣ ล็อกอินเข้าสู่ระบบให้เรียบร้อย</li>
          <li>2️⃣ เข้าเมนู <strong>ผลการสัมภาษณ์</strong></li>
          <li>3️⃣ ลิงก์สัมภาษณ์จะอยู่ในโปรไฟล์ของตำแหน่งงานที่ท่านสมัครไว้</li>
        </ul>

        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0" />

        <div class="section-box" style="border-left-color: #ffc107">
          <p><strong>👔 การแต่งกายสุภาพเรียบร้อย</strong></p>
        </div>

        <p><strong>🔎 ข้อแนะนำในการเตรียมตัวสัมภาษณ์</strong></p>
        <ul class="steps">
          <li>
            ติดตั้งแอปพลิเคชัน Microsoft Teams ในการสัมภาษณ์ออนไลน์
            และทดสอบว่าทุกอย่างพร้อมใช้งาน
          </li>
          <li>ตรวจสอบสัญญาณอินเทอร์เน็ตให้พร้อมใช้งาน</li>
          <li>เลือกสถานที่ในการสัมภาษณ์ให้เหมาะสม ไม่มีเสียงรบกวน</li>
          <li>เตรียมทุกอย่างให้พร้อมก่อนเวลานัดสัมภาษณ์ อย่างน้อย 10 นาที</li>
        </ul>

        <p style="padding-top: 10px">
          ทั้งนี้ บริษัทฯ ขออนุญาตเก็บข้อมูลใบสมัครของท่านไว้ในฐานข้อมูล
          เพื่อพิจารณาสำหรับโอกาสในการร่วมงานในอนาคต หากมีตำแหน่งที่เหมาะสม
          บริษัทฯ จะติดต่อท่านอีกครั้ง
        </p>
        <span
          >บริษัทฯ ขอขอบพระคุณในความสนใจ
          และหวังเป็นอย่างยิ่งว่าจะมีโอกาสได้ร่วมงานกับท่านในอนาคต</span
        >

        <div class="footer-text">
          <p style="color: #0ea60e; font-weight: 600">
            Com7 (Public) Company Limited <br />
            549/1 Sanphawut Rd., Bangna Tai, Bangna, Bangkok 10260<br />
            <span style="color: #000"
              >Tel: 02-017-7777 ต่อ 7209, 7208, 7210, 7217, 7230</span
            >
          </p>
        </div>
      </div>
    </div>
  </body>
</html>`;
    return { subject, htmlBody };
  }

  if (!data.type && !data.eventName) {
    const subject = "แจ้งผลการสมัครสัมภาษณ์";
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #111;">
        <p>เรียนคุณ ${fullname}</p>
        <p>การสมัครสัมภาษณ์ของคุณมีการเปลี่ยนแปลงสถานะ กรุณาติดต่อฝ่ายทรัพยากรบุคคลหากมีข้อสงสัย</p>
        <p style="color:#6b7280; font-size:12px; margin-top:24px;">อีเมลนี้ส่งผ่านช่องทางสำรองของระบบ</p>
      </div>`;
    return { subject, htmlBody };
  }

  const subject = `นัดหมายสัมภาษณ์ - ${eventName}`;
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px; color: #111;">
      <p>เรียนคุณ ${fullname}</p>
      <p>งาน: <b>${eventName}</b>${data.buName ? ` (${data.buName})` : ""}</p>
      <p>วันที่: ${data.date || "-"}</p>
      <p>เวลา: ${data.timePeriod || "-"}</p>
      <p>สถานที่: ${data.location || "-"}</p>
      <p style="color:#6b7280; font-size:12px; margin-top:24px;">อีเมลนี้ส่งผ่านช่องทางสำรองของระบบ</p>
    </div>`;
  return { subject, htmlBody };
}