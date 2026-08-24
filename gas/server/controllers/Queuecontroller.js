// ==========================================
// QUEUE SYSTEM CONSTANTS & CONFIGURATIONS
// ==========================================
const PARTICIPATION_STATUS = ["เข้าร่วม", "Walk-In", "ไม่เข้าร่วม"];
const QUEUE_LOG_SS_ID = "1egEcYtBeBfxiu82lqW4vEmuFK4uTCL-Vdei-ZSUDF8A";
const QUEUE_LOG_SHEET_NAME = "Queue";
const QUEUE_LOG_HEADERS = ["No", "Event", "ชื่อ", "เวลาที่เข้าร่วมงาน", "ช่วงเวลาสัมภาษณ์", "สถานะงาน", "ผู้อนุมัติเช็คชื่อ", "สถานะเช็คชื่อ", "Interview_Id", "ส่งอีเมลคิวแล้ว", "Queue_Admin_Status", " "];

// ==========================================
// CORE QUEUE FUNCTIONS
// ==========================================

function getQueueLogSheet() {
  const ss = SpreadsheetApp.openById(QUEUE_LOG_SS_ID);
  let sheet = ss.getSheetByName(QUEUE_LOG_SHEET_NAME) || ss.getSheets()[0];

  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_LOG_SHEET_NAME);
  }

  const lastColumn = Math.max(sheet.getLastColumn(), QUEUE_LOG_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const hasHeaders = QUEUE_LOG_HEADERS.every(
    (header, index) => headers[index] === header,
  );

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, QUEUE_LOG_HEADERS.length).setValues([QUEUE_LOG_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function extractDateKey(timestampStr) {
  if (!timestampStr) return "";

  if (Object.prototype.toString.call(timestampStr) === "[object Date]") {
    if (isNaN(timestampStr.getTime())) return "";
    const tz = Session.getScriptTimeZone();
    return Utilities.formatDate(timestampStr, tz, "yyyy-MM-dd");
  }

  const s = String(timestampStr).trim();

  const thaiMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (thaiMatch) {
    const day = Number(thaiMatch[1]);
    const month = Number(thaiMatch[2]);
    let year = Number(thaiMatch[3]);
    if (year > 2400) year -= 543;
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    let year = Number(isoMatch[1]);
    if (year > 2400) year -= 543;
    return `${year}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    const tz = Session.getScriptTimeZone();
    return Utilities.formatDate(parsed, tz, "yyyy-MM-dd");
  }

  const now = new Date();
  return now.getFullYear() + "-" +
    String(now.getMonth() + 1).padStart(2, "0") + "-" +
    String(now.getDate()).padStart(2, "0");
}

function normalizeNameKey(value) {
  return value ? String(value).trim().replace(/\s+/g, " ").toLowerCase() : "";
}

function formatJoinedAtForDisplay(value) {
  if (!value) return "-";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (isNaN(value.getTime())) return "-";
    const tz = Session.getScriptTimeZone();
    let gregorianYear = Number(Utilities.formatDate(value, tz, "yyyy"));

    // กันข้อมูลเก่าที่เคยถูกบันทึกด้วย toLocaleString("th-TH") มาก่อน
    // ซึ่งทำให้ Sheets auto-parse ปี พ.ศ. เป็นปีของ Date object ไปแล้ว (เช่น 2569)
    // ถ้าไม่กันไว้ การ +543 อีกครั้งด้านล่างจะได้ปีเพี้ยนเป็น 3112 แบบเดิม
    if (gregorianYear > 2400) {
      gregorianYear -= 543;
    }

    const buddhistYear = gregorianYear + 543;
    const datePart = Utilities.formatDate(value, tz, "dd/MM/") + buddhistYear;
    const timePart = Utilities.formatDate(value, tz, "HH:mm:ss");
    return `${datePart} ${timePart}`;
  }

  return String(value).trim() || "-";
}

function appendQueueLog(eventName, candidateName, joinedAt, interviewPeriod, interviewStatus, approvedBy, participationStatus, interviewId, employeeCategory) {
  const sheet = getQueueLogSheet();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);

  const normalizedEventValue = normalizeNameKey(eventName);
  const normalizedCandidateValue = normalizeNameKey(candidateName);
  const rawEventValue = String(eventName || "").trim();
  const rawCandidateValue = String(candidateName || "").trim();
  const interviewIdValue = String(interviewId || "").trim();
  const employeeCategoryValue = String(employeeCategory || "").trim();
  const todayKey = extractDateKey(joinedAt || new Date().toLocaleDateString("th-TH"));

  const existingRowIndex = rows.findIndex((row) => {
    const rowDateMatches = extractDateKey(row[3]) === todayKey;
    if (!rowDateMatches) return false;

    if (interviewIdValue && row[8]) {
      return String(row[8]).trim() === interviewIdValue;
    }
    return (
      normalizeNameKey(row[1]) === normalizedEventValue &&
      normalizeNameKey(row[2]) === normalizedCandidateValue
    );
  });

  if (existingRowIndex !== -1) {
    const rowNum = existingRowIndex + 2;
    sheet.getRange(rowNum, 6).setValue(interviewStatus || "");
    sheet.getRange(rowNum, 7).setValue(approvedBy || "");
    sheet.getRange(rowNum, 8).setValue(participationStatus || "");
    sheet.getRange(rowNum, 9).setValue(interviewIdValue || rows[existingRowIndex][8] || "");
    if (employeeCategoryValue && !String(rows[existingRowIndex][11] || "").trim()) {
      sheet.getRange(rowNum, 12).setValue(employeeCategoryValue);
    }

    return {
      queueNo: rows[existingRowIndex][0],
      joinedAt: rows[existingRowIndex][3],
      rowNum,
      employeeCategory: String(rows[existingRowIndex][11] || employeeCategoryValue || "").trim(),
    };
  }

  const todayQueues = rows.filter((row) => {
    return (
      normalizeNameKey(row[1]) === normalizedEventValue &&
      extractDateKey(row[3]) === todayKey &&
      normalizeNameKey(String(row[11] || "")) === normalizeNameKey(employeeCategoryValue)
    );
  });

  const queueNo = todayQueues.length + 1;

  sheet.appendRow([
    queueNo,
    rawEventValue,
    rawCandidateValue,
    joinedAt,
    interviewPeriod || "",
    interviewStatus || "",
    approvedBy || "",
    participationStatus || "",
    interviewIdValue,
    "",
    "",
    employeeCategoryValue,
  ]);

  return { queueNo, joinedAt, rowNum: sheet.getLastRow(), employeeCategory: employeeCategoryValue };
}

function updateQueueAdminStatus(updateData) {
  try {
    const { interviewId, status, eventName, candidateName } = updateData || {};
    const validStatuses = ["ยังไม่สัมภาษณ์", "สัมภาษณ์แล้ว", "ไม่เข้าร่วม"];
    if (!validStatuses.includes(status)) {
      return { success: false, message: "สถานะไม่ถูกต้อง" };
    }

    const sheet = getQueueLogSheet();
    const values = sheet.getDataRange().getValues();
    const rows = values.slice(1);

        let rowIndex = -1;
    if (interviewId) {
      // ค้นหาแถวล่างสุด (ล่าสุด) ที่ตรงกับ Interview_Id แทนแถวแรกสุด
      // เพื่อให้สอดคล้องกับ getQueueCandidates ที่ยึดแถวล่างสุดเป็นค่าปัจจุบันเสมอ
      // (ถ้ามีแถวซ้ำจาก appendQueueLog เช่น เช็คชื่อคนละวันแต่ Interview_Id เดิม
      // จะได้ไม่เขียนสถานะไปคนละแถวกับที่หน้าเว็บอ่านมาแสดง)
      const targetInterviewId = String(interviewId).trim();
      for (let i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][8] || "").trim() === targetInterviewId) {
          rowIndex = i;
          break;
        }
  }
}

    if (rowIndex === -1 && eventName && candidateName) {
      const eventValue = normalizeNameKey(eventName);
      const candidateValue = normalizeNameKey(candidateName);
      for (let i = rows.length - 1; i >= 0; i--) {
        if (
          normalizeNameKey(rows[i][1]) === eventValue &&
          normalizeNameKey(rows[i][2]) === candidateValue
        ) {
          rowIndex = i;
          break;
        }
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบข้อมูลคิวสำหรับอัปเดต" };
    }

    const rowNum = rowIndex + 2;
    sheet.getRange(rowNum, 11).setValue(status);

    return { success: true, message: "เปลี่ยนสถานะสัมภาษณ์เรียบร้อยแล้ว" };
  } catch (error) {
    console.error("Error in updateQueueAdminStatus:", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด" };
  }
}

function setQueueLogEmailStatus(rowNum, sent) {
  if (!rowNum) return;
  try {
    const sheet = getQueueLogSheet();
    sheet.getRange(rowNum, 10).setValue(sent ? "TRUE" : "FALSE");
  } catch (e) {
    console.error("setQueueLogEmailStatus error:", e);
  }
}

function findQueueLog(eventName, candidateName, interviewId) {
  const sheet = getQueueLogSheet();
  const values = sheet.getDataRange().getValues();
  const eventValue = normalizeNameKey(eventName);
  const candidateValue = normalizeNameKey(candidateName);
  const interviewIdValue = String(interviewId || "").trim();
  const todayKey2 = extractDateKey(new Date().toLocaleString("th-TH"));

  if (interviewIdValue) {
    const rowById = values.slice(1).find((item) => String(item[8] || "").trim() === interviewIdValue);
    if (rowById) {
      return {
        queueNo: rowById[0],
        eventName: rowById[1],
        candidateName: rowById[2],
        joinedAt: formatJoinedAtForDisplay(rowById[3]),
      };
    }
  }

  const row = values.slice(1).find((item) => {
    const rowDate = extractDateKey(item[3]);
    return (
      normalizeNameKey(item[1]) === eventValue &&
      normalizeNameKey(item[2]) === candidateValue &&
      (rowDate === todayKey2 || !rowDate)
    );
  });

  if (!row) return null;

  return {
    queueNo: row[0],
    eventName: row[1],
    candidateName: row[2],
    joinedAt: formatJoinedAtForDisplay(row[3]),
  };
}

function getQueueDashboard(eventName) {
  try {
    if (!eventName) return { success: false, message: "กรุณาระบุชื่ออีเวนต์" };

    const eventSheet = getEventSheet();
    const eventData = getRowsFromSheet(eventSheet);
    const event = eventData.find((e) => e.Event_Name === eventName);

    if (!event || !event.Sheet_ID) return { success: false, message: "ไม่พบข้อมูลอีเวนต์: " + eventName };

    const ss = SpreadsheetApp.openById(event.Sheet_ID);
    if (!ss) return { success: false, message: "ไม่สามารถเปิด Spreadsheet ได้" };

    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    const transactionSheet = ss.getSheetByName(EVENT_SHEETS.TRANSACTION);

    if (!interviewSheet) return { success: false, message: "ไม่พบชีท Interviews" };

    const interviews = getRowsFromSheet(interviewSheet);
    const transactions = transactionSheet ? getRowsFromSheet(transactionSheet) : [];

    const queueSheet = getQueueLogSheet();
    const queueValues = queueSheet.getDataRange().getValues();
    const queueRows = queueValues.slice(1);

    const participationTransMap = {};
    transactions.forEach((tx) => {
      const txStatus = String(tx[5] || tx.Status || "").trim();
      if (!PARTICIPATION_STATUS.includes(txStatus)) return;

      const txInterviewId = String(tx[1] || tx.Interview_Id || "").trim();
      if (!txInterviewId) return;

      if (!participationTransMap[txInterviewId]) {
        participationTransMap[txInterviewId] = tx;
      } else {
        const existingTime = String(participationTransMap[txInterviewId][0] || "");
        const newTime = String(tx[0] || tx.TimeStamp || "");
        if (newTime > existingTime) {
          participationTransMap[txInterviewId] = tx;
        }
      }
    });

    const result = interviews.map((interview) => {
      const interviewId = interview.Interview_Id || "";
      const fullname = interview.Fullname || "";
      const participationStatus = interview.Participation_Status || "ไม่เข้าร่วม";
      const interviewStatus = interview.Status || "รอการพิจารณา";
      const interviewPeriod = interview.Interview_Period || "";

      const queueRow = queueRows.find((row) => {
        return (
          String(row[1] || "").trim() === String(eventName).trim() &&
          String(row[2] || "").trim() === String(fullname).trim()
        );
      });

      const queueNo = queueRow ? queueRow[0] : null;
      const joinedAt = queueRow ? formatJoinedAtForDisplay(queueRow[3]) : null;

      const txRow = participationTransMap[interviewId];
      let approvedBy = "-";
      if (txRow) {
        approvedBy = String(txRow[3] || txRow.FullName || txRow[2] || "").trim() || "-";
      }

      return {
        queueNo: queueNo || "-",
        fullname,
        interviewStatus,
        joinedAt: joinedAt || "-",
        interviewPeriod,
        approvedBy,
        participationStatus,
      };
    });

    result.sort((a, b) => {
      const aChecked = a.participationStatus === "เข้าร่วม" || a.participationStatus === "Walk-In";
      const bChecked = b.participationStatus === "เข้าร่วม" || b.participationStatus === "Walk-In";
      if (aChecked && !bChecked) return -1;
      if (!aChecked && bChecked) return 1;
      const aNo = typeof a.queueNo === "number" ? a.queueNo : parseInt(a.queueNo) || 9999;
      const bNo = typeof b.queueNo === "number" ? b.queueNo : parseInt(b.queueNo) || 9999;
      return aNo - bNo;
    });

    return { success: true, data: result };
  } catch (error) {
    console.error("Error in getQueueDashboard:", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด: " + error.message };
  }
}

function getQueueCandidates(query, userRole) {
  try {
    const { page = 1, limit = 20, statusFilter, employeeCategory, eventName } = query || {};
    const pageNum = Math.max(1, parseInt(page));
    const limitNum = Math.max(1, parseInt(limit));

    const result = getAllInterviews({ ...query, page: 1, limit: 999999 }, userRole);

    if (!result.success || !result.data || result.data.length === 0) return result;

    let interviews = result.data;

    // ── แสดงเฉพาะคิวของ "วันนี้" เท่านั้น (กันไม่ให้ "ไม่เข้าร่วม" ของ Event/วันเก่าปนมากับคิววันนี้) ──
    // ถ้าผู้ใช้เจาะจงเลือก Event จาก dropdown เอง จะไม่จำกัดแค่วันนี้ เพื่อให้ยังดูคิวย้อนหลังของ Event นั้นได้
    if (!eventName) {
      const todayDateKey = extractDateKey(new Date());
      interviews = interviews.filter((interview) => {
        // บั๊กเดิม: เช็คแค่วันที่ของสลอตที่นัดหมายไว้ (slotDetail.Date) เพียง
        // อย่างเดียว ทำให้ผู้สมัครที่ "เช็คชื่อแล้ว" จริงในวันนี้ (Participation_Status
        // เป็น "เข้าร่วม"/"Walk-In") แต่สลอตที่ผูกไว้ไม่ตรงกับวันนี้พอดี — หรือไม่มี
        // slotDetail เลย (Slot_Id ไม่ตรงกับชีท Slots เช่นกรณี Walk-In ที่ไม่ได้จอง
        // สลอตล่วงหน้า) — ถูกกรองทิ้งไปทั้งหมด ทั้งที่เพิ่งเช็คชื่อเข้างานไปจริงๆ
        // แก้โดยยึดวันที่เช็คชื่อจริง (จาก transaction ล่าสุดที่เปลี่ยนสถานะเป็น
        // เข้าร่วม/Walk-In) เป็นหลักสำหรับคนที่เช็คชื่อแล้ว แล้วค่อย fallback ไปดู
        // วันที่ของสลอตเหมือนเดิมสำหรับคนที่ยังไม่เช็คชื่อ (เพื่อให้ยังเห็นคิวที่นัดไว้
        // วันนี้แต่ยังไม่มาเช็คชื่อ)
        const participation = String(interview.Participation_Status || "").trim();
        const isCheckedIn = participation === "เข้าร่วม" || participation === "Walk-In";

        if (isCheckedIn) {
          let checkInDateKey = null;
          if (Array.isArray(interview.transactionDetail)) {
            const checkInTx = interview.transactionDetail
              .filter((tx) => tx.Status_Changes === "เข้าร่วม" || tx.Status_Changes === "Walk-In")
              .sort((a, b) => parseThaiDateTime(b.TimeStamp).getTime() - parseThaiDateTime(a.TimeStamp).getTime())[0];
            if (checkInTx) {
              checkInDateKey = extractDateKey(checkInTx.TimeStamp);
            }
          }
          if (checkInDateKey) {
            return checkInDateKey === todayDateKey;
          }
          // ไม่มี transaction ที่ระบุวันที่เช็คชื่อได้ชัดเจน ให้ตกไปเช็คจากสลอตแทน (fallback ด้านล่าง)
        }

        const slot = interview.slotDetail;
        if (!slot || !slot.Date) return false;
        try {
          const slotDateKey = Utilities.formatDate(
            normalizeDate(slot.Date),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          );
          return slotDateKey === todayDateKey;
        } catch (e) {
          return false;
        }
      });
    }

    if (employeeCategory && employeeCategory !== "all") {
      interviews = interviews.filter((interview) => {
        const positionText = String(
          interview.Position_Applied || interview.Employee_Category || interview.Position_Type || ""
        ).trim();
        return getEmployeeCategoryFromJobType(positionText) === employeeCategory;
      });
    }

    // หมายเหตุ: เคยลองแคชค่าดิบของชีทนี้ด้วย CacheService แต่พบว่า Date
    // object ในคอลัมน์ joinedAt (row[3]) จะถูกแปลงเป็น string ตอน
    // JSON.stringify ทำให้ extractDateKey()/formatJoinedAtForDisplay()
    // ด้านล่าง (ที่เช็คชนิดข้อมูลด้วย Object.prototype.toString) ตรวจไม่
    // เจอว่าเป็น Date แล้ว parse ผิดรูปแบบ จึงตัดสินใจไม่แคชจุดนี้ (หน้านี้
    // มี Admin ใช้งานไม่ถึง 10 คน ความเสี่ยงต่อ logic ไม่คุ้มกับที่จะได้)
    const queueSheet = getQueueLogSheet();
    const queueValues = queueSheet.getDataRange().getValues();
    const todayKey = extractDateKey(new Date().toLocaleString("th-TH"));

    const queueMap = {};       
    const queueMapByName = {}; 

    for (let i = 1; i < queueValues.length; i++) {
      const row = queueValues[i];
      const rowDateKey = extractDateKey(row[3]);

      const eventName = String(row[1] || "").trim();
      const candidateName = String(row[2] || "").trim();
      const normalizedEventName = normalizeNameKey(eventName);
      const normalizedCandidateName = normalizeNameKey(candidateName);
      const queueNo = row[0];
      const approvedBy = row[6];
      const rowInterviewId = String(row[8] || "").trim();
      const emailSentRaw = String(row[9] || "").trim().toUpperCase();
      const entry = {
        queueNo,
        joinedAt: formatJoinedAtForDisplay(row[3]),
        approvedBy,
        emailSent: emailSentRaw === "TRUE",
        hasEmailRecord: emailSentRaw === "TRUE" || emailSentRaw === "FALSE",
        queueAdminStatus: String(row[10] || "").trim(),
        queueEmployeeCategory: String(row[11] || "").trim(),
      };

      if (rowInterviewId) queueMap[rowInterviewId] = entry;
      if (rowDateKey !== todayKey) continue;

      const nameKey = normalizedEventName + "||" + normalizedCandidateName;
      if (nameKey !== "||") queueMapByName[nameKey] = entry;
    }

    const now = new Date();

    const enriched = interviews.map((interview) => {
      const eventName = (interview.eventDetail && interview.eventDetail.Event_Name) || "";
      const fullname = interview.Fullname || "";
      const nameKey = normalizeNameKey(eventName) + "||" + normalizeNameKey(fullname);
      const queueEntry = queueMap[String(interview.Interview_Id || "").trim()] || queueMapByName[nameKey];
      const queueAdminStatus = queueEntry && queueEntry.queueAdminStatus ? queueEntry.queueAdminStatus : undefined;

      const participation = String(interview.Participation_Status || "").trim();
      const isCheckedIn = participation === "เข้าร่วม" || participation === "Walk-In";

      let slotEndDateTime = null;
      try {
        const slot = interview.slotDetail;
        if (slot && slot.Date && slot.End_Time) {
          const slotDate = Utilities.formatDate(normalizeDate(slot.Date), Session.getScriptTimeZone(), "yyyy-MM-dd");
          slotEndDateTime = combineDateTime(slotDate, formatTimeTh(new Date(slot.End_Time)));
        }
      } catch (e) {
        slotEndDateTime = null;
      }

      const isOverdue = !!(slotEndDateTime && now.getTime() > slotEndDateTime.getTime());
      const adminStatus = isCheckedIn ? "ยังไม่สัมภาษณ์" : (isOverdue ? "ไม่เข้าร่วม" : "รอเช็คชื่อ");
      let displayStatus = "รอเช็คชื่อ"; 
      if (isCheckedIn) {
        displayStatus = "เช็คชื่อแล้ว"; 
      } else if (isOverdue) {
        displayStatus = "ไม่เข้าร่วม"; 
      }

      let emailStatus = "no_queue";
      if (queueEntry) {
        emailStatus = queueEntry.hasEmailRecord ? (queueEntry.emailSent ? "sent" : "failed") : "unknown";
      }

      let fallbackCheckInTime = null;
      let fallbackApprovedBy = null;
      if (isCheckedIn && !queueEntry && Array.isArray(interview.transactionDetail)) {
        const checkInTx = interview.transactionDetail
          .filter((tx) => tx.Status_Changes === "เข้าร่วม" || tx.Status_Changes === "Walk-In")
          .sort((a, b) => parseThaiDateTime(b.TimeStamp).getTime() - parseThaiDateTime(a.TimeStamp).getTime())[0];

        if (checkInTx) {
          fallbackCheckInTime = formatJoinedAtForDisplay(checkInTx.TimeStamp);
          fallbackApprovedBy = checkInTx.Name || null;
          emailStatus = "sent";
        }
      }

      return {
        ...interview,
        queueNo: queueEntry ? queueEntry.queueNo : "-",
        joinedAt: queueEntry ? queueEntry.joinedAt : (fallbackCheckInTime || "-"),
        approvedBy: queueEntry && queueEntry.approvedBy ? queueEntry.approvedBy : (fallbackApprovedBy || "-"),
        Queue_Display_Status: displayStatus,
        Queue_Admin_Status: queueAdminStatus || adminStatus,
        Queue_Employee_Category: (queueEntry && queueEntry.queueEmployeeCategory) || interview.Employee_Category || interview.Position_Type || "",
        emailStatus,
      };
    });

    enriched.sort((a, b) => {
      const aChecked = a.Queue_Display_Status === "เช็คชื่อแล้ว";
      const bChecked = b.Queue_Display_Status === "เช็คชื่อแล้ว";
      if (aChecked && !bChecked) return -1;
      if (!aChecked && bChecked) return 1;

      if (aChecked && bChecked) {
        const aNo = parseInt(a.queueNo) || 9999;
        const bNo = parseInt(b.queueNo) || 9999;
        return aNo - bNo;
      }
      return 0;
    });

    const counts = {
  total: enriched.length,
  checkedIn: enriched.filter((c) => c.Queue_Display_Status === "เช็คชื่อแล้ว").length,
  waiting: enriched.filter((c) => c.Queue_Admin_Status === "ยังไม่สัมภาษณ์").length,
  interviewed: enriched.filter((c) => c.Queue_Admin_Status === "สัมภาษณ์แล้ว").length,
  absent: enriched.filter((c) => c.Queue_Display_Status === "ไม่เข้าร่วม").length,
};

    const filteredByStatus = statusFilter && statusFilter !== "all"
  ? enriched.filter((c) => {
    // "ยังไม่สัมภาษณ์" และ "สัมภาษณ์แล้ว" ต้องเช็คจาก Queue_Admin_Status
    if (statusFilter === "ยังไม่สัมภาษณ์" || statusFilter === "สัมภาษณ์แล้ว") {
      return c.Queue_Admin_Status === statusFilter;
    }
    // "เช็คชื่อแล้ว" ต้องรวมทั้งคนที่ "สัมภาษณ์แล้ว" และ "ยังไม่สัมภาษณ์"
    // เพราะทั้งสองสถานะนี้คือคนที่เช็คชื่อเข้างานแล้วทั้งคู่
    if (statusFilter === "เช็คชื่อแล้ว") {
      return (
        c.Queue_Admin_Status === "ยังไม่สัมภาษณ์" ||
        c.Queue_Admin_Status === "สัมภาษณ์แล้ว"
      );
    }
    // ส่วน "ไม่เข้าร่วม" เช็คจาก Queue_Display_Status
    return c.Queue_Display_Status === statusFilter;
  })
  : enriched;

    const totalItems = filteredByStatus.length;
    const totalPages = Math.max(1, Math.ceil(totalItems / limitNum));
    const startIndex = (pageNum - 1) * limitNum;
    const pagedData = filteredByStatus.slice(startIndex, startIndex + limitNum);

    return {
      success: true,
      data: pagedData,
      pagination: {
        totalItems,
        totalPages,
        currentPage: pageNum,
        limit: limitNum,
        hasNextPage: pageNum < totalPages,
        hasPreviousPage: pageNum > 1,
      },
      counts,
    };
  } catch (error) {
    console.error("Error in getQueueCandidates:", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * ส่งอีเมลแจ้งคิว: พยายามส่งผ่าน GmailApp เป็นหลักเหมือนเดิมทุกอย่าง
 * (subject, htmlBody, sender name เหมือนเดิม) ถ้าเกิด error ตอนส่ง
 * (เช่น GmailApp quota หมดรายวัน, network error) จะ fallback ไปส่งผ่าน
 * MailApp.sendEmail แทนโดยอัตโนมัติ แทนที่จะโยน error กลับไปให้ผู้ใช้เห็น
 * ทันที (พฤติกรรมเดิม)
 *
 * หมายเหตุ: GmailApp และ MailApp ใช้ quota "อีเมล/วัน" ก้อนเดียวกันของ
 * บัญชี Google ที่รันสคริปต์ (ไม่เหมือนกรณี UrlFetchApp vs MailApp ใน
 * services/sendEmail.js ที่ quota คนละก้อนกัน) ดังนั้น fallback นี้จะ
 * ช่วยกรณี GmailApp error ชั่วคราว/รูปแบบเรียกผิดพลาดเป็นหลัก ไม่ได้ช่วย
 * กรณีโควต้าอีเมลรายวันหมดจริงๆ (ถ้าหมดจริง จะ fail ทั้งสองช่องทาง) —
 * จึงเพิ่มการเช็ค remaining quota ก่อนเพื่อคืนข้อความแจ้งเตือนที่ชัดเจน
 * แทนการโยน error ดิบๆ กลับไป
 *
 * Return shape เหมือนเดิมทุกอย่าง { success, message } เพิ่มแค่ field
 * "channel" (ไม่บังคับใช้ ฝั่งเรียกเดิมไม่ต้องแก้อะไร) บอกว่าส่งผ่าน
 * ช่องทางไหนสำเร็จ
 */
function sendQueueEmail(payload) {
  if (!payload) return { success: false, message: "No data provided" };

  const eventName = payload.eventName || "กิจกรรมสัมภาษณ์งาน";
  const fullname = payload.fullname || "";
  const queueSubject = `[COM7 Event] แจ้งลำดับคิวสัมภาษณ์งาน: ${eventName} (คุณ ${fullname})`;

  // 1) ช่องทางหลัก: GmailApp (พฤติกรรมเดิมทุกอย่าง)
  try {
    GmailApp.sendEmail(payload.to, queueSubject, "", {
      htmlBody: payload.htmlBody,
      name: "COM7 Event"
    });

    return { success: true, message: "ส่งอีเมลแจ้งคิวสำเร็จผ่าน GmailApp", channel: "gmail" };
  } catch (error) {
    console.error("sendQueueEmail: GmailApp failed (" + error.message + "). Falling back to MailApp.");
  }

  // 2) ช่องทางสำรอง: MailApp (เนื้อหา/หัวข้ออีเมลเหมือนเดิมทุกอย่าง)
  try {
    let remainingQuota = null;
    try {
      remainingQuota = MailApp.getRemainingDailyQuota();
    } catch (e) {
      console.warn("sendQueueEmail fallback: cannot read MailApp quota: " + e.message);
    }

    if (remainingQuota !== null && remainingQuota <= 0) {
      console.error("sendQueueEmail fallback: MailApp daily quota exhausted too.");
      return {
        success: false,
        message: "โควต้าอีเมลของระบบหมดสำหรับวันนี้ (ทั้ง GmailApp และ MailApp)",
        channel: "mailapp-fallback",
      };
    }

    MailApp.sendEmail({
      to: payload.to,
      subject: queueSubject,
      htmlBody: payload.htmlBody,
      name: "COM7 Event",
    });

    return { success: true, message: "ส่งอีเมลแจ้งคิวสำเร็จผ่านช่องทางสำรอง (MailApp)", channel: "mailapp-fallback" };
  } catch (error) {
    console.error("Error in sendQueueEmail (fallback):", error);
    return { success: false, message: error.message };
  }
}