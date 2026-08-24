const EVENT_STATUS = ["TRUE", "FALSE"];
const EVENT_TYPE = ["ONSITE", "ONLINE"];
const ONSITE_INTERVIEW_IMAGE =
  "https://media.istockphoto.com/id/1298405314/vector/job-interview.jpg?s=612x612&w=0&k=20&c=F3P4brlXN7S35fe73OrxrKs0-FMc3VoMSuv6I6VIcGg=";
const ONLINE_INTERVIEW_IMAGE =
  "https://media.licdn.com/dms/image/v2/D4D12AQGlAzsMD0-5uA/article-cover_image-shrink_720_1280/article-cover_image-shrink_720_1280/0/1691685823372?e=2147483647&v=beta&t=i5FXGmJnVPEXw0h8ZOHZHQxlKnkMyfZegBjG0y5o__E";

function getEventsFolder() {
  return DriveApp.getFolderById(EVENTS_FOLDER_ID);
}

function getEventSheet() {
  return SpreadsheetApp.openById(MAIN_SS_ID).getSheetByName(MAIN_SHEETS.EVENTS);
}

function getJobTypeSheet() {
  return SpreadsheetApp.openById(MAIN_SS_ID).getSheetByName(
    MAIN_SHEETS.JOB_TYPE,
  );
}

function resolveCandidateName(event) {
  const candidateNames = [
    event?.Fullname,
    event?.Name,
    event?.Candidate_Name,
    event?.candidateName,
    event?.fullname,
    event?.userName,
    event?.username,
  ];

  const foundName = candidateNames.find((value) => {
    if (value === undefined || value === null) return false;
    return String(value).trim() !== "";
  });

  return foundName ? String(foundName).trim() : "";
}

function resolveQueueNo(event) {
  const directValues = [
    event?.queueNo,
    event?.Queue_No,
    event?.queue_no,
    event?.No,
    event?.Queue_Number,
    event?.queueNumber,
  ];

  const directValue = directValues.find((value) => {
    if (value === undefined || value === null) return false;
    return String(value).trim() !== "";
  });

  if (directValue !== undefined) {
    return String(directValue).trim();
  }

  const candidateName = resolveCandidateName(event);
  const eventName = event?.Event_Name || event?.eventName || "";

  if (candidateName && eventName) {
    const queueInfo = findQueueLog(eventName, candidateName);
    if (queueInfo && queueInfo.queueNo !== undefined && queueInfo.queueNo !== null) {
      return String(queueInfo.queueNo).trim();
    }
  }

  return "-";
}

function getEventsData(query = {}, userRole) {
  try {
    const eventSheet = getEventSheet();
    if (!eventSheet) return { success: false, message: "ไม่พบชีทนี้" };

    let eventData = getRowsFromSheet(eventSheet);

    const cleanData = eventData.filter(e => e.Event_Name && e.Event_Name.trim() !== "");

    let filteredData = filteredEventByQuery(cleanData, query, userRole);

    // Sort before paginating so order is consistent across all pages
    filteredData.sort((a, b) => {
      if (userRole === SUPER_ADMIN_ROLE) {
        const dateA = new Date(normalizeDate(a.TimeStamp)).getTime() || 0;
        const dateB = new Date(normalizeDate(b.TimeStamp)).getTime() || 0;
        return dateB - dateA;
      }
      const dateA = new Date(normalizeDate(a.Closing_At)).getTime() || 0;
      const dateB = new Date(normalizeDate(b.Closing_At)).getTime() || 0;
      return dateA - dateB;
    });

    // เพิ่มข้อมูลคิวและสถานะ
    const enriched = filteredData.map(event => {
      const candidateName = resolveCandidateName(event);
      const queueInfo = candidateName
        ? findQueueLog(event.Event_Name || "", candidateName)
        : null;
      const resolvedQueueNo = resolveQueueNo(event);

      return {
        ...event,
        queueNo: resolvedQueueNo,
        Queue_No: resolvedQueueNo,
        joinedAt: queueInfo ? queueInfo.joinedAt : null,
        Participation_Status: event.Participation_Status || "ยังไม่เช็คชื่อ"
      };
    });

    // Pagination
    const page = Number(query.page) || 1;
    const limit = Number(query.limit) || 10;
    const skip = (page - 1) * limit;
    const total = enriched.length;

    const paginatedData = enriched.slice(skip, skip + limit);

    return {
      success: true,
      data: paginatedData,
      pagination: {
        currentPage: page,
        limit,
        total,
        totalPages: Math.ceil(total / limit)
      }
    };
  } catch (error) {
    console.error("Error in getEventsData:", error);
    return { success: false, message: error.message };
  }
}

// คืน Event ทั้งหมดในระบบ (ไม่กรองตามการลงทะเบียนของ user) ใช้สำหรับหน้าตรวจสอบลำดับคิว
function getAllEventsForQueue() {
  try {
    const eventSheet = getEventSheet();
    if (!eventSheet) {
      return { success: false, message: "ไม่พบชีทนี้" };
    }

    const eventData = getRowsFromSheet(eventSheet);

    const cleanData = eventData.filter(
      (event) => event.Event_Name && String(event.Event_Name).trim() !== "",
    );

    // เรียงตามวันที่ล่าสุดก่อน
    cleanData.sort((a, b) => {
      const dateA = new Date(normalizeDate(a.Date)).getTime() || 0;
      const dateB = new Date(normalizeDate(b.Date)).getTime() || 0;
      return dateB - dateA;
    });

    return {
      success: true,
      data: cleanData,
    };
  } catch (error) {
    console.error("Error in getAllEventsForQueue:", error);
    return { success: false, message: error.message };
  }
}

function createEvent(eventData) {
  try {
    // Validate data
    if (!eventData) {
      return {
        success: false,
        message: "กรุณากรอกข้อมูลให้ครบถ้วน",
      };
    }

    const {
      eventName,
      type,
      location,
      jobType,
      openingDate,
      openingTime,
      closingDate,
      closingTime,
      userRole,
    } = eventData;

    // Validate event name
    if (
      !eventName ||
      !type ||
      !jobType ||
      !openingDate ||
      !openingTime ||
      !closingDate ||
      !closingTime
    ) {
      return {
        success: false,
        message: "กรุณากรอกข้อมูลให้ครบถ้วน",
      };
    }

    // Validate user role
    if (userRole === USER_ROLE || userRole === MANAGER_ROLE) {
      return {
        success: false,
        message: "คุณไม่สามารถสร้างชีทนี้ได้",
      };
    }

    // Validate event name
    const isEventNameExists = checkEventNameExists(eventName);
    if (isEventNameExists) {
      return {
        success: false,
        message: "ชื่ออีเวนต์นี้ถูกใช้ไปแล้ว",
      };
    }

    if (!EVENT_TYPE.includes(type)) {
      return {
        success: false,
        message: "ประเภทอีเวนต์ไม่ถูกต้อง",
      };
    }

    const jobTypeSheet = getJobTypeSheet();
    const jobTypeRows = getRowsFromSheet(jobTypeSheet);

    const isJobTypeExists = jobTypeRows.some((job) => job.Name === jobType);
    if (!isJobTypeExists) {
      return {
        success: false,
        message: "ประเภทงานไม่ถูกต้อง",
      };
    }

    // Create spreadsheet
    const { ss_id, ss_url } = createSpreadSheet(eventData);
    if (!ss_id && !ss_url) {
      return {
        success: false,
        message: "เกิดข้อผิดพลาดในการสร้างชีท",
      };
    }

    // Create payload
    const openingAt = combineDateTime(openingDate, openingTime);
    const closingAt = combineDateTime(closingDate, closingTime);

    if (openingAt >= closingAt) {
      return {
        success: false,
        message: "วันเวลาปิดอีเวนต์ต้องมากกว่าวันเวลาเปิดอีเวนต์",
      };
    }

    const formatedOpeningAt = openingAt.toLocaleString("th-TH", {
      timeZone: "Asia/Bangkok",
    });
    const formatedClosingAt = closingAt.toLocaleString("th-TH", {
      timeZone: "Asia/Bangkok",
    });
    const now = new Date().toLocaleString("th-TH", {
      timeZone: "Asia/Bangkok",
    });
    const eventImage =
      type === EVENT_TYPE[0] ? ONSITE_INTERVIEW_IMAGE : ONLINE_INTERVIEW_IMAGE;

    // Create new event
    const newEvent = [
      ss_id,
      eventName,
      type,
      formatedOpeningAt,
      formatedClosingAt,
      location,
      jobType,
      ss_url,
      eventImage,
      EVENT_STATUS[0],
      now,
    ];

    // Append new event
    const eventSheet = getEventSheet();
    if (!eventSheet) {
      return {
        success: false,
        message: "ไม่พบชีทนี้",
      };
    }

    eventSheet.appendRow(newEvent);

    try {
      CacheService.getScriptCache().remove("dropdown_events_raw");
    } catch (e) {
      console.warn("Cannot invalidate dropdown_events_raw cache: " + e.message);
    }

    return {
      success: true,
      message: "สร้างอีเวนต์สําเร็จ",
    };
  } catch (error) {
    console.error("Error in getEventsSheetName: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function getEventById(eventId) {
  try {
    if (!eventId) {
      return { success: false, message: "ไม่พบข้อมูลอีเวนต์" };
    }

    // === จุดที่แก้: เช็ค cache ก่อนยิง batchGet ===
    const cache = CacheService.getScriptCache();
    const cacheKey = "event_detail_" + eventId;
    try {
      const cached = getCacheChunked(cache, cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (e) {
      console.warn("Cannot read event_detail cache: " + e.message);
    }

    const responseMain = callSheetsApiWithBackoff(() =>
      Sheets.Spreadsheets.Values.batchGet(MAIN_SS_ID, {
        ranges: [MAIN_SHEETS.EVENTS],
        valueRenderOption: "FORMATTED_VALUE",
      }),
    );

    const eventSheetData =
      mergeMapValuesToRows(responseMain.valueRanges[0].values) || [];
    const eventData = eventSheetData.find(
      (event) => event.Sheet_ID === eventId,
    );

    if (!eventData) {
      return { success: false, message: "ไม่พบข้อมูลอีเวนต์" };
    }

    const responseEvent = callSheetsApiWithBackoff(() =>
      Sheets.Spreadsheets.Values.batchGet(eventId, {
        ranges: [EVENT_SHEETS.HISTORY],
        valueRenderOption: "FORMATTED_VALUE",
      }),
    );

    const historyValues =
      mergeMapValuesToRows(responseEvent.valueRanges[0].values) || [];

    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) {
      return { success: false, message: "ไม่พบชีทนี้" };
    }

    const ssName = ss.getName();
    const ssUrl = ss.getUrl();
    const ssId = ss.getId();

    const payload = {
      success: true,
      data: { eventData, ssName, ssUrl, ssId, historyValues },
    };

    // === จุดที่แก้: เก็บ cache ไว้ 5 นาที (300 วิ) ===
    try {
      putCacheChunked(cache, cacheKey, JSON.stringify(payload), 300);
    } catch (e) {
      console.warn("Cannot write event_detail cache: " + e.message);
    }

    return payload;
  } catch (error) {
    console.error("Error in getEventById: ", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด" };
  }
}

function getHistoryByEventId(eventId) {
  try {
    if (!eventId) {
      return { success: false, message: "ไม่พบข้อมูลอีเวนต์" };
    }

    // === จุดที่แก้: เช็ค cache ก่อนยิง batchGet (TTL สั้นกว่า event detail เพราะ history เปลี่ยนบ่อยกว่า) ===
    const cache = CacheService.getScriptCache();
    const cacheKey = "event_history_" + eventId;
    try {
      const cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (e) {
      console.warn("Cannot read event_history cache: " + e.message);
    }

    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) {
      return { success: false, message: "ไม่พบชีทนี้" };
    }

    const historySheet = ss.getSheetByName(EVENT_SHEETS.HISTORY);
    if (!historySheet) {
      return { success: false, message: "ไม่พบชีทประวัติ" };
    }

    const responseEvent = callSheetsApiWithBackoff(() =>
      Sheets.Spreadsheets.Values.batchGet(eventId, {
        ranges: [EVENT_SHEETS.HISTORY],
        valueRenderOption: "FORMATTED_VALUE",
      }),
    );

    const historyValues =
      mergeMapValuesToRows(responseEvent.valueRanges[0].values) || [];

    const sortedHistory = historyValues.slice().reverse();

    const payload = { success: true, data: sortedHistory };

    // === จุดที่แก้: เก็บ cache 90 วิ (สั้นเพราะ history อาจถูกเพิ่มบ่อย) ===
    try {
      cache.put(cacheKey, JSON.stringify(payload), 90);
    } catch (e) {
      console.warn("Cannot write event_history cache: " + e.message);
    }

    return payload;
  } catch (error) {
    console.error("Error in getHistoryByEventId: ", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด" };
  }
}

function updateEventData(eventId, eventData) {
  try {
    // Validate the data
    if (!eventId || !eventData) {
      return {
        success: false,
        message: "กรุณากรอกข้อมูลให้ครบถ้วน",
      };
    }

    const { eventName, userName, userRole, action } = eventData;

    // Validate user role
    if (userRole === USER_ROLE || userRole === MANAGER_ROLE) {
      return {
        success: false,
        message: "คุณไม่สามารถแก้ไขชีทนี้ได้",
      };
    }

    // Get spreadsheet
    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) {
      return {
        success: false,
        message: "ไม่พบชีทนี้",
      };
    }

    // Update spreadsheet name
    eventName && ss.setName(eventName);

    // Update sheet name in database
    const isUpdateInMainSheet = updateEventDataInMainSheet(eventId, eventData);
    if (isUpdateInMainSheet.success === false) {
      return {
        success: isUpdateInMainSheet.success,
        message: isUpdateInMainSheet.message,
      };
    }

    const isUpdated = updateHistoryLogs(ss, userName, userRole, action);
    if (isUpdated.success === false) {
      return {
        success: isUpdated.success,
        message: isUpdated.message,
      };
    }

    try {
      CacheService.getScriptCache().remove("dropdown_events_raw");
    } catch (e) {
      console.warn("Cannot invalidate dropdown_events_raw cache: " + e.message);
    }

    return {
      success: true,
      message: "อัปเดตข้อมูลสําเร็จ",
    };
  } catch (error) {
    console.error("Error in getEventById: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function deleteEventData(deleteData) {
  try {
    // Validate the data
    const { eventId, userRole } = deleteData;
    if (!eventId) {
      return {
        success: false,
        message: "ไม่พบข้อมูลอีเวนต์",
      };
    }

    // Validate user role
    if (userRole !== SUPER_ADMIN_ROLE) {
      return {
        success: false,
        message: "คุณไม่สามารถลบชีทนี้ได้",
      };
    }

    // Delete event data from database
    const eventSheet = getEventSheet();
    const eventData = getRowsFromSheet(eventSheet);
    const index = eventData.findIndex((event) => event.Sheet_ID === eventId);

    if (index === -1) {
      return {
        success: false,
        message: "ไม่พบข้อมูลอีเวนต์",
      };
    }

    const targetRow = index + 2;
    eventSheet.deleteRow(targetRow);

    try {
      CacheService.getScriptCache().remove("dropdown_events_raw");
    } catch (e) {
      console.warn("Cannot invalidate dropdown_events_raw cache: " + e.message);
    }

    return {
      success: true,
      message: "ลบข้อมูลสําเร็จ",
    };
  } catch (error) {
    console.error("Error in deleteEventData: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}