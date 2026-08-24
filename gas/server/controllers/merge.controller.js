function getFullEventStructure() {
  // แคช 3 นาที ลดจำนวนครั้งที่ยิง UrlFetchApp — ฟังก์ชันนี้ถูกเรียกจาก
  // หน้าแรกและหน้าค้นหา event ซึ่งมีผู้เข้าใช้งานจำนวนมาก
  const cacheKey = "full_event_structure";
  const cache = CacheService.getScriptCache();
  try {
    const cached = getCacheChunked(cache, cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {
    console.warn("Cannot read full_event_structure cache: " + e.message);
  }

  try {
    const eventSheet = SpreadsheetApp.openById(MAIN_SS_ID).getSheetByName(
      MAIN_SHEETS.EVENTS,
    );
    const eventRows = getRowsFromSheet(eventSheet);

    // เดิมกรองเฉพาะ Event ที่ยังอยู่ในช่วง Opening_At < now < Closing_At
    // (เปิดรับสมัครอยู่ตอนนี้เท่านั้น) ตามคำขอ ให้แสดงทุก Event โดยไม่กรอง
    // ตามช่วงวันที่ ยังคงกรองเฉพาะ Status ที่เผยแพร่แล้วไว้เหมือนเดิม
    const filteredEventRows = eventRows.filter((event) => {
      return (
        String(event.Status).toLowerCase() ===
        String(EVENT_STATUS[0]).toLowerCase()
      );
    });

    if (filteredEventRows.length === 0) {
      return { success: true, data: [] };
    }

    // แก้ไข: เดิมยิง UrlFetchApp.fetchAll() ตรงไปที่ Sheets REST API ทำให้
    // ชนโควต้ารายวันของ UrlFetchApp ("Service invoked too many times for one
    // day: urlfetch") ฟังก์ชันนี้ถูกเรียกทุกครั้งที่มีคนเปิดหน้าแรก จึงเป็น
    // ตัวการหลักที่ทำให้โควต้าหมดเร็ว เปลี่ยนมาใช้ Advanced Sheets Service
    // แทนเพราะใช้โควต้าคนละก้อนกับ UrlFetchApp
    const fullStructure = filteredEventRows.map((data) => {
      const eventObj = {
        eventId: data.Sheet_ID,
        eventName: data.Event_Name,
        eventType: data.Event_Type,
        eventJobType: data.Job_Type,
        eventStart: normalizeDate(data.Opening_At),
        eventEnd: normalizeDate(data.Closing_At),
        slots: [],
      };

      try {
        const res = callSheetsApiWithBackoff(() =>
          Sheets.Spreadsheets.Values.get(
            data.Sheet_ID,
            EVENT_SHEETS.SLOTS,
            { valueRenderOption: "FORMATTED_VALUE" },
          ),
        );
        const values = res.values;
        if (values && values.length > 1) {
          const slotValues = values.slice(1);
          const slotsMap = new Map();

          slotValues.forEach((row) => {
            const [uuid, date, startTime, endTime, buName, capacity] = row;
            const uniqueSlotKey = `${date}_${startTime}_${endTime}`;

            if (!slotsMap.has(uniqueSlotKey)) {
              slotsMap.set(uniqueSlotKey, {
                slotId: uniqueSlotKey,
                date: date,
                startTime: startTime,
                endTime: endTime,
                bus: [],
              });
            }

            slotsMap.get(uniqueSlotKey).bus.push({
              allocationId: uuid,
              buName: buName,
              capacity: parseInt(capacity) || 0,
            });
          });

          eventObj.slots = Array.from(slotsMap.values());
        }
      } catch (innerErr) {
        console.warn(
          `Cannot fetch slots for event ${eventObj.eventName}: ${innerErr.message}`,
        );
      }

      return eventObj;
    });

    const result = JSON.parse(
      JSON.stringify({
        success: true,
        data: fullStructure,
      }),
    );

    try {
      // ใช้ putCacheChunked แทน cache.put() ตรงๆ เพราะข้อมูลรวมทุก event
      // พร้อม slots ทั้งหมด มีโอกาสโตเกินลิมิต 100KB ของ CacheService ต่อ key
      // เดียวเมื่อจำนวน event/slot เพิ่มขึ้น (pattern เดียวกับ raw_interviews)
      putCacheChunked(cache, cacheKey, JSON.stringify(result), 900); // 15 นาที (เพิ่มจาก 3 นาที เพื่อลดจำนวนครั้งที่ยิง Sheets API)
    } catch (e) {
      console.warn("Cannot write full_event_structure cache: " + e.message);
    }

    return result;
  } catch (err) {
    console.error("Error in getEventAndSlots controller: ", err.message);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function getFullManagerStructure(query = {}) {
  const { permission, username, page = 1, limit = 15 } = query;

  const pageNum = Math.max(1, parseInt(page));
  const limitNum = Math.max(1, parseInt(limit));

  if (!permission || !username) {
    return {
      success: false,
      message: "ไม่พบข้อมูลนี้",
    };
  }

  try {
    // แก้ไข: เดิมยิง UrlFetchApp.fetch() ตรงไปที่ Sheets REST API ทำให้ชน
    // โควต้ารายวันของ UrlFetchApp เปลี่ยนมาใช้ Advanced Sheets Service แทน
    let res;
    try {
      res = Sheets.Spreadsheets.Values.get(
        permission,
        EVENT_SHEETS.TRANSACTION,
        { valueRenderOption: "FORMATTED_VALUE" },
      );
    } catch (e) {
      return {
        success: false,
        message: "ดึงข้อมูลจาก Sheet ล่าช้าหรือไม่ตอบสนอง",
      };
    }

    const transactionRows = mergeMapValuesToRows(res.values);

    const filteredData = transactionRows.filter(
      (data) => String(data.Username) === String(username),
    );

    filteredData.sort((a, b) => {
      const timeA = new Date(normalizeDate(a.TimeStamp)).getTime() || 0;
      const timeB = new Date(normalizeDate(b.TimeStamp)).getTime() || 0;
      return timeB - timeA;
    });

    const totalItems = filteredData.length;
    const totalPages = Math.ceil(totalItems / limitNum);

    const startIndex = (pageNum - 1) * limitNum;
    const endIndex = startIndex + limitNum;

    const paginatedData = filteredData.slice(startIndex, endIndex);

    const result = {
      success: true,
      data: paginatedData,
      pagination: {
        totalItems: totalItems,
        totalPages: totalPages,
        currentPage: pageNum,
        limit: limitNum,
        hasNextPage: pageNum < totalPages,
        hasPreviousPage: pageNum > 1,
      },
    };

    return JSON.parse(JSON.stringify(result));
  } catch (err) {
    console.error("Error in getFullManagerSructure: ", err.message);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด: " + err.message,
    };
  }
}

function getEventAndSlots(data) {
  try {
    const { eventId, dateQuery } = data;
    const event = getEventById(eventId);
    if (!event.success) return event;

    const eventDetail = event.data.eventdata;
    let selectedDate;
    let slotsData;

    if (dateQuery) {
      // Specific date requested — single read
      const slots = getAllSlotsData({ date: dateQuery }, eventId);
      if (!slots.success) return slots;
      const dateData =
        slots.data.dates && slots.data.dates.length > 0
          ? slots.data.dates[0]
          : null;
      slotsData = dateData ? dateData.slots : [];

      if (slotsData.length === 0) {
        // Queried date is now empty (e.g. after deletion) — fall back to earliest available date
        const allSlots = getAllSlotsData({}, eventId);
        if (allSlots.success && allSlots.data.dates && allSlots.data.dates.length > 0) {
          const sortedDates = allSlots.data.dates.sort(
            (a, b) => new Date(a.date) - new Date(b.date),
          );
          selectedDate = sortedDates[0].date;
          slotsData = sortedDates[0].slots;
        } else {
          selectedDate = dateQuery;
        }
      } else {
        selectedDate = dateQuery;
      }
    } else {
      // No date — read once, pick the earliest future date from the result
      const allSlots = getAllSlotsData({}, eventId);
      if (!allSlots.success) return allSlots;

      if (allSlots.data.dates && allSlots.data.dates.length > 0) {
        const sortedDates = allSlots.data.dates.sort(
          (a, b) => new Date(a.date) - new Date(b.date),
        );
        selectedDate = sortedDates[0].date;
        slotsData = sortedDates[0].slots;
      } else {
        selectedDate = new Date().toISOString().split("T")[0];
        slotsData = [];
      }
    }

    return JSON.parse(
      JSON.stringify({
        success: true,
        data: { event: eventDetail, slots: slotsData, selectedDate },
      }),
    );
  } catch (err) {
    console.error("Error in getEventAndSlots controller: ", err.message);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function getEventAndSlotData(data) {
  try {
    if (!data) {
      return {
        success: false,
        message: "ไม่พบข้อมูลนี้",
      };
    }

    const { eventId, slotId } = data;
    if (!eventId || !slotId) {
      return {
        success: false,
        message: "ไม่พบข้อมูลนี้",
      };
    }

    const event = getEventById(eventId);
    if (event.success === false) {
      return {
        success: event.success,
        message: event.message,
      };
    }

    const slot = getSlotById(eventId, slotId);
    if (slot.success === false) {
      return {
        success: event.success,
        message: event.message,
      };
    }

    const eventDetail = event.data.eventdata;

    const response = {
      success: true,
      data: {
        event: eventDetail,
        slot: slot.data,
      },
    };

    return JSON.parse(JSON.stringify(response));
  } catch (error) {
    console.error("Error in getEventAndSlotData controller: ", error.message);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function getEventAndBuStructure() {
  // ใส่แคช (5 นาที) เพื่อลดจำนวนครั้งที่ยิง UrlFetchApp ไปยัง Sheets API
  // ฟังก์ชันนี้ถูกเรียกจากหลายหน้า รวมถึงหน้าแรกที่ผู้ใช้ทุกคนเข้าถึงได้
  // ถ้าไม่แคช จะยิง request ใหม่ทุกครั้งที่มีคนเปิดหน้า จน quota รายวันของ
  // UrlFetchApp หมด (Service invoked too many times for one day: urlfetch)
  const cacheKey = "event_bu_structure";
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {
    console.warn("Cannot read event_bu_structure cache: " + e.message);
  }

  // แก้ไข: ป้องกัน "cache stampede" — เดิมถ้ามีหลาย request เข้ามาพร้อมกัน
  // ตอน cache หมดอายุพอดี (เช่นตอนเปิดหน้าแรกพร้อมกันหลายคน) แต่ละ request
  // จะวิ่งเข้ามาดึงข้อมูลจาก Sheets API ซ้ำกันทั้งหมด (คนละก้อน ทวีคูณตาม
  // จำนวนคน) จนชนโควต้า "Read requests per minute per user" ของ
  // sheets.googleapis.com ใช้ LockService เพื่อให้มีแค่ execution เดียวที่
  // ยิง Sheets API จริง ส่วนที่เหลือรอแล้วอ่านจาก cache ที่ execution แรก
  // เขียนไว้แทน (ผลลัพธ์ที่ return ออกไปเหมือนเดิมทุกประการ)
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    console.warn("Cannot acquire lock for event_bu_structure: " + e.message);
  }

  try {
    // เช็ค cache อีกครั้งหลังได้ lock เผื่อ execution ก่อนหน้าเพิ่งเขียนเสร็จ
    try {
      const cachedAfterLock = cache.get(cacheKey);
      if (cachedAfterLock) return JSON.parse(cachedAfterLock);
    } catch (e) {
      console.warn("Cannot read event_bu_structure cache: " + e.message);
    }

    const eventSheet = callSheetsApiWithBackoff(() =>
      Sheets.Spreadsheets.Values.batchGet(MAIN_SS_ID, {
        ranges: [MAIN_SHEETS.EVENTS],
        valueRenderOption: "FORMATTED_VALUE",
      }),
    );
    const eventData = mergeMapValuesToRows(eventSheet.valueRanges[0].values);

    if (eventData.length === 0) return [];

    // แก้ไข: เดิมยิง UrlFetchApp.fetchAll() ตรงไปที่ Sheets REST API ทำให้
    // ชนโควต้ารายวันของ UrlFetchApp เปลี่ยนมาใช้ Advanced Sheets Service แทน
    // (และห่อด้วย callSheetsApiWithBackoff เพื่อ retry อัตโนมัติถ้าชน
    // quota "Read requests per minute" ซึ่งเป็น error ชั่วคราว)
    const payload = eventData.map((row) => {
      const eventId = row.Sheet_ID;
      const eventName = row.Event_Name;
      let buList = [];

      try {
        const res = callSheetsApiWithBackoff(() =>
          Sheets.Spreadsheets.Values.get(row.Sheet_ID, EVENT_SHEETS.SLOTS, {
            valueRenderOption: "FORMATTED_VALUE",
          }),
        );
        const allSlotData = mergeMapValuesToRows(res.values);

        const uniqueBuNames = [
          ...new Set(
            allSlotData
              .map((data) => (data.Bu_Name ? String(data.Bu_Name).trim() : ""))
              .filter((name) => name),
          ),
        ];

        buList = uniqueBuNames.map((name) => ({
          buName: name,
        }));
      } catch (e) {
        console.warn(`Error getting BUs for ${eventName}: ${e.message}`);
      }

      return {
        eventId: eventId,
        eventName: eventName,
        bu: buList,
      };
    });

    const sortedPayload = payload.sort((a, b) =>
      a.eventName.localeCompare(b.eventName),
    );

    try {
      cache.put(cacheKey, JSON.stringify(sortedPayload), 900); // 15 นาที (เพิ่มจาก 5 นาที)
    } catch (e) {
      console.warn("Cannot write event_bu_structure cache: " + e.message);
    }

    return sortedPayload;
  } catch (error) {
    console.error(
      "Error in getEventAndBuStructure controller: ",
      error.message,
    );
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // ไม่มี lock ให้ release ก็ไม่เป็นไร (เช่นตอน waitLock ล้มเหลวไปแล้ว)
    }
  }
}

/**
 * ดึงข้อมูล Dashboard รองรับ Filter: Date, BU, และ EventName
 */
function getAllDataForDashBoard(query = {}) {
  try {
    // Return cached result if available (10-minute TTL)
    const cacheKey =
      "dashboard_" +
      (query.eventName || "") +
      "|" +
      (query.buName || "") +
      "|" +
      (query.startDate || "") +
      "|" +
      (query.endDate || "");
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }

    const result = {
      kpi: {
        totalEvents: 0,
        totalCandidates: 0,
        totalOffers: 0,
        conversionRate: 0,
        totalCapacity: 0,
        capacityUtilization: 0,
      },
      charts: {
        statusDistribution: {},
        candidateTrend: {
          labels: [],
          data: [],
        },
      },
    };

    const formatDateKey = (dateInput) => {
      if (!dateInput) return null;
      const date = new Date(dateInput);
      return Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
    };

    const filterEventName = query.eventName || null;
    const filterBu = query.buName || null;
    const filterStart = query.startDate
      ? new Date(query.startDate).getTime()
      : null;
    const filterEnd = query.endDate ? new Date(query.endDate).getTime() : null;

    const eventSheet = getEventSheet();
    let eventData = getRowsFromSheet(eventSheet);

    if (!eventData || eventData.length === 0)
      return { success: true, payload: result };

    if (filterEventName) {
      eventData = eventData.filter((e) => e.Event_Name === filterEventName);
    }

    result.kpi.totalEvents = eventData.length;
    const trendMap = {};

    const VALID_STATUSES = new Set([
      "ผ่าน",
      "ไม่ผ่าน",
      "รอการพิจารณา",
      "เก็บไว้พิจารณา",
      "ไม่ผ่านการพิจารณา",
    ]);

    // แก้ไข: เดิมยิง UrlFetchApp.fetchAll() ตรงไปที่ Sheets REST API ทำให้ชน
    // โควต้ารายวันของ UrlFetchApp เปลี่ยนมาใช้ Advanced Sheets Service แทน
    // เพราะใช้โควต้าคนละก้อนกับ UrlFetchApp (ยังคง cache ผลลัพธ์ 10 นาทีไว้
    // ด้านบนเหมือนเดิมเพื่อลดจำนวนครั้งที่ต้องยิงจริงๆ)
    const ranges = [EVENT_SHEETS.SLOTS, EVENT_SHEETS.INTERVIEWS];

    eventData.forEach((event) => {
      try {
        const response = Sheets.Spreadsheets.Values.batchGet(
          event.Sheet_ID,
          {
            ranges: ranges,
            valueRenderOption: "FORMATTED_VALUE",
          },
        );
        if (!response.valueRanges || response.valueRanges.length < 2) return;

        const slotValues = response.valueRanges[0].values;
        const intValues = response.valueRanges[1].values;

        // --- Slots: sum Capacity column (col 6 / index 5, same layout as before) ---
        if (slotValues && slotValues.length > 1) {
          for (let i = 1; i < slotValues.length; i++) {
            const cap = slotValues[i][5];
            result.kpi.totalCapacity += parseInt(cap) || 0;
          }
        }

        // --- Interviews: detect column positions from header row (works for
        // both old sheets and new sheets with the extra Shop column) ---
        if (intValues && intValues.length > 1) {
          const intHeaders = intValues[0];
          const buColIdx = intHeaders.indexOf("Bu_Name");
          const statusColIdx = intHeaders.indexOf("Status");
          const tsColIdx = intHeaders.indexOf("TimeStamp");

          if (buColIdx === -1 || statusColIdx === -1 || tsColIdx === -1) {
            console.warn("Missing expected headers in interviews sheet for event: " + event.Event_Name);
            return;
          }

          for (let i = 1; i < intValues.length; i++) {
            const row = intValues[i];
            const buName = String(row[buColIdx] || "");
            const status = String(row[statusColIdx] || "");
            const timeStampStr = String(row[tsColIdx] || "");

            if (!status && !timeStampStr) continue;
            if (!VALID_STATUSES.has(status)) continue;

            const cleanTimestamp = timeStampStr.replace(/\s+/g, " ").trim();
            const parsedDate = parseThaiDateTime(cleanTimestamp);
            const interviewDate = isNaN(parsedDate.getTime())
              ? 0
              : parsedDate.getTime();

            const dateValid =
              (!filterStart || interviewDate >= filterStart) &&
              (!filterEnd || interviewDate <= filterEnd);
            const buValid = !filterBu || buName === filterBu;

            if (!dateValid || !buValid) continue;

            result.kpi.totalCandidates++;

            result.charts.statusDistribution[status] =
              (result.charts.statusDistribution[status] || 0) + 1;

            if (status === "ผ่าน") {
              result.kpi.totalOffers++;
            }

            const dateKey = formatDateKey(parsedDate);
            if (dateKey) {
              trendMap[dateKey] = (trendMap[dateKey] || 0) + 1;
            }
          }
        }
      } catch (e) {
        console.warn(
          "Error processing event " + event.Event_Name + ": " + e.message,
        );
      }
    });

    const sortedDates = Object.keys(trendMap).sort();
    result.charts.candidateTrend.labels = sortedDates;
    result.charts.candidateTrend.data = sortedDates.map(
      (date) => trendMap[date],
    );

    if (result.kpi.totalCandidates > 0) {
      result.kpi.conversionRate = (
        (result.kpi.totalOffers / result.kpi.totalCandidates) *
        100
      ).toFixed(2);
    }
    if (result.kpi.totalCapacity > 0) {
      result.kpi.capacityUtilization = (
        result.kpi.totalCandidates / result.kpi.totalCapacity
      ).toFixed(2);
    }

    const finalResult = { success: true, payload: result };

    try {
      cache.put(cacheKey, JSON.stringify(finalResult), 600);
    } catch (cacheErr) {
      console.warn("Dashboard cache write failed: " + cacheErr.message);
    }

    return finalResult;
  } catch (error) {
    console.error("Error in getAllDataForDashBoard: ", error.message);
    return { success: false, message: error.message };
  }
}