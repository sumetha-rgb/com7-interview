function filterSlotsByQuery(eventId, query) {
  const { date, startTime } = query;

  const dateValue = date
    ? String(new Date(date).toLocaleDateString())
    : "";

  const startTimeValue =
    startTime instanceof Date
      ? Utilities.formatDate(startTime, "Asia/Bangkok", "HH:mm")
      : String(startTime);

  //    Get slot sheet
  const slotSheet = SpreadsheetApp.openById(eventId).getSheetByName(
    EVENT_SHEETS.SLOTS
  );

  //    Get slot values
  const slotsData = slotSheet.getDataRange().getValues().slice(1);

  const formattedSlotValues = slotsData.map((slot) => {
    return {
      Slot_Id: slot[0],
      Date: slot[1],
      Start_Time:
        slot[2] instanceof Date
          ? Utilities.formatDate(slot[2], "Asia/Bangkok", "HH:mm")
          : slot[2],
      End_Time:
        slot[3] instanceof Date
          ? Utilities.formatDate(slot[3], "Asia/Bangkok", "HH:mm")
          : slot[3],
      Bu_Name: slot[4],
      Capacity: slot[5],
    };
  });

  //    Filter slot values
  const filteredSlotValues = formattedSlotValues.filter((slot) => {
    if (dateValue === "") {
      const now = toDateOnly(new Date());
      const slotDate = toDateOnly(normalizeDate(slot.Date));
      return slotDate >= now;
    } else {
      const normalizeSlotDate = normalizeDate(slot.Date).toLocaleDateString()
      if (normalizeSlotDate !== dateValue) return false;
    }

    if (startTime && startTimeValue !== "") {
      if (slot.Start_Time !== startTimeValue) return false;
    }

    return true;
  });

  return filteredSlotValues;
}

function formatTime(timeValue) {
  return new Date(timeValue).toLocaleTimeString("en-GB", {
    hour12: false,
    hour: "2-digit",
    minute: "2-digit",
  });
}

function timeToMinutes(time) {
  const [h, m] = time.split(":").map(Number);
  return h * 60 + m;
}

function getMergedSlotsData() {
  // แคช 3 นาที ลดจำนวนครั้งที่ต้องยิง Sheets API จริงๆ — ฟังก์ชันนี้ถูกเรียก
  // จาก checkSlotStatus() ทุกครั้งที่ผู้สมัครเปิด/รีเฟรชหน้าค้นหา event-slot
  // TTL สั้นกว่า getFullEventStructure (15 นาที) เพราะข้อมูล capacity ของ
  // slot เปลี่ยนบ่อยกว่า (ผู้สมัครลงทะเบียน/ยกเลิกตลอดเวลา) จึงไม่อยากให้
  // เก่าเกินไปจนโชว์ที่ว่างผิดพลาด แต่ก็ยังช่วยลดโควต้าได้มากในช่วง burst
  const cacheKey = "merged_slots_data";
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {
    console.warn("Cannot read merged_slots_data cache: " + e.message);
  }

  const eventSheet = getEventSheet();
  const eventData = getRowsFromSheet(eventSheet);

  const activeEvents = eventData.filter(
    (event) => event.Status !== EVENT_STATUS[1]
  );
  if (activeEvents.length === 0) return [];

  // แก้ไข: เดิมยิง UrlFetchApp.fetchAll() ตรงไปที่ Sheets REST API ทำให้ชน
  // โควต้ารายวันของ UrlFetchApp ("Service invoked too many times for one
  // day: urlfetch") ฟังก์ชันนี้ถูกเรียกทุกครั้งที่ผู้สมัครเช็คสถานะ slot
  // (checkSlotStatus) ซึ่งเกิดบ่อยมากฝั่ง candidate-facing เปลี่ยนมาใช้
  // Advanced Sheets Service (Sheets.Spreadsheets.Values.get) แทน เพราะใช้
  // โควต้าคนละก้อนกับ UrlFetchApp เหมือน 5 จุดก่อนหน้าที่แก้ไปแล้ว
  let hasFailedRequest = false;

  const result = activeEvents.map((event) => {
    const eventObj = {
      sheetId: event.Sheet_ID,
      eventName: event.Event_Name,
      eventType: event.Event_Type,
      location: event.Location || "",
      imageUrl: event.Image_Url,
      dates: [],
    };

    try {
      const res = callSheetsApiWithBackoff(() =>
        Sheets.Spreadsheets.Values.get(
          event.Sheet_ID,
          EVENT_SHEETS.SLOTS,
          { valueRenderOption: "FORMATTED_VALUE" },
        ),
      );

      const values = res.values;
      if (!values || values.length <= 1) return eventObj;

      const groupedByDate = {};
      values.slice(1).forEach((row) => {
        const date = String(row[1] || "");
        const startTime = String(row[2] || "");
        const endTime = String(row[3] || "");
        const capacity = row[5] || 0;

        if (!date) return;
        if (!groupedByDate[date]) groupedByDate[date] = [];
        groupedByDate[date].push({ startTime, endTime, capacity });
      });

      for (const [date, slots] of Object.entries(groupedByDate)) {
        eventObj.dates.push({ date, slots });
      }
    } catch (e) {
      hasFailedRequest = true;
      console.warn(
        `Cannot fetch slots for ${event.Event_Name}: ${e.message}`
      );
    }

    return eventObj;
  });

  // ถ้ามี event ไหนดึงไม่สำเร็จ (เช่น ตอน quota เต็ม/Sheets API error ชั่วคราว)
  // ห้ามแคชผลลัพธ์นี้ไว้ ไม่งั้นหน้าเว็บจะโชว์ slot ไม่ครบ/ว่างเปล่าต่อไป
  // อีกหลายนาทีทั้งที่จริงๆ ข้อมูลมีอยู่
  if (!hasFailedRequest) {
    try {
      cache.put(cacheKey, JSON.stringify(result), 180); // 3 นาที
    } catch (e) {
      console.warn("Cannot write merged_slots_data cache: " + e.message);
    }
  }

  return result;
}

function validateSlotData(eventId, slots, date) {
  const eventSheet = getEventSheet();
  const eventData = getRowsFromSheet(eventSheet);
  const event = eventData.find((event) => event.Sheet_ID === eventId);
  if (!event) {
    return {
      success: false,
      message: "ไม่พบชีทนี้",
    };
  }

  const storeOpeningAt = normalizeDate(event.Opening_At);
  const storeClosingAt = normalizeDate(event.Closing_At);
  const requestDate = toDateOnly(new Date(date));

  if (requestDate < toDateOnly(storeOpeningAt) || requestDate > toDateOnly(storeClosingAt)) {
    return {
      success: false,
      message: "กรุณาเลือกวันที่ที่อยู่ในช่วงเปิดรับผู้สมัคร",
    };
  }

  // Insert to sheet
  const slotSheet = SpreadsheetApp.openById(eventId).getSheetByName(
    EVENT_SHEETS.SLOTS
  );
  if (!slotSheet) {
    return {
      success: false,
      message: "ไม่พบชีทนี้",
    };
  }

  const slotData = getRowsFromSheet(slotSheet);

  for (const slot of slots) {
    if (!slot.startTime || !slot.endTime) {
      return {
        success: false,
        message: "ข้อมูลไม่ครบ",
      };
    }

    // Check exists slots
    const isExistsSlot = slotData.find((storeSlot) => {
      const exists = normalizeDate(storeSlot.Date).toLocaleDateString() === new Date(date).toLocaleDateString() && formatTimeTh(new Date(storeSlot.Start_Time)) === slot.startTime && formatTimeTh(new Date(storeSlot.End_Time)) === slot.endTime
      return !!exists
    })
    if (isExistsSlot) {
      return {
        success: false,
        message: "สล็อตนี้ถูกสร้างเปิดแล้ว",
        date: `${new Date(date).toLocaleDateString("th-TH")} ${slot.startTime} - ${slot.endTime}`
      }
    }

    const startSlotDate = combineDateTime(date, slot.startTime);
    const endSlotDate = combineDateTime(date, slot.endTime);

    if (startSlotDate < storeOpeningAt) {
      return {
        success: false,
        message: "ไม่สามารถเปิดสล็อตก่อนเวลาเปิดอีเวนต์ได้",
        date: `${new Date(date).toLocaleDateString("th-TH")} ${slot.startTime}`,
      };
    }

    if (startSlotDate >= storeClosingAt) {
      return {
        success: false,
        message: "ไม่สามารถเปิดสล็อตหลังเวลาเปิดอีเวนต์ได้",
        date: `${new Date(date).toLocaleDateString("th-TH")} ${slot.startTime}`,
      };
    }

    if (endSlotDate > storeClosingAt) {
      return {
        success: false,
        message: "ไม่สามารถเปิดสล็อตช่วงเวลานี้ได้เนื่องจากเกินเวลาเปิดอีเวนต์",
        date: `${new Date(date).toLocaleDateString("th-TH")} ${slot.startTime} - ${slot.endTime}`,
      };
    }
  }

  return true
}