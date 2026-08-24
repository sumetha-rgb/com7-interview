// ==========================================================================
// INTERVIEW SYSTEM CONSTANTS
// ==========================================================================
const INTERVIEW_STATUS = [
  "รอการพิจารณา",
  "ผ่าน",
  "ไม่ผ่าน",
  "เก็บไว้พิจารณา",
  "ไม่ผ่านการพิจารณา",
];

// หมายเหตุ: ส่วนของ PARTICIPATION_STATUS และ QUEUE Constants ถูกย้ายไปที่ QueueService.gs แล้ว
// Google Apps Script สามารถเรียกข้ามไฟล์ในโปรเจกต์เดียวกันได้โดยไม่ต้องประกาศซ้ำ

// ==========================================================================
// INTERVIEW CORE FUNCTIONS
// ==========================================================================

function createInterview(interviewData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const {
      eventId,
      slotId,
      candidateId,
      fullname,
      nickname,
      birthdate,
      age,
      gender,
      weight,
      height,
      nationality,
      nationalId,
      maritalStatus,
      militaryStatus,
      currentAddress,
      phoneNumber,
      email,
      lineId,
      highestEducation,
      institution,
      faculty,
      major,
      workLocation,
      workPeriod,
      workPosition,
      workDescription,
      saleExperience,
      positionType,
      positionApplied,
      preferredProvince,
      preferredDistrict,
      expectedSalary,
      availableStartDate,
      shop,
      file,
    } = interviewData;

    // Validate data
    if (!eventId || !slotId) {
      return { success: false, message: "ไม่พบห้องสัมภาษณ์นี้" };
    }

    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) {
      return { success: false, message: "ไม่พบชีทนี้" };
    }

    const slotSheet = ss.getSheetByName(EVENT_SHEETS.SLOTS);
    if (!slotSheet) {
      return { success: false, message: "ไม่พบชีทนี้" };
    }
    const slotData = getRowsFromSheet(slotSheet);
    const slot = slotData.find((slot) => slot.Slot_Id === slotId);
    if (!slot) {
      return { success: false, message: "สล็อคนี้ยังไม่เปิดให้สัมภาษณ์" };
    }

    // Validate date
    const now = new Date();
    const formattedNowDate = now.toLocaleString("th-TH");

    const slotDate = Utilities.formatDate(
      normalizeDate(slot.Date),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    const targetDate = combineDateTime(
      slotDate,
      formatTimeTh(new Date(slot.End_Time)),
    );

    if (now > targetDate) {
      return { success: false, message: "ช่วงเวลนี้ปิดแล้ว" };
    }

    // Validate user
    const candidateSheet = getCandidateSheet();
    const candidateData = getRowsFromSheet(candidateSheet);
    const candidate = candidateData.find(
      (data) => data.Candidate_Id === candidateId,
    );
    if (!candidate) {
      return { success: false, message: "ไม่พบผู้สมัครนี้" };
    }

    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    const interviewDataAll = getRowsFromSheet(interviewSheet);

    const isUserExists = interviewDataAll.find(
      (row) => row.Candidate_Id === candidateId && row.Slot_Id === slotId,
    );
    if (isUserExists) {
      return { success: false, message: "คุณได้ลงทะเบียนแล้ว" };
    }

    const isFullnameExists = interviewDataAll.find(
      (row) => row.Fullname === fullname,
    );
    if (isFullnameExists) {
      return { success: false, message: "ชื่อนี้ถูกลงทะเบียนไปแล้ว" };
    }

    if (nationalId.length !== 13) {
      return { success: false, message: "เลขบัตรประชาชนไม่ถูกต้อง" };
    }

    const isNationalIdExists = interviewDataAll.find(
      (row) => row.National_Id === nationalId,
    );
    if (isNationalIdExists) {
      return { success: false, message: "เลขบัตรประชาชนนี้ถูกลงทะเบียนไปแล้ว" };
    }

    if (phoneNumber.length !== 10) {
      return { success: false, message: "เบอร์โทรศัพท์ไม่ถูกต้อง" };
    }

    const isPhoneNumberExists = interviewDataAll.find(
      (row) => row.Phone_Number === phoneNumber,
    );
    if (isPhoneNumberExists) {
      return { success: false, message: "เบอร์โทรศัพท์นี้ถูกลงทะเบียนไปแล้ว" };
    }

    const isEmailExists = interviewDataAll.find((row) => row.Email === email);
    if (isEmailExists) {
      return { success: false, message: "อีเมลนี้ถูกลงทะเบียนไปแล้ว" };
    }

    const isLineIdExists = interviewDataAll.find(
      (row) => row.Line_Id === lineId,
    );
    if (isLineIdExists) {
      return { success: false, message: "ไอดีไลน์นี้ถูกลงทะเบียนไปแล้ว" };
    }

    // ----- อัปโหลดไฟล์ หรือ ถ่วงสร้าง PDF -----
    let previewUrl = "";
    let shouldGeneratePDF = false;

    if (file && file.base64Data && file.fileName && file.mimeType) {
      try {
        const fileUrl = uploadFileToDrive(
          file.base64Data,
          file.fileName,
          file.mimeType,
        );
        if (fileUrl) {
          previewUrl = fileUrl.embedUrl;
        }
      } catch (e) {
        console.error("Error in uploadFileToDrive: ", e);
      }
    } else {
      shouldGeneratePDF = true;
    }

    // create payload
    const interviewId = Utilities.getUuid();

    const interviewDate = normalizeDate(new Date(slot.Date)).toLocaleDateString(
      "th-Th",
    );

    const startTime = Utilities.formatDate(
      new Date(slot.Start_Time),
      Session.getScriptTimeZone(),
      "HH:mm",
    );
    const endTime = Utilities.formatDate(
      new Date(slot.End_Time),
      Session.getScriptTimeZone(),
      "HH:mm",
    );

    const timePeriod = `${startTime}-${endTime}`;
    const buName = `'${slot.Bu_Name}`;
    const emailValue = String(email).toLowerCase();
    const formattedBirthDate = new Date(birthdate).toLocaleDateString("th-Th");

    const formattedAvailableStartDate = availableStartDate
      ? new Date(availableStartDate).toLocaleDateString("th-TH")
      : "";

    // Create new interview initial row (without calendar_id first)
    const newInterview = [
      interviewId,
      slotId,
      candidateId,
      interviewDate,
      timePeriod,
      shop || "",
      buName,
      fullname,
      nickname,
      formattedBirthDate,
      age,
      gender,
      weight,
      height,
      nationality,
      nationalId,
      maritalStatus,
      militaryStatus,
      currentAddress,
      "'" + phoneNumber,
      emailValue,
      "'" + lineId,
      highestEducation,
      institution,
      faculty,
      major,
      workLocation,
      workPeriod,
      workPosition,
      workDescription,
      saleExperience,
      positionType,
      positionApplied,
      preferredProvince,
      preferredDistrict,
      expectedSalary,
      formattedAvailableStartDate,
      previewUrl,
      INTERVIEW_STATUS[0],
      "", // calendar_id placeholder
      formattedNowDate,
      "", // interview link
      PARTICIPATION_STATUS[2],
    ];

    interviewSheet.appendRow(newInterview);

    const capacityStore = Number(slot.Capacity);
    if (capacityStore > 0) {
      const rowIndex = findRowIndex(slotData, slotId, "Slot_Id");
      const capacity = Number(capacityStore) - 1;
      slotSheet.getRange(rowIndex + 2, 6).setValue(capacity);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "ทะเบียนสัมภาษณ์สําเร็จเสร็จสิ้น",
      pendingExtras: true,
      extrasPayload: {
        interviewId: interviewId,
        eventId: eventId,
        slotId: slotId,
        candidateId: candidateId,
        shouldGeneratePDF: shouldGeneratePDF,
      },
    };
  } catch (error) {
    console.error("Error in createInterview: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  } finally {
    lock.releaseLock();
  }
}

function processInterviewExtras(payload) {
  try {
    const { interviewId, eventId, slotId, candidateId, shouldGeneratePDF } =
      payload;

    const ss = SpreadsheetApp.openById(eventId);
    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    const slotSheet = ss.getSheetByName(EVENT_SHEETS.SLOTS);

    const interviewDataAll = getRowsFromSheet(interviewSheet);
    const interviewRowBase = findRowIndex(
      interviewDataAll,
      interviewId,
      "Interview_Id",
    );
    const actualRow = interviewRowBase !== -1 ? interviewRowBase + 2 : -1;

    if (actualRow === -1)
      return { success: false, message: "interview not found" };

    const interview = interviewDataAll[interviewRowBase];
    const slotData = getRowsFromSheet(slotSheet);
    const slot = slotData.find((s) => s.Slot_Id === slotId);

    let previewUrl = interview.Resume_Url;

    // 1. Generate PDF
    if (shouldGeneratePDF) {
      const candidateSheet = getCandidateSheet();
      const candidateData = getRowsFromSheet(candidateSheet);
      const candidate = candidateData.find(
        (c) => c.Candidate_Id === candidateId,
      );

      const dataForPdf = {
        profileImageUrl: candidate ? candidate.Profile_Url : "",
        formattedBirthDate: parseThaiDateTime(interview.Birthdate).toLocaleDateString(
          "th-TH",
          { day: "numeric", month: "long", year: "numeric" },
        ),
        formattedStartDate: parseThaiDateTime(
          interview.Available_Start_Date,
        ).toLocaleDateString("th-TH", {
          day: "numeric",
          month: "long",
          year: "numeric",
        }),
        fullname: interview.Fullname,
        nickname: interview.Nickname,
        age: interview.Age,
        gender: interview.Gender,
        weight: interview.Weight,
        height: interview.Height,
        nationality: interview.Nationality,
        nationalId: interview.National_Id,
        maritalStatus: interview.Marital_Status,
        militaryStatus: interview.Military_Status,
        currentAddress: interview.Current_Address,
        phoneNumber: "'" + String(interview.Phone_Number).replace(/'/g, ""),
        email: interview.Email,
        lineId: "'" + String(interview.Line_Id).replace(/'/g, ""),
        highestEducation: interview.Highest_Education,
        institution: interview.Institution,
        faculty: interview.Faculty,
        major: interview.Major,
        workLocation: interview.Work_Location,
        workPeriod: interview.Work_Period,
        workPosition: interview.Work_Position,
        workDescription: interview.Work_Description,
        saleExperience: interview.Sale_Experience,
        positionType: interview.Position_Type,
        positionApplied: interview.Position_Applied,
        preferredProvince: interview.Preferred_Province,
        preferredDistrict: interview.Preferred_District,
        expectedSalary: interview.Expected_Salary,
      };

      const fileUrl = createInterviewPDF(dataForPdf);
      if (fileUrl && fileUrl.success) {
        previewUrl = fileUrl.embedUrl;
        interviewSheet.getRange(actualRow, 38).setValue(previewUrl);
      }
    }

    // 2. Calendar and Email
    const eventSheet = getEventSheet();
    const eventData = getRowsFromSheet(eventSheet);
    const event = eventData.find((data) => data.Sheet_ID === eventId);

    if (event) {
      const emailDate = normalizeDate(new Date(slot.Date)).toLocaleDateString(
        "th-Th",
        {
          year: "numeric",
          month: "long",
          day: "numeric",
        },
      );

      const calendarPayload = {
        event,
        slot,
        eventName: event.Event_Name,
        email: interview.Email,
        fullname: interview.Fullname,
        buName: interview.Shop,
        location: event.Location,
      };

      const calendar_id = createCalendarEvent(calendarPayload);
      if (calendar_id) {
        interviewSheet.getRange(actualRow, 40).setValue(calendar_id);
      }

      const payloadEmail = {
        to: interview.Email,
        eventName: event.Event_Name,
        buName: interview.Shop,
        fullname: interview.Fullname,
        date: emailDate,
        timePeriod: interview.Interview_Period,
        type: event.Event_Type,
        location: event.Location,
      };
      sendEmail(payloadEmail);
    }

    return { success: true };
  } catch (e) {
    console.error("processInterviewExtras error: ", e);
    return { success: false };
  }
}

function getAllInterviews(query = {}, userRole) {
  try {
    const {
      search,
      candidateId,
      eventName,
      province,
      permissionType,
      page = 1,
      limit = 10,
    } = query;

    const pageNum = Math.max(1, parseInt(page));
    const limitNum = Math.max(1, parseInt(limit));

    if (!AUTH_ROLES.includes(userRole)) {
      return {
        success: false,
        message: "คุณไม่ได้รับอนุญาติให้อ่านข้อมูล",
      };
    }

    const isUser = userRole === USER_ROLE;
    const isManager = userRole === MANAGER_ROLE;

    if (isUser && !candidateId) {
      return {
        success: true,
        data: [],
        pagination: {
          totalItems: 0,
          totalPages: 0,
          currentPage: pageNum,
          limit: limitNum,
        },
      };
    }

    const eventSheet = getEventSheet();
    const eventData = getRowsFromSheet(eventSheet);

    // ถ้าผู้ใช้เจาะจง Event มาแล้ว (เลือกจาก dropdown) ให้ดึงเฉพาะ Event นั้น
    // ไม่ต้องดึงข้อมูลทุก Event มา filter ทีหลัง เพราะช้าและสิ้นเปลือง quota
    const hasEventNameFilter = !!eventName;

    // หมายเหตุ: เดิมมีเงื่อนไขจำกัดเฉพาะ Event ที่ยังไม่ปิดเกิน 7 วัน เพื่อลด
    // จำนวน Sheet ที่ต้องยิง fetch (ประหยัด UrlFetchApp quota) แต่ตามคำขอ
    // ให้แสดงข้อมูลผู้สมัครทั้งหมดโดยไม่กรองช่วงวันที่ จึงตัดเงื่อนไขนี้ออก
    // ข้อควรระวัง: ถ้ามี Event สะสมในระบบเยอะ การดึงทุก Event ทุกครั้งจะยิง
    // UrlFetchApp มากขึ้น มีโอกาสชน quota รายวันง่ายขึ้นกว่าเดิม (ตัวแคชที่
    // ทำไว้ก่อนหน้านี้จะช่วยลดผลกระทบส่วนนี้อยู่บ้าง)

    const validEvents = eventData.filter((event) => {
      if (permissionType && permissionType !== event.Event_Type) {
        return false;
      }
      if (isManager && String(event.Status).toUpperCase() !== "TRUE") {
        return false;
      }
      if (hasEventNameFilter) {
        return event.Event_Name === eventName;
      }
      return true;
    });

    if (validEvents.length === 0) {
      return {
        success: true,
        data: [],
        pagination: {
          totalItems: 0,
          totalPages: 0,
          currentPage: pageNum,
          limit: limitNum,
        },
      };
    }

    // --- OPTIMIZATION: PARALLEL HTTP REQUESTS + CACHE ---
    // แคชผลดิบ (ก่อน filter/sort/pagination) ตามชุด event ที่ต้องดึง
    // เพราะ filter/sort/pagination ทำในหน่วยความจำ ไม่เปลือง quota
    // แต่การยิง UrlFetchApp ไปยัง Sheets API ต่างหากที่ทำให้ quota
    // รายวันหมด (Service invoked too many times for one day: urlfetch)
    // เมื่อฟังก์ชันนี้ถูกเรียกซ้ำๆ ทุกครั้งที่เปลี่ยนหน้า/พิมพ์ค้นหา
    const rawCacheKeySeed =
      "raw_interviews_" + validEvents.map((e) => e.Sheet_ID).sort().join(",");
    // ตอนนี้ไม่กรองช่วงวันที่แล้ว จำนวน event ที่ดึงพร้อมกันอาจเยอะขึ้นมาก
    // ทำให้ cache key ยาวเกิน 250 ตัวอักษรที่ CacheService รองรับ จึงแปลง
    // เป็น hash สั้นๆ แทนถ้ายาวเกินไป
    const rawCacheKey =
      rawCacheKeySeed.length > 200
        ? "raw_interviews_" +
        Utilities.base64EncodeWebSafe(
          Utilities.computeDigest(
            Utilities.DigestAlgorithm.MD5,
            rawCacheKeySeed,
          ),
        )
        : rawCacheKeySeed;
    const cache = CacheService.getScriptCache();
    let allInterviews;
    let cachedRaw = null;
    try {
      // ใช้ getCacheChunked แทน cache.get ตรงๆ เพราะข้อมูล raw_interviews
      // ของบาง event set มีขนาดใหญ่กว่า 100KB ที่ CacheService รองรับต่อ key
      // เดิมเขียนด้วย cache.put() ตรงๆ จึง throw "Argument too large: value"
      // เงียบๆ (ถูกจับด้วย try/catch) ทำให้ไม่เคยแคชสำเร็จเลย ต้องยิง
      // Sheets API ใหม่ทุกครั้งที่มีคนค้นหา/เปลี่ยนหน้า
      cachedRaw = getCacheChunked(cache, rawCacheKey);
    } catch (e) {
      console.warn("Cannot read raw_interviews cache: " + e.message);
    }

    if (cachedRaw) {
      allInterviews = JSON.parse(cachedRaw);
    } else {
      // แก้ไข: เดิมยิง UrlFetchApp.fetchAll() ตรงไปที่ Sheets REST API ทำให้
      // ชนโควต้ารายวันของ UrlFetchApp ("Service invoked too many times for
      // one day: urlfetch") ซึ่งเป็นสาเหตุที่หน้าฐานข้อมูลผู้สมัครดึงข้อมูล
      // ไม่ขึ้น เปลี่ยนมาใช้ Advanced Sheets Service (Sheets.Spreadsheets.
      // Values.batchGet) แทน เพราะใช้โควต้าคนละก้อนกับ UrlFetchApp (เหมือนที่
      // getAllInterviewsWithoutInterviewLink ใช้อยู่แล้วและทำงานได้ปกติ)
      const ranges = [
        EVENT_SHEETS.INTERVIEWS,
        EVENT_SHEETS.SLOTS,
        EVENT_SHEETS.TRANSACTION,
      ];

      const mapValuesToRows = (values) => {
        if (!values || values.length < 2) return [];
        const header = values[0];
        const rows = [];
        for (let i = 1; i < values.length; i++) {
          const rowData = values[i];
          if (rowData.join("").trim() === "") continue;
          const rowObj = {};
          header.forEach((h, colIndex) => {
            rowObj[h] =
              rowData[colIndex] !== undefined ? rowData[colIndex] : "";
          });
          rows.push(rowObj);
        }
        return rows;
      };

      allInterviews = [];
      let hasFailedRequest = false;

      validEvents.forEach((event) => {
        try {
          const response = callSheetsApiWithBackoff(() =>
            Sheets.Spreadsheets.Values.batchGet(event.Sheet_ID, {
              ranges: ranges,
              valueRenderOption: "FORMATTED_VALUE",
            }),
          );

          if (!response.valueRanges || response.valueRanges.length < 3) {
            return;
          }

          const interviewData = mapValuesToRows(
            response.valueRanges[0].values,
          );
          const slotData = mapValuesToRows(response.valueRanges[1].values);
          const transactionData = mapValuesToRows(
            response.valueRanges[2].values,
          );

          const slotMap = {};
          for (let i = 0; i < slotData.length; i++) {
            slotMap[slotData[i].Slot_Id] = slotData[i];
          }

          const transactionMap = {};
          for (let i = 0; i < transactionData.length; i++) {
            const data = transactionData[i];
            if (!transactionMap[data.Interview_Id]) {
              transactionMap[data.Interview_Id] = [];
            }
            transactionMap[data.Interview_Id].push(data);
          }

          for (let i = 0; i < interviewData.length; i++) {
            const interview = interviewData[i];
            allInterviews.push({
              ...interview,
              eventDetail: event,
              slotDetail: slotMap[interview.Slot_Id] || null,
              transactionDetail: transactionMap[interview.Interview_Id] || [],
            });
          }
        } catch (e) {
          hasFailedRequest = true;
          console.warn(
            `Sheets API request failed for event ${event && event.Event_Name}: ${e.message}`,
          );
        }
      });

      // ถ้ามี request ไหนล้มเหลว (เช่น ตอน quota เต็ม หรือ Sheets API
      // ตอบ error ชั่วคราว) ห้ามแคชผลลัพธ์นี้ไว้ ไม่งั้นหน้าเว็บจะโชว์
      // "ไม่พบข้อมูล" ต่อไปอีกหลายนาทีทั้งที่จริงๆ ข้อมูลมีอยู่
      if (!hasFailedRequest) {
        try {
          // เพิ่มเวลาแคชจาก 2 นาที เป็น 20 นาที เพื่อลดจำนวนครั้งที่ต้องยิง
          // ไปอ่าน Sheet จริง (หน้านี้ถูกเรียกทุกครั้งที่พิมพ์ค้นหา/เปลี่ยนหน้า)
          // ใช้ putCacheChunked แทน cache.put() ตรงๆ เพราะข้อมูลรวมทุก event
          // มักมีขนาดเกิน 100KB (ลิมิตต่อ key ของ CacheService) การ put ตรงๆ
          // จะ throw "Argument too large: value" เงียบๆ ทำให้ไม่เคยแคชสำเร็จ
          putCacheChunked(cache, rawCacheKey, JSON.stringify(allInterviews), 1200); // 20 นาที
        } catch (e) {
          console.warn("Cannot write raw_interviews cache: " + e.message);
        }
      }
    }

    const filteredInterviews = allInterviews.filter((interview) => {
      if (isUser && candidateId !== String(interview.Candidate_Id))
        return false;
      if (isManager && interview.Status === "ไม่ผ่านการพิจารณา") return false;
      if (eventName && eventName !== interview.eventDetail.Event_Name)
        return false;
      if (province && province !== interview.Preferred_Province) return false;

      const searchValue = search
        ? String(search).toLowerCase().replace(/\s+/g, "")
        : "";

      if (search && searchValue !== "") {
        const matchesSearch = Object.entries(interview).some(([key, value]) => {
          if (
            typeof value === "object" ||
            value === null ||
            value === undefined
          )
            return false;

          const cleanValue = String(value).toLowerCase().replace(/\s+/g, "");
          return cleanValue.includes(searchValue);
        });

        if (!matchesSearch) return false;
      }
      return true;
    });

    // --- SORTING LOGIC
    filteredInterviews.sort((a, b) => {
      const dateA = parseThaiDateTime(a.TimeStamp);
      const dateB = parseThaiDateTime(b.TimeStamp);
      return dateB.getTime() - dateA.getTime();
    });

    // --- PAGINATION LOGIC
    const totalItems = filteredInterviews.length;
    const totalPages = Math.ceil(totalItems / limitNum);
    const startIndex = (pageNum - 1) * limitNum;
    const endIndex = startIndex + limitNum;
    const paginatedData = filteredInterviews.slice(startIndex, endIndex);

    return {
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
  } catch (error) {
    console.error("Error in getAllInterviews: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function getAllInterviewsWithoutInterviewLink(query = {}) {
  try {
    const { eventName, buName, page = 1, limit = 10 } = query;

    if (!eventName || !buName) {
      return {
        success: true,
        data: [],
        pagination: { totalItems: 0, totalPages: 0, currentPage: 1, limit: limit },
      };
    }

    const pageNum = Math.max(1, parseInt(page));
    const limitNum = Math.max(1, parseInt(limit));

    // === จุดที่แก้: cache key ผูกกับ eventName เท่านั้น (ไม่ผูก buName/page เพราะ filter/pagination ทำในหน่วยความจำได้) ===
    const cache = CacheService.getScriptCache();
    const rawCacheKey = "raw_interviews_nolink_" + eventName;

    let interviewData, slotData, transactionData, eventData;
    let cachedRaw = null;

    try {
      cachedRaw = getCacheChunked(cache, rawCacheKey);
    } catch (e) {
      console.warn("Cannot read raw_interviews_nolink cache: " + e.message);
    }

    if (cachedRaw) {
      const parsed = JSON.parse(cachedRaw);
      interviewData = parsed.interviewData;
      slotData = parsed.slotData;
      transactionData = parsed.transactionData;
      eventData = parsed.eventData;
    } else {
      const responseEvents = callSheetsApiWithBackoff(() =>
        Sheets.Spreadsheets.Values.batchGet(MAIN_SS_ID, {
          ranges: [MAIN_SHEETS.EVENTS],
          valueRenderOption: "FORMATTED_VALUE",
        }),
      );

      const rawEventData =
        mergeMapValuesToRows(responseEvents.valueRanges[0].values) || [];
      eventData = rawEventData.find((event) => event.Event_Name === eventName);

      if (!eventData) {
        return { success: false, message: "ไม่พบข้อมูล Event: " + eventName };
      }

      const eventId = eventData.Sheet_ID;
      const ranges = [
        EVENT_SHEETS.INTERVIEWS,
        EVENT_SHEETS.SLOTS,
        EVENT_SHEETS.TRANSACTION,
      ];

      const response = callSheetsApiWithBackoff(() =>
        Sheets.Spreadsheets.Values.batchGet(eventId, {
          ranges: ranges,
          valueRenderOption: "FORMATTED_VALUE",
        }),
      );

      if (!response.valueRanges || response.valueRanges.length < 3) {
        return { success: false, message: "ไม่พบข้อมูล Event: " + eventName };
      }

      interviewData = mergeMapValuesToRows(response.valueRanges[0].values);
      slotData = mergeMapValuesToRows(response.valueRanges[1].values);
      transactionData = mergeMapValuesToRows(response.valueRanges[2].values);

      // === จุดที่แก้: เก็บ cache 5 นาที ===
      try {
        putCacheChunked(cache, rawCacheKey, JSON.stringify({
          interviewData, slotData, transactionData, eventData
        }), 300);
      } catch (e) {
        console.warn("Cannot write raw_interviews_nolink cache: " + e.message);
      }
    }

    const formatBuName = (value = "") => value.trim().toLowerCase();

    const filteredInterviews = interviewData.filter((interview) => {
      if (buName && formatBuName(buName) !== formatBuName(interview.Bu_Name))
        return false;
      if (interview.Status !== "รอการพิจารณา") return false;
      if (!interview.Calendar_Id) return false;
      if (interview.Interview_Link) return false;
      return true;
    });

    filteredInterviews.sort((a, b) => {
      const dateA = parseThaiDateTime(a.TimeStamp);
      const dateB = parseThaiDateTime(b.TimeStamp);
      return dateB.getTime() - dateA.getTime();
    });

    const totalItems = filteredInterviews.length;
    const totalPages = Math.ceil(totalItems / limitNum);
    const startIndex = (pageNum - 1) * limitNum;
    const endIndex = startIndex + limitNum;

    const slicedData = filteredInterviews.slice(startIndex, endIndex);

    const paginatedData = slicedData.map((item) => ({
      ...item,
      eventDetail: eventData,
      slotDetail: slotData.find((slot) => slot.Slot_Id === item.Slot_Id) || null,
      transactionDetail: transactionData.filter(
        (transaction) => transaction.Interview_Id === item.Interview_Id
      ),
    }));

    return {
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
  } catch (error) {
    console.error("Error in getAllInterviews: ", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด" };
  }
}

function getInterviewsById(getData) {
  try {
    if (!getData) {
      return {
        success: false,
        message: "ไม่พบข้อมูลนี้",
      };
    }

    const { eventId, interviewId, candidateId, userRole } = getData;

    if (!AUTH_ROLES.includes(userRole)) {
      return {
        success: false,
        message: "คุณไม่ได้รับอนุญาติให้อ่านข้อมูล",
      };
    }

    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) {
      return {
        success: false,
        message: "ไม่พบชีทนี้",
      };
    }

    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    if (!interviewSheet) {
      return {
        success: false,
        message: "ไม่พบชีทนี้",
      };
    }

    const rawData = getRowsFromSheet(interviewSheet);
    let dataObj = {};

    if (userRole === USER_ROLE) {
      const isYourInterview = rawData.find((row) => {
        return (
          row.Interview_Id === interviewId && row.Candidate_Id === candidateId
        );
      });

      if (!isYourInterview) {
        return {
          success: false,
          message: "คุณไม่ได้รับอนุญาติให้อ่านข้อมูล",
        };
      }
      dataObj = isYourInterview;
    } else {
      const filteredData = rawData.filter((row) => {
        return row.Interview_Id === interviewId;
      });
      dataObj = filteredData[0] || {};   // ← เพิ่ม [0] ดึง object ตัวแรกออกจาก array
    }

    return {
      success: true,
      data: dataObj,
    };
  } catch (error) {
    console.error("Error in getAllInterviews: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function updateDetailInterview(updateData) {
  try {
    if (!updateData) {
      return {
        success: false,
        message: "ไม่พบข้อมูลนี้",
      };
    }

    const { userRole, userId, eventId, interviewId, file } = updateData;

    if (userRole !== USER_ROLE) {
      return {
        success: false,
        message: "คุณไม่ได้รับอนุญาติให้แก้ข้อมูล",
      };
    }

    const ss = SpreadsheetApp.openById(eventId);
    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);

    const rowNum = findRowIndex(interviewSheet, interviewId, "Interview_Id");
    const actualRow = rowNum + 2;

    const interviewRange = interviewSheet.getRange(
      actualRow,
      1,
      1,
      interviewSheet.getLastColumn(),
    );
    const interview = interviewRange.getValues()[0];

    const index = {
      slotId: 1,
      candidateId: 2,
      shop: 5,
      buName: 6,
      nickname: 8,
      birthdate: 9,
      age: 10,
      gender: 11,
      weight: 12,
      height: 13,
      nationality: 14,
      maritalStatus: 16,
      militaryStatus: 17,
      currentAddress: 18,
      highestEducation: 22,
      institution: 23,
      faculty: 24,
      major: 25,
      workLocation: 26,
      workPeriod: 27,
      workPosition: 28,
      workDescription: 29,
      saleExperience: 30,
      positionType: 31,
      positionApplied: 32,
      preferredProvince: 33,
      preferredDistrict: 34,
      expectedSalary: 35,
      availableStartDate: 36,
    };

    if (!interview) {
      return {
        success: false,
        message: "ไม่พบข้อมูลนี้",
      };
    }

    if (userId !== interview[index.candidateId]) {
      return {
        success: false,
        message: "คุณไม่ได้รับอนุญาติให้แก้ไขข้อมูล",
      };
    }

    const now = new Date();
    const actualDate = normalizeDate(interview[3]);
    actualDate.setDate(actualDate.getDate() - 3);

    if (now > actualDate) {
      return {
        success: false,
        message: "หมดเขตการแก้ไชข้อมูลแล้ว",
      };
    }

    const storeData = {};
    const newData = {};
    let modified = false;

    if (file) {
      if (interview[37]) {
        deleteFileByUrl(interview[37]);
      }

      const fileUrl = uploadFileToDrive(
        file.base64Data,
        file.fileName,
        file.mimeType,
      );

      if (fileUrl) {
        interview[37] = fileUrl.embedUrl;
        modified = true;
      }
    }

    if (updateData.birthdate) {
      updateData.birthdate = new Date(updateData.birthdate).toLocaleDateString(
        "th-TH",
      );
    }

    if (updateData.availableStartDate) {
      updateData.availableStartDate = new Date(
        updateData.availableStartDate,
      ).toLocaleDateString("th-TH");
    }

    Object.keys(updateData).forEach((key) => {
      if (index[key] !== undefined) {
        if (updateData[key] !== interview[index[key]]) {
          console.log(`➡️ Updating field: ${key}`);
          storeData[key] = interview[index[key]];
          newData[key] = updateData[key];
          console.log(`📝 ${key}: ${storeData[key]} -> ${newData[key]}`);

          interview[index[key]] = updateData[key];
          modified = true;
        }
      }
    });

    if (interview[19]) {
      const rawPhone = String(interview[19]).replace(/^'+/, "");
      interview[19] = `'${rawPhone}`;
    }

    if (!modified) {
      return {
        success: false,
        message: "ไม่มีการแก้ไขข้อมูล",
      };
    }

    interviewRange.setValues([interview]);

    return {
      success: true,
      message: "แก้ไขข้อมูลสําเร็จ",
    };
  } catch (error) {
    console.error("Error in updateDetailInterview: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function updateSlotInterview(updateData) {
  try {
    if (!updateData) {
      return {
        success: false,
        message: "กรุณากรอกข้อมูลให้ครบถ้วน",
      };
    }

    const { eventId, interviewId, slotId, shop } = updateData;
    const mainSheet = SpreadsheetApp.openById(MAIN_SS_ID);

    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) {
      return { success: false, message: "ไม่พบชีทนี้" };
    }

    const eventSheet = mainSheet.getSheetByName(MAIN_SHEETS.EVENTS);
    const eventRows = getRowsFromSheet(eventSheet);
    const event = eventRows.find((data) => data.Sheet_ID === eventId);

    const eventName = event.Event_Name;
    const location = event.Location;
    const type = event.Event_Type;

    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    if (!interviewSheet) {
      return { success: false, message: "ไม่พบชีทนี้" };
    }

    const slotSheet = ss.getSheetByName(EVENT_SHEETS.SLOTS);
    const slotData = getRowsFromSheet(slotSheet);
    const slot = slotData.find((slot) => slot.Slot_Id === slotId);
    if (!slot) {
      return { success: false, message: "สล็อคนี้ยังไม่เปิดให้สัมภาษณ์" };
    }

    const now = new Date();
    const slotDate = Utilities.formatDate(
      normalizeDate(slot.Date),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    const targetDate = combineDateTime(
      slotDate,
      formatTimeTh(new Date(slot.End_Time)),
    );

    if (now > targetDate) {
      return { success: false, message: "ช่วงเวลนี้ปิดแล้ว" };
    }

    const rowNum = findRowIndex(interviewSheet, interviewId, "Interview_Id");
    if (rowNum === -1) {
      return { success: false, message: "ไม่พบข้อมูลนี้ในระบบ" };
    }
    const actualRow = rowNum + 2;

    const interviewRange = interviewSheet.getRange(
      actualRow,
      1,
      1,
      interviewSheet.getLastColumn(),
    );
    const interview = interviewRange.getValues()[0];

    const index = {
      slotId: 1,
      interviewDate: 3,
      interviewPeriod: 4,
      shop: 5,
      buName: 6,
    };

    const storeSlotId = interview[index.slotId];

    if (storeSlotId === slotId) {
      return { success: false, message: "คุณได้ลงทะเบียนแล้ว" };
    }

    const interviewDate = normalizeDate(new Date(slot.Date)).toLocaleDateString(
      "th-Th",
    );

    const startTime = Utilities.formatDate(
      new Date(slot.Start_Time),
      Session.getScriptTimeZone(),
      "HH:mm",
    );
    const endTime = Utilities.formatDate(
      new Date(slot.End_Time),
      Session.getScriptTimeZone(),
      "HH:mm",
    );
    const timePeriod = `${startTime}-${endTime}`;

    const newData = {
      slotId: slotId,
      interviewDate: interviewDate,
      interviewPeriod: timePeriod,
      buName: slot.Bu_Name,
    };

    if (typeof shop === "string") {
      newData.shop = shop;
    }

    let modified = false;
    Object.keys(newData).forEach((key) => {
      if (index[key] !== undefined) {
        const colIndex = index[key];
        const oldValue = interview[colIndex];
        const newValue = newData[key];

        if (String(oldValue) !== String(newValue)) {
          console.log(`➡️ Updating field: ${key}`);
          console.log(`📝 Change: ${oldValue} -> ${newValue}`);
          interview[colIndex] = newValue;
          modified = true;
        }
      }
    });

    if (!modified) {
      return { success: false, message: "ไม่มีการแก้ไขข้อมูล" };
    }

    if (interview[19]) {
      const rawPhone = String(interview[19]).replace(/^'+/, "");
      interview[19] = `'${rawPhone}`;
    }

    // update old slot capacity
    const oldSlotRowIndex = findRowIndex(slotSheet, storeSlotId, "Slot_Id");
    const oldSlotRange = slotSheet.getRange(
      oldSlotRowIndex + 2,
      1,
      1,
      slotSheet.getLastColumn(),
    );
    const oldSlotData = oldSlotRange.getValues()[0];
    oldSlotData[5] = Number(oldSlotData[5]) + 1;
    oldSlotData[4] = "'" + String(oldSlotData[4]);
    oldSlotRange.setValues([oldSlotData]);

    // update new slot capacity
    const newSlotRowIndex = findRowIndex(slotSheet, slotId, "Slot_Id");
    const newSlotRange = slotSheet.getRange(
      newSlotRowIndex + 2,
      1,
      1,
      slotSheet.getLastColumn(),
    );
    const newSlotData = newSlotRange.getValues()[0];
    newSlotData[5] = Math.max(0, Number(newSlotData[5]) - 1);
    newSlotData[4] = "'" + String(newSlotData[4]);
    newSlotRange.setValues([newSlotData]);

    const email = interview[20];
    const fullname = interview[7];
    const calendarId = interview[39];
    const emailDate = normalizeDate(new Date(slot.Date)).toLocaleDateString(
      "th-Th",
      { year: "numeric", month: "long", day: "numeric" },
    );

    const updateCalandarPayload = {
      calendarId,
      slot,
      eventName,
      email,
      fullname,
      buName: interview.Shop,
      location,
    };

    updateCalendarEvent(updateCalandarPayload);

    const sendMailPayload = {
      to: email,
      eventName,
      buName: interview.Shop,
      fullname,
      date: emailDate,
      timePeriod,
      type,
      location,
    };

    sendEmail(sendMailPayload);
    interviewRange.setValues([interview]);

    return {
      success: true,
      message: "เปลี่ยนเวลาสัมภาษณ์เสร็จสิ้น",
    };
  } catch (error) {
    console.error("Error in updateSlotInterview: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function updateStatusInterview(updateData) {
  try {
    const {
      eventId,
      interviewId,
      candidateName,
      userName,
      fullName,
      userRole,
      status,
      result_bu,
      result_store_id,
      result_position,
      description,
    } = updateData;

    if (userRole === USER_ROLE) {
      return {
        success: false,
        message: "คุณไม่ได้รับอนุญาติให้อ่านข้อมูล",
      };
    }

    const mainSS = getSpreadsheet();
    const eventSheet = mainSS.getSheetByName(MAIN_SHEETS.EVENTS);
    const eventData = getRowsFromSheet(eventSheet);

    const event = eventData.find((item) => eventId === item["Sheet_ID"]);
    const event_job_type = event["Job_Type"];

    const STORE_JOB_TYPES = [
      "พนักงานประจำหน้าร้านสาขา True",
      "พนักงานประจำหน้าร้านสาขา Studio7",
      "พนักงานประจำหน้าร้านสาขา BaNaNa",
      "พนักงานประจำหน้าร้านสาขา Dtac",
    ];

    const isStore = STORE_JOB_TYPES.includes(event_job_type);

    if (!INTERVIEW_STATUS.includes(status)) {
      return {
        success: false,
        message: "สถานะไม่ถูกต้อง",
      };
    }

    if (status === INTERVIEW_STATUS[3] && !description) {
      return {
        success: false,
        message: "กรุณากรอกเหตุผล",
      };
    }

    if (status === INTERVIEW_STATUS[1] && isStore) {
      if (!result_bu) return { success: false, message: "กรุณากรอกหน่วยงาน" };
      if (!result_store_id) return { success: false, message: "กรุณากรอกสาขา" };
      if (!result_position) return { success: false, message: "กรุณากรอกตำแหน่ง" };
    }

    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) return { success: false, message: "ไม่พบชีทนี้" };

    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    if (!interviewSheet) return { success: false, message: "ไม่พบชีทนี้" };

    const index = findRowIndex(interviewSheet, interviewId, "Interview_Id");
    const actualIndex = index + 2;
    if (index === -1) return { success: false, message: "ไม่พบข้อมูลนี้" };

    const row = interviewSheet
      .getRange(actualIndex, 1, 1, interviewSheet.getLastColumn())
      .getValues()[0];

    interviewSheet.getRange(actualIndex, 39).setValue(status);

    if (status === INTERVIEW_STATUS[1]) {
      interviewSheet.getRange(actualIndex, 44).setValue(result_bu);
      interviewSheet.getRange(actualIndex, 45).setValue(result_store_id);
      interviewSheet.getRange(actualIndex, 46).setValue(result_position);
    }

    const interviewName = row[7];
    const transactionSheet = ss.getSheetByName(EVENT_SHEETS.TRANSACTION);
    const now = new Date().toLocaleString("th-TH");

    const newHistory = [
      now,
      interviewId,
      "'" + userName,
      fullName,
      candidateName,
      status,
      description,
    ];

    transactionSheet.appendRow(newHistory);

    const responseMessage = `คุณได้บันทึกข้อมูลสําเร็จเสร็จสิ้น ชื่อผู้สมัคร: ${interviewName} สถานะ: ${status} ${status === INTERVIEW_STATUS[3] ? `หมายุเหตุ: ${description}` : ""}`;

    return {
      success: true,
      message: responseMessage,
    };
  } catch (error) {
    console.error("Error in updateStatusInterview: ", error);
    return {
      success: false,
      message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด",
    };
  }
}

function updateParticipationInterview(updateData) {
  try {
    const {
      eventId,
      interviewId,
      candidateName,
      userName,
      fullName,
      userRole,
      status,
    } = updateData;

    if (!PARTICIPATION_STATUS.includes(status)) {
      return { success: false, message: "สถานะไม่ถูกต้อง" };
    }

    const ss = SpreadsheetApp.openById(eventId);
    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    const transactionSheet = ss.getSheetByName(EVENT_SHEETS.TRANSACTION);

    const interviewDataAll = getRowsFromSheet(interviewSheet);
    const rowIndex = findRowIndex(interviewDataAll, interviewId, "Interview_Id");
    if (rowIndex === -1) return { success: false, message: "ไม่พบข้อมูล" };

    const actualRow = rowIndex + 2;
    const interview = interviewDataAll[rowIndex];

    // อัพเดทสถานะเช็คชื่อหน้าสาขาในชีทหลัก
    interviewSheet.getRange(actualRow, 43).setValue(status);

    const now = new Date(); // ใช้ Date object ตรงๆ (ค.ศ.) ห้ามใช้ toLocaleString("th-TH") เพราะจะได้ปี พ.ศ. เป็นสตริง
    // แล้ว Sheets จะ auto-parse ปี พ.ศ. เป็นปีจริง ทำให้ formatJoinedAtForDisplay +543 ซ้ำ (บั๊ก 3112)
    const eventResult = getEventById(eventId);
    const event = (eventResult && eventResult.success && eventResult.data)
      ? eventResult.data.eventData
      : null;

    let result;
    let emailResult = { success: false, message: "ไม่มีการส่งอีเมล" };

    if (status === "เข้าร่วม" || status === "Walk-In") {
      // แก้ไข: เดิมใช้ interview.Employee_Category (ว่างเปล่าจนกว่าแอดมินจะสรุปผล
      // หลังสัมภาษณ์) หรือ interview.Position_Type (ค่าเดียวกันทั้ง Event มาจาก
      // eventJobType ไม่ใช่ตำแหน่งที่ผู้สมัครเลือกจริง) ทำให้เลขคิวแยกกลุ่มผิด —
      // เปลี่ยนมาใช้ interview.Position_Applied ซึ่งเป็นค่าที่ผู้สมัครกรอกจริง
      // ("พนักงานประจำหน้าร้านสาขา True/Studio7/BaNaNa/Dtac", "พนักงานสำนักงานใหญ่",
      // "นักศึกษาฝึกงาน") แล้วแปลงผ่าน getEmployeeCategoryFromJobType() ให้เหลือ
      // แค่ 2 กลุ่มสำหรับนับเลขคิว: "พนักงานหน้าร้าน" (รวมทุกสาขา นับไล่เลขต่อกัน
      // ในวันเดียวกัน) กับ "ออฟฟิศ" (พนักงานสำนักงานใหญ่ + นักศึกษาฝึกงาน นับรวม
      // อยู่ในอีกคิวหนึ่ง แยกจากคิวหน้าร้านโดยสิ้นเชิง) ถ้าข้อความไม่เข้าเงื่อนไขใดเลย
      // (กรณีข้อมูลเก่า/ผิดปกติ) จะ fallback ไปใช้ค่าดิบเหมือนพฤติกรรมเดิม
      const rawPositionText =
        interview.Position_Applied || interview.Employee_Category || interview.Position_Type || "";
      const employeeCategory = getEmployeeCategoryFromJobType(rawPositionText) || rawPositionText;

      // เรียกข้ามไฟล์ไปยังคิวเซอร์วิสที่แยกไว้เพื่อบันทึกและรันคิว
      result = appendQueueLog(
        event ? event.Event_Name : "",
        candidateName,
        now,
        interview.Interview_Period || "",
        interview.Status || "รอการพิจารณา",
        fullName,
        status,
        interviewId,
        employeeCategory
      );

      // ส่งอีเมลส่งคิวสัมภาษณ์
      emailResult = sendCheckInEmail({
        fullname: candidateName,
        email: interview.Email || "",
        eventName: event ? event.Event_Name : "",
        queueNo: result.queueNo,
        date: formatThaiLongDate(interview.Interview_Date),
        time: interview.Interview_Period,
        positionApplied: interview.Position_Applied || "",
        employeeCategory: employeeCategory
      });

      // ปรับปรุงผลล็อกการส่งอีเมลลง Queue Sheet ข้ามไฟล์
      setQueueLogEmailStatus(result.rowNum, emailResult && emailResult.success);
    }

    transactionSheet.appendRow([now, interviewId, "'" + userName, fullName, candidateName, status]);

    const checkedInCategory = result && result.employeeCategory ? result.employeeCategory : "";
    const categorySuffix = checkedInCategory ? ` (ตำแหน่งที่สมัคร: ${checkedInCategory})` : "";

    return {
      success: true,
      message: `คิวสัมภาษณ์ลำดับที่ ${result?.queueNo || '-'}${categorySuffix}`,
      queueNo: result?.queueNo,
      employeeCategory: checkedInCategory,
      emailSent: !!(emailResult && emailResult.success),
      emailMessage: emailResult ? emailResult.message : "ไม่มีการส่งอีเมล"
    };
  } catch (error) {
    console.error("Error in updateParticipationInterview: ", error);
    return { success: false, message: "เกิดข้อผิดพลาด" };
  }
}

function deleteInterview(deleteData) {
  try {
    const { userRole, userName, userId, eventId, interviewIds, description } =
      deleteData;

    const isUser = userRole === USER_ROLE;
    const targetIds = Array.isArray(interviewIds) ? interviewIds : [interviewIds];

    if (targetIds.length === 0 || !eventId) {
      return {
        success: false,
        message: "กรุณากรอกข้อมูลให้ครบถ้วน",
      };
    }

    if (isUser && !description) {
      return {
        success: false,
        message: "กรุณากรอกเหตุผล",
      };
    }

    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) return { success: false, message: "ไม่พบชีทนี้" };

    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    const slotSheet = ss.getSheetByName(EVENT_SHEETS.SLOTS);
    const historySheet = ss.getSheetByName(EVENT_SHEETS.HISTORY);
    const transactionSheet = ss.getSheetByName(EVENT_SHEETS.TRANSACTION);

    if (!interviewSheet || !slotSheet || !historySheet) {
      return { success: false, message: "ไม่พบชีทที่ต้องการ" };
    }

    const interviewData = getRowsFromSheet(interviewSheet);
    const slotData = getRowsFromSheet(slotSheet);

    const interviewIndexMap = {};
    interviewData.forEach((row, i) => { interviewIndexMap[row.Interview_Id] = i; });

    const slotIndexMap = {};
    slotData.forEach((row, i) => { slotIndexMap[row.Slot_Id] = i; });

    const rowsToDelete = [];
    const rowsToChangeStatus = [];
    const logsToAppend = [];

    for (const id of targetIds) {
      const interview = interviewData[interviewIndexMap[id]];
      if (!interview) continue;

      const candidateName = interview.Fullname;
      const candidateEmail = interview.Email;
      const candidateCalendarId = interview.Calendar_Id;

      if (isUser && interview.Candidate_Id !== userId) {
        return {
          success: false,
          message: `คุณไม่ได้รับอนุญาตให้ลบข้อมูล ID: ${id}`,
        };
      }

      const slotId = interview.Slot_Id;
      const slotIndex = slotIndexMap[slotId];

      if (slotIndex !== undefined) {
        const slotActualIndex = slotIndex + 2;
        const currentCapacity = Number(slotData[slotIndex]["Capacity"]);
        slotSheet.getRange(slotActualIndex, 6).setValue(currentCapacity + 1);
        slotData[slotIndex]["Capacity"] = currentCapacity + 1;
      }

      const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
      let message = "";

      if (isUser) {
        message = `${userName} ได้ทำการยกเลิกการสมัครสัมภาษณ์เนื่องจาก ${description}`;
        logsToAppend.push([now, userName, message]);
      } else {
        message = `${userName} ได้ทำการแก้ไขสถานะสมัครสัมภาษณ์เนื่องจาก ${description}`;
      }

      const rowIndex = interviewIndexMap[id];
      if (rowIndex !== undefined) {
        const actualRow = rowIndex + 2;
        if (isUser) {
          rowsToDelete.push(actualRow);
        } else {
          rowsToChangeStatus.push(actualRow);
        }
      }

      if (candidateCalendarId && candidateCalendarId !== "") {
        deleteCalendarEvent(candidateCalendarId);
      }

      if (!isUser) {
        const payload = { to: candidateEmail, fullname: candidateName };
        sendEmail(payload);

        const taSheet = getTASheet();
        const taData = getRowsFromSheet(taSheet);
        const taUser = taData.find((user) => user.Name === userName);

        const nowTx = new Date().toLocaleString("th-TH");
        const status = INTERVIEW_STATUS[4];

        const newHistory = [
          nowTx,
          interview.Interview_Id,
          taUser.Username,
          userName,
          candidateName,
          status,
          description,
        ];
        transactionSheet.appendRow(newHistory);
      }
    }

    if (logsToAppend.length > 0) {
      logsToAppend.forEach((log) => historySheet.appendRow(log));
    }

    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach((row) => interviewSheet.deleteRow(row));
    }

    if (rowsToChangeStatus.length > 0) {
      rowsToChangeStatus.sort((a, b) => b - a);
      rowsToChangeStatus.forEach((row) => {
        interviewSheet.getRange(row, 39).setValue(INTERVIEW_STATUS[4]);
      });
    }

    return {
      success: true,
      message: isUser ? "ยกเลิกการสมัครเสร็จสิ้น" : "อัปเดตสถานะเป็นไม่ผ่านการพิจารณาเสร็จสิ้น",
    };
  } catch (error) {
    console.error("Error in deleteInterview: ", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด" };
  }
}

function addLinktoSheet(data) {
  try {
    if (!data) return { success: false, message: "ไม่พบข้อมูลนี้" };

    const { eventName, interviewIds, interviewLink } = data;

    if (!eventName || !interviewIds || interviewIds.length === 0 || !interviewLink) {
      return { success: false, message: "ข้อมูลไม่ครบ" };
    }

    const targetIds = Array.isArray(interviewIds) ? interviewIds : [interviewIds];
    const mainSS = SpreadsheetApp.openById(MAIN_SS_ID);
    const eventSheet = mainSS.getSheetByName(MAIN_SHEETS.EVENTS);
    const eventData = getRowsFromSheet(eventSheet).find((event) => event.Event_Name === eventName);

    if (!eventData) return { success: false, message: "ไม่พบ Event Name นี้ในระบบ" };

    const eventId = eventData["Sheet_ID"];
    const ss = SpreadsheetApp.openById(eventId);
    if (!ss) return { success: false, message: "ไม่พบชีทนี้" };

    const interviewSheet = ss.getSheetByName(EVENT_SHEETS.INTERVIEWS);
    if (!interviewSheet) return { success: false, message: "ไม่พบชีทนี้" };

    const lastRow = interviewSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "ไม่มีข้อมูลในชีท" };

    const headers = interviewSheet.getRange(1, 1, 1, interviewSheet.getLastColumn()).getValues()[0];

    const interviewIdIndex = headers.indexOf("Interview_Id");
    const interviewLinkIndex = headers.indexOf("Interview_Link");
    const emailIndex = headers.indexOf("Email");
    const statusIndex = headers.indexOf("Status");
    const calendarIdIndex = headers.indexOf("Calendar_Id");
    const interviewDateIndex = headers.indexOf("Interview_Date");

    if (interviewIdIndex === -1 || interviewLinkIndex === -1 || emailIndex === -1 || statusIndex === -1 || interviewDateIndex === -1) {
      return { success: false, message: "ไม่พบโครงสร้างคอลัมน์ที่จำเป็น" };
    }

    // === จุดที่แก้: อ่านข้อมูลทั้งตารางครั้งเดียว แทนการ getRange/getValues แยก 6 ครั้ง ===
    const allValues = interviewSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

    let updateCount = 0;

    for (let i = 0; i < allValues.length; i++) {
      const row = allValues[i];
      const currentId = row[interviewIdIndex];
      const currentStatus = String(row[statusIndex]).trim();
      const currentCalendarId = row[calendarIdIndex];
      const currentLink = row[interviewLinkIndex];

      if (
        targetIds.includes(currentId) &&
        currentStatus === "รอการพิจารณา" &&
        currentCalendarId !== "" &&
        currentLink === ""
      ) {
        row[interviewLinkIndex] = interviewLink;
        updateCount++;

        const recipientEmail = row[emailIndex];
        const formattedDate = normalizeDate(row[interviewDateIndex]).toLocaleDateString("th-TH", {
          year: "numeric", month: "long", day: "numeric"
        });

        console.log(`Sending mail to: ${recipientEmail} for ID: ${currentId}`);

        if (recipientEmail) {
          const payload = {
            to: recipientEmail,
            eventName,
            date: formattedDate,
            type: "CALL_LINK",
            interviewLink,
          };
          sendEmail(payload);
        }
      }
    }

    if (updateCount > 0) {
      // เขียนกลับทั้งตารางครั้งเดียว (เดิมเขียนเฉพาะคอลัมน์ link แยก ยังคงพฤติกรรมเดิม)
      interviewSheet.getRange(2, 1, allValues.length, headers.length).setValues(allValues);
      return {
        success: true,
        message: `เพิ่มลิงค์สมัครสัมภาษณ์เสร็จสิ้น (${updateCount} รายการ)`,
      };
    } else {
      return {
        success: false,
        message: "ไม่พบข้อมูลที่ตรงเงื่อนไข หรือมีลิงค์อยู่แล้ว",
      };
    }
  } catch (error) {
    console.error("Error in addLinktoSheet: ", error);
    return { success: false, message: "เซิร์ฟเวอร์เกิดข้อผิดพลาด" };
  }
}

// ==========================================================================
// INTERVIEW STANDALONE HELPERS (FOR CHECK-IN PROCESS)
// ==========================================================================

/**
 * แปลงวันที่สตริงให้อยู่ในรูปภาษาไทยแบบยาว (พ.ศ.) สำหรับการส่งเมลคิว
 */
function formatThaiLongDate(dateStr) {
  if (!dateStr) return "-";
  const thaiMonths = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];
  try {
    const s = String(dateStr).trim();
    const match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) {
      const day = parseInt(match[1], 10);
      const month = parseInt(match[2], 10);
      let year = parseInt(match[3], 10);
      if (year < 2400) year += 543;
      return `${day} ${thaiMonths[month - 1]} ${year}`;
    }
    const isoMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      let year = parseInt(isoMatch[1], 10);
      const month = parseInt(isoMatch[2], 10);
      const day = parseInt(isoMatch[3], 10);
      if (year < 2400) year += 543;
      return `${day} ${thaiMonths[month - 1]} ${year}`;
    }
    return s;
  } catch (e) {
    return String(dateStr);
  }
}

/**
 * ส่งอีเมลแจ้งลำดับคิวและชุดข้อมูลการเช็คชื่อผู้สมัคร
 */
function sendCheckInEmail(data) {
  try {
    const email = data.email || (data.interview ? data.interview.Email : null);

    if (!email) {
      console.warn("ไม่พบอีเมลผู้สมัครสำหรับส่งแจ้งคิว");
      return { success: false, message: "ไม่พบอีเมล" };
    }

    const displayPosition = data.positionApplied || data.employeeCategory || "";

    const employeeCategoryRow = displayPosition
      ? `<div style="border: 0.5px solid #e5e7eb; border-radius: 14px; padding: 16px; grid-column: span 2;">
       <div style="font-size: 11px; color: #9ca3af; font-weight: 400; margin-bottom: 6px;">ตำแหน่งที่สมัคร</div>
       <div style="font-size: 15px; color: #111111; font-weight: 500;">${displayPosition}</div>
     </div>`
      : "";

    const queueNoStr = String(data.queueNo || '-');
    const queueFontSize =
      queueNoStr.length <= 1 ? 44 :
        queueNoStr.length === 2 ? 40 :
          queueNoStr.length === 3 ? 36 : 30;

    const htmlBody = `
      <div style="font-family: 'Prompt', 'Sarabun', Arial, sans-serif; background: #f4f5f7; padding: 24px 0;">
        <div style="max-width: 600px; margin: 0 auto; border-radius: 20px; overflow: hidden; background: #ffffff; border: 0.5px solid #d1d5db;">
          <div style="background: #ffffff; padding: 28px 30px; border-bottom: 3px solid #5cb030; text-align: center;">
            <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/COM7_Logo.svg/1280px-COM7_Logo.svg.png" alt="COM7" style="height: 60px; object-fit: contain;" />
          </div>
          <div style="padding: 26px 30px 4px 30px;">
            <p style="margin: 0 0 4px 0; font-size: 13px; color: #9ca3af; font-weight: 400;">เรียนคุณ</p>
            <p style="margin: 0; font-size: 20px; font-weight: 500; color: #111111;">${data.fullname}</p>
            <p style="margin: 8px 0 0 0; font-size: 13px; color: #6b7280; font-weight: 400; line-height: 1.7;">ลำดับคิวสัมภาษณ์ของคุณในงาน ${data.eventName || '-'} โปรดนำข้อมูลนี้ไปใช้ประกอบการเข้ารับการสัมภาษณ์</p>
          </div>
          
          
          <div style="margin: 22px 30px 24px 30px; text-align: center;">
            <div style="display: inline-block; border-radius: 20px; background-color: #EFF8EA; padding: 22px 20px; min-width: 110px; border: 1px solid #d4ebd0;">
              <div style="font-size: ${parseInt(queueFontSize) + 20}px; font-weight: 700; color: #059669; line-height: 1; margin-bottom: 8px;">
                ${queueNoStr}
              </div>
              <div style="font-size: 13px; color: #059669; font-weight: 500; letter-spacing: 0.5px;">
                ลำดับคิวของคุณ
              </div>
            </div>
          </div>
          
          <div style="margin: 0 30px 24px 30px; display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
            <div style="border: 0.5px solid #e5e7eb; border-radius: 14px; padding: 16px;">
              <div style="font-size: 11px; color: #9ca3af; font-weight: 400; margin-bottom: 6px;">วันที่สัมภาษณ์</div>
              <div style="font-size: 15px; color: #111111; font-weight: 500;">${data.date || '-'}</div>
            </div>
            <div style="border: 0.5px solid #e5e7eb; border-radius: 14px; padding: 16px;">
              <div style="font-size: 11px; color: #9ca3af; font-weight: 400; margin-bottom: 6px;">เวลาสัมภาษณ์</div>
              <div style="font-size: 15px; color: #111111; font-weight: 500;">${data.time || '-'}</div>
            </div>
            ${employeeCategoryRow}
          </div>
          <div style="margin: 0 30px 24px 30px; padding-top: 20px; border-top: 0.5px dashed #d1d5db; font-family: sans-serif;">
            <p style="margin: 0 0 12px 0; font-size: 13px; font-weight: 600; color: #111111;">ข้อควรทราบ</p>
            <p style="margin: 0; font-size: 13px; color: #6b7280; font-weight: 400; line-height: 1.9;">
              • กรุณามาให้ทันเวลานัดหมาย เพื่อให้การสัมภาษณ์เป็นไปอย่างราบรื่น<br>
              • หากมีการเปลี่ยนแปลง กรุณาแจ้งฝ่ายทรัพยากรบุคคลล่วงหน้า<br>
            </p>
          </div>
          <div style="margin: 0 30px 28px 30px; background: #eff8ea; border-radius: 0 14px 14px 0; padding: 18px 20px; border-left: 3px solid #5cb030;">
            <div style="font-size: 13px; font-weight: 600; color: #111111; margin-bottom: 10px;">ฝ่ายทรัพยากรบุคคล</div>
            <div style="font-size: 12px; color: #6b7280; margin-bottom: 10px; font-weight: 400; line-height: 1.8;">
              ร่วมเป็นส่วนหนึ่งของบริษัท COM7 แค่เพียงคุณสมัครงานเข้ามาเพื่อคว้าโอกาสดีๆกับทางบริษัท
            </div>
            <div style="font-size: 13px; color: #111111; font-weight: 500; margin-bottom: 14px;">
              📞 02-017-7777 ต่อ 7212, 7208, 7209, 7210
            </div>
            <div style="display: flex; gap: 12px;">
              <a href="https://web.facebook.com/Comsevencareer?_rdc=1&_rdr#" target="_blank">
                <img src="https://img.icons8.com/fluency/48/facebook-new.png" width="36" height="36" alt="Facebook" style="display:block; border-radius:50%;" />
              </a>
              <a href="https://www.instagram.com/comsevencareer?igshid=OGQ5ZDc2ODk2ZA%3D%3D" target="_blank">
                <img src="https://img.icons8.com/fluency/48/instagram-new.png" width="36" height="36" alt="Instagram" style="display:block; border-radius:50%;" />
              </a>
              <a href="https://page.line.me/comsevencareer?openQrModal=true" target="_blank">
                <img src="https://img.icons8.com/color/48/line-me.png" width="36" height="36" alt="Line" style="display:block; border-radius:50%;" />
              </a>
              <a href="https://www.tiktok.com/@comsevencareer?_t=8fmsCUSQOxe&_r=1" target="_blank">
                <img src="https://img.icons8.com/color/48/tiktok--v1.png" width="36" height="36" alt="TikTok" style="display:block; border-radius:50%;" />
              </a>
            </div>
          </div>
          <div style="background: linear-gradient(135deg, #059669 0%, #10b981 100%); padding: 16px; text-align: center;">
            <span style="color: #ffffff; font-size: 11px; font-weight: 400;">&copy; 2026 Comseven Public Company Limited</span>
          </div>
        </div>
      </div>
`;

    MailApp.sendEmail({
      to: email,
      subject: `คิวสัมภาษณ์ของคุณ - ลำดับที่ ${data.queueNo} | ${data.eventName || 'COM7'}`,
      htmlBody: htmlBody,
      name: "COM7 Recruitment"
    });

    console.log(`✅ ส่งอีเมลเช็คชื่อให้ ${email} สำเร็จ (คิวที่ ${data.queueNo})`);
    return { success: true };
  } catch (error) {
    console.error("Error sending check-in email:", error);
    return { success: false, message: error.message };
  }
}