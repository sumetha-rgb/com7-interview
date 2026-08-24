function getEventNameDropdown(userRole) {
  const roleValue = userRole ? userRole : "USER";

  // --- OPTIMIZATION: cache รายชื่อ Event ดิบ (ก่อน filter ตาม role/เวลา) 3
  // นาที เพราะฟังก์ชันนี้ถูกเรียกแทบทุกครั้งที่เปิดหน้าเว็บ/ฟอร์มสมัคร แต่
  // ข้อมูล Event เปลี่ยนไม่บ่อยขนาดต้องอ่านชีทสดทุกครั้ง การ filter ตาม
  // role/เวลายังทำสดเหมือนเดิมทุกครั้ง (แค่แคชข้อมูลดิบ) ผลลัพธ์จึงตรงกับ
  // เดิมทุกกรณี
  const cache = CacheService.getScriptCache();
  const cacheKey = "dropdown_events_raw";
  let eventData;
  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      eventData = JSON.parse(cached);
    }
  } catch (e) {
    console.warn("Cannot read dropdown_events_raw cache: " + e.message);
  }

  if (!eventData) {
    const eventSheet = getEventSheet();
    eventData = getRowsFromSheet(eventSheet);
    try {
      cache.put(cacheKey, JSON.stringify(eventData), 180); // 3 นาที
    } catch (e) {
      console.warn("Cannot write dropdown_events_raw cache: " + e.message);
    }
  }

  if (roleValue === USER_ROLE || roleValue === MANAGER_ROLE) {
    console.log(roleValue === USER_ROLE || roleValue === MANAGER_ROLE);
    const now = new Date();
    return eventData
      .filter((event) => {
        return (
          now > normalizeDate(event.Opening_At) &&
          now < normalizeDate(event.Closing_At) &&
          String(event.Status).toLowerCase() ===
            String(EVENT_STATUS[0]).toLowerCase()
        );
      })
      .map((event) => event.Event_Name);
  }

  return eventData.map((event) => event.Event_Name);
}

function getJobTypeDropdown() {
  // --- OPTIMIZATION: cache 5 นาที เพราะรายการตำแหน่งงานเปลี่ยนไม่บ่อยเลย
  const cache = CacheService.getScriptCache();
  const cacheKey = "dropdown_job_types";
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {
    console.warn("Cannot read dropdown_job_types cache: " + e.message);
  }

  const jobTypeSheet = getJobTypeSheet();
  const jobTypeData = getRowsFromSheet(jobTypeSheet);
  const result = jobTypeData
    .map((job) => {
      return {
        name: job.Name,
      };
    })
    .sort();

  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 300); // 5 นาที
  } catch (e) {
    console.warn("Cannot write dropdown_job_types cache: " + e.message);
  }

  return result;
}