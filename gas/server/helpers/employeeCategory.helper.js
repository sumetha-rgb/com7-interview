/**
 * แปลงข้อความตำแหน่งงาน (Position_Applied) ให้เป็นหมวดกว้าง "พนักงานหน้าร้าน" หรือ "ออฟฟิศ"
 * ใช้ includes() (keyword match) แทน exact match เพื่อรองรับได้ทั้งข้อมูลเก่า
 * (string รวมสาขา เช่น "...True & Studio7 & BaNaNa") และข้อมูลใหม่ (แยกสาขาเดี่ยว)
 */
function getEmployeeCategoryFromJobType(positionText) {
  const text = String(positionText || "").trim();
  if (!text) return "";
  if (text.includes("หน้าร้าน")) return "พนักงานหน้าร้าน";
  if (text.includes("สำนักงานใหญ่") || text.includes("ฝึกงาน")) return "ออฟฟิศ";
  return "";
}