function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
const API_KEY = "Uandes@2026";
function removeAccents(str) {
  return String(str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
