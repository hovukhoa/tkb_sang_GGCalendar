// Code.gs
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔔 Thời Khóa Biểu')
    .addItem('📅 Cập nhật lịch dạy', 'showDatePickerDialog')
    .addToUi();
}

// Hiển thị hộp thoại chọn ngày
function showDatePickerDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DatePicker')
    .setWidth(300)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Chọn ngày Thứ Hai');
}

// Nhận ngày từ DatePicker.html và tạo sự kiện
function createEventsFromSelectedDate(dateString) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MAIN");
  var calendar = CalendarApp.getCalendarsByName("HOVUKHOA")[0]; // Tên của lịch đặt ở đây
  
  if (!calendar) {
    SpreadsheetApp.getUi().alert("Không tìm thấy lịch tên 'HOVUKHOA'");
    return;
  }
  
  var startDate = new Date(dateString); // Ngày Thứ Hai được chọn
  var days = {"T2": 0, "T3": 1, "T4": 2, "T5": 3, "T6": 4, "T7": 5};
  var timeSlots = {"S1": ["07:00", "07:45"], "S2": ["07:50", "08:35"],
                   "S3": ["08:55", "09:40"], "S4": ["09:45", "10:30"], "S5": ["10:40", "11:25"],
                   "C1": ["12:45", "13:30"], "C2": ["13:35", "14:20"],
                   "C3": ["14:40", "15:25"], "C4": ["15:30", "16:15"], "C5": ["16:25", "17:10"]};
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) { // Bỏ qua tiêu đề
    var session = data[i][0]; // S1, S2, ..., C5
    if (!timeSlots[session]) continue;
    
    for (var j = 1; j < 7; j++) { // Cột T2 -> T7
      var classInfo = data[i][j];
      if (classInfo) {
        var lessonDate = new Date(startDate);
        lessonDate.setDate(startDate.getDate() + days[Object.keys(days)[j-1]]);
        
        var startTime = timeSlots[session][0];
        var endTime = timeSlots[session][1];
        
        var eventStart = new Date(lessonDate.toDateString() + " " + startTime);
        var eventEnd = new Date(lessonDate.toDateString() + " " + endTime);
        
        calendar.createEvent("Dạy lớp " + classInfo, eventStart, eventEnd);
      }
    }
  }
  SpreadsheetApp.getUi().alert("Đã cập nhật lịch vào Google Calendar!");
}