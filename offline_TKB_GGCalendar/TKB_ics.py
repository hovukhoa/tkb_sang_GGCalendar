import pandas as pd
from ics import Calendar, Event
from datetime import datetime, timedelta
import pytz

# Múi giờ Việt Nam
tz = pytz.timezone("Asia/Ho_Chi_Minh")

# Định nghĩa thời gian tiết học
def get_time_slot(session, period):
    morning_slots = [
        ("07:00", "07:45"), ("07:50", "08:35"),
        ("08:55", "09:40"), ("09:45", "10:30"), ("10:40", "11:25")
    ]
    afternoon_slots = [
        ("12:45", "13:30"), ("13:35", "14:20"),
        ("14:40", "15:25"), ("15:30", "16:15"), ("16:25", "17:10")
    ]
    
    if session == "Sáng":
        return morning_slots[period - 1]
    elif session == "Chiều":
        return afternoon_slots[period - 1]
    else:
        raise ValueError("Lỗi xác định buổi học!")

# Đọc file Excel
def read_schedule(file_path):
    df = pd.read_excel(file_path, header=None)
    days = ["T2", "T3", "T4", "T5", "T6", "T7"]
    schedule = []
    
    for index, row in df.iterrows():
        period_info = row[0]  # Lấy thông tin tiết học (S1, S2, C1, C2,...)
        
        if isinstance(period_info, str) and period_info.startswith("S"):
            session = "Sáng"
            period = int(period_info[1])  # Lấy số tiết
        elif isinstance(period_info, str) and period_info.startswith("C"):
            session = "Chiều"
            period = int(period_info[1])  # Lấy số tiết
        else:
            continue  # Bỏ qua dòng không hợp lệ
        
        for i, day in enumerate(days):
            class_info = row[i + 1]  # Lấy tên lớp
            if pd.notna(class_info):
                schedule.append((day, session, period, class_info))
    
    return schedule

# Tạo file ICS
def create_ics(schedule):
    calendar = Calendar()
    start_date = datetime(2025, 3, 3)  # Ngày bắt đầu tuần 25
    day_mapping = {"T2": 0, "T3": 1, "T4": 2, "T5": 3, "T6": 4, "T7": 5}
    
    for day, session, period, class_info in schedule:
        lesson_date = start_date + timedelta(days=day_mapping[day])
        start_time, end_time = get_time_slot(session, period)
        
        start_datetime = tz.localize(datetime.strptime(f"{lesson_date.date()} {start_time}", "%Y-%m-%d %H:%M"))
        end_datetime = tz.localize(datetime.strptime(f"{lesson_date.date()} {end_time}", "%Y-%m-%d %H:%M"))
        
        event = Event()
        event.name = f"Dạy lớp {class_info}"
        event.begin = start_datetime
        event.end = end_datetime
        calendar.events.add(event)
    
    with open("TKB.ics", "w", encoding="utf-8") as f:
        f.writelines(calendar)
    
    print("✅ Đã tạo file TKB.ics (Định dạng giờ đã chuẩn)")

# Chạy chương trình
file_path = "GV.xlsx"  # Đường dẫn file Excel
schedule = read_schedule(file_path)
create_ics(schedule)
