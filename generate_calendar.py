import pandas as pd
import json
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

EXCEL_PATH = os.path.join('Temploy', 'Посещаемость кружки и ИМ 2025_2026.xlsx')
OUTPUT_JSON = 'calendar_data.json'

xls = pd.ExcelFile(EXCEL_PATH)

# ============ ЕЛЕНА ИМ ============
df = pd.read_excel(xls, sheet_name='Елена ИМ', header=None)

# Student 1: Настя Кудисова (химия/физика) - columns 0-3
student1_name = "Настя Кудисова"
student1_subjects = "химия / физика"
student1_lessons = []

for i in range(1, len(df)):
    date_val = df.iloc[i, 1]
    subject = df.iloc[i, 2]
    status = df.iloc[i, 3]

    if pd.notna(date_val) and pd.notna(subject):
        date_str = pd.Timestamp(date_val).strftime('%Y-%m-%d') if pd.notna(date_val) else None
        student1_lessons.append({
            'date': date_str,
            'subject': str(subject).strip(),
            'status': str(status).strip() if pd.notna(status) else 'нет данных'
        })

# Student 2: Харченко Матвей - columns 5-9
student2_name = "Харченко Матвей"
student2_subjects = "3 кл (русский, литература, математика, окружающий, английский, ЭИ)"
student2_lessons = []

for i in range(1, len(df)):
    date_val = df.iloc[i, 7]
    subject = df.iloc[i, 8]
    status = df.iloc[i, 9]

    if pd.notna(date_val) and pd.notna(subject):
        date_str = pd.Timestamp(date_val).strftime('%Y-%m-%d') if pd.notna(date_val) else None
        if date_str:
            student2_lessons.append({
                'date': date_str,
                'subject': str(subject).strip(),
                'status': str(status).strip() if pd.notna(status) else 'нет данных'
            })

# ============ ПОСЕЩАЕМОСТЬ ============
df_att = pd.read_excel(xls, sheet_name='ПОСЕЩАЕМОСТЬ', header=None)

groups = []

# Parse group blocks
i = 0
while i < len(df_att):
    group_name = df_att.iloc[i, 0]
    if pd.notna(group_name) and isinstance(group_name, str) and group_name.strip():
        # Found a group header row - get dates from columns 1+
        dates = []
        for c in range(1, len(df_att.columns)):
            d = df_att.iloc[i, c]
            if pd.notna(d):
                try:
                    dates.append(pd.Timestamp(d).strftime('%Y-%m-%d'))
                except:
                    pass

        # Read students below
        students = []
        j = i + 1
        while j < len(df_att):
            student_name = df_att.iloc[j, 0]
            if pd.isna(student_name) or not isinstance(student_name, str) or not student_name.strip():
                # Check if it's just a row with False values (empty student slot)
                has_any_true = False
                for c in range(1, min(len(dates) + 1, len(df_att.columns))):
                    if df_att.iloc[j, c] == True:
                        has_any_true = True
                        break
                if not has_any_true:
                    break
                j += 1
                continue

            attendance = {}
            for idx, date in enumerate(dates):
                col = idx + 1
                if col < len(df_att.columns):
                    val = df_att.iloc[j, col]
                    attendance[date] = bool(val) if pd.notna(val) else False

            students.append({
                'name': student_name.strip(),
                'attendance': attendance
            })
            j += 1

        if dates and students:
            groups.append({
                'name': group_name.strip(),
                'dates': dates,
                'students': students
            })
        i = j
    else:
        i += 1

# ============ BUILD OUTPUT ============
calendar_data = {
    'individual_lessons': [
        {
            'student': student1_name,
            'description': student1_subjects,
            'lessons': student1_lessons
        },
        {
            'student': student2_name,
            'description': student2_subjects,
            'lessons': student2_lessons
        }
    ],
    'group_attendance': groups
}

with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(calendar_data, f, ensure_ascii=False, indent=2)

print(f"Данные сохранены в {OUTPUT_JSON}")
print(f"Индивидуальные занятия:")
print(f"  {student1_name}: {len(student1_lessons)} занятий")
print(f"  {student2_name}: {len(student2_lessons)} занятий")
print(f"Групповые кружки: {len(groups)}")
for g in groups:
    print(f"  {g['name']}: {len(g['students'])} учеников, {len(g['dates'])} дат")
