import pandas as pd
import json
import os
import sys
import hashlib

sys.stdout.reconfigure(encoding='utf-8')

EXCEL_PATH = os.path.join('Temploy', 'Посещаемость кружки и ИМ 2025_2026.xlsx')
DATA_FILE = 'data.json'

def make_id(text):
    return hashlib.md5(text.encode()).hexdigest()[:10]

xls = pd.ExcelFile(EXCEL_PATH)

students = []
lessons = []
attendance = {}

# ============ ЕЛЕНА ИМ — Настя Кудисова ============
s1_id = make_id('Настя Кудисова')
students.append({'id': s1_id, 'name': 'Кудисова Настя', 'desc': 'химия, физика', 'group': ''})

df = pd.read_excel(xls, sheet_name='Елена ИМ', header=None)

for i in range(1, len(df)):
    date_val = df.iloc[i, 1]
    subject = df.iloc[i, 2]
    status = df.iloc[i, 3]
    if pd.notna(date_val) and pd.notna(subject):
        ds = pd.Timestamp(date_val).strftime('%Y-%m-%d')
        subj = str(subject).strip()
        lid = make_id(f's1_{ds}_{subj}')
        lessons.append({'id': lid, 'studentId': s1_id, 'subject': subj, 'date': ds, 'time': ''})
        st = str(status).strip().lower() if pd.notna(status) else ''
        if st in ('пришла', 'пришел'):
            attendance[f'{ds}_{s1_id}_{subj}'] = True
        elif 'отмена' in st or 'перенесли' in st:
            attendance[f'{ds}_{s1_id}_{subj}'] = False

# ============ ЕЛЕНА ИМ — Харченко Матвей ============
s2_id = make_id('Харченко Матвей')
students.append({'id': s2_id, 'name': 'Харченко Матвей', 'desc': '3 кл: русский, литература, математика, окружающий, английский, ЭИ', 'group': ''})

for i in range(1, len(df)):
    date_val = df.iloc[i, 7]
    subject = df.iloc[i, 8]
    status = df.iloc[i, 9]
    if pd.notna(date_val) and pd.notna(subject):
        ds = pd.Timestamp(date_val).strftime('%Y-%m-%d')
        subj = str(subject).strip()
        # skip non-lesson entries
        if subj in ('каникулы', 'не было занятий'):
            continue
        if 'перенесли' in subj:
            continue
        lid = make_id(f's2_{ds}_{subj}_{i}')
        lessons.append({'id': lid, 'studentId': s2_id, 'subject': subj, 'date': ds, 'time': ''})
        st = str(status).strip().lower() if pd.notna(status) else ''
        if st in ('пришла', 'пришел'):
            attendance[f'{ds}_{s2_id}_{subj}'] = True
        elif 'отмена' in st or 'перенесли' in st:
            attendance[f'{ds}_{s2_id}_{subj}'] = False

# ============ ПОСЕЩАЕМОСТЬ — групповые кружки ============
df_att = pd.read_excel(xls, sheet_name='ПОСЕЩАЕМОСТЬ', header=None)

i = 0
while i < len(df_att):
    group_name = df_att.iloc[i, 0]
    if pd.notna(group_name) and isinstance(group_name, str) and group_name.strip():
        gname = group_name.strip()
        dates = []
        for c in range(1, len(df_att.columns)):
            d = df_att.iloc[i, c]
            if pd.notna(d):
                try:
                    dates.append(pd.Timestamp(d).strftime('%Y-%m-%d'))
                except:
                    pass

        j = i + 1
        while j < len(df_att):
            sname = df_att.iloc[j, 0]
            if pd.isna(sname) or not isinstance(sname, str) or not sname.strip():
                has_any = any(df_att.iloc[j, c] == True for c in range(1, min(len(dates)+1, len(df_att.columns))))
                if not has_any:
                    break
                j += 1
                continue

            sname = sname.strip()
            sid = make_id(f'grp_{gname}_{sname}')
            students.append({'id': sid, 'name': sname, 'desc': gname, 'group': gname})

            for idx, ds in enumerate(dates):
                col = idx + 1
                if col < len(df_att.columns):
                    val = df_att.iloc[j, col]
                    present = bool(val) if pd.notna(val) else False
                    lid = make_id(f'grp_{sid}_{ds}')
                    lessons.append({'id': lid, 'studentId': sid, 'subject': gname, 'date': ds, 'time': ''})
                    if present:
                        attendance[f'{ds}_{sid}_{gname}'] = True
                    else:
                        attendance[f'{ds}_{sid}_{gname}'] = False
            j += 1
        i = j
    else:
        i += 1

# ============ SAVE ============
data = {
    'students': students,
    'lessons': lessons,
    'attendance': attendance
}

with open(DATA_FILE, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f'Импортировано:')
print(f'  Учеников: {len(students)}')
print(f'  Занятий: {len(lessons)}')
print(f'  Отметок посещаемости: {len(attendance)}')
print(f'Данные сохранены в {DATA_FILE}')
