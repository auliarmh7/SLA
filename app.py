import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import io

# =========================
# UI HEADER
# =========================
st.set_page_config(page_title="SLA Pengaduan", layout="wide")
st.title("📊 Perhitungan SLA Pengaduan Otomatis")

st.info("""
Cara penggunaan:
1. Upload file Excel
2. Klik tombol 'Proses Hitung SLA'
3. Download hasil
""")

# =========================
# INPUT LIBUR NASIONAL
# =========================
st.sidebar.header("⚙️ Pengaturan")

default_libur = [
    '2025-01-01','2025-01-27','2025-01-29',
    '2025-03-28','2025-03-29','2025-03-31',
    '2025-04-01','2025-04-02','2025-04-03','2025-04-04',
    '2025-04-07','2025-04-18','2025-04-20',
    '2025-05-01','2025-05-12','2025-05-13','2025-05-29','2025-05-30',
    '2025-06-01','2025-06-06','2025-06-09','2025-06-27',
    '2025-08-17','2025-08-18','2025-09-05',
    '2025-12-25','2025-12-26'
]

libur_input = st.sidebar.text_area(
    "Daftar Libur (pisahkan dengan koma)",
    ",".join(default_libur)
)

libur_nasional = [pd.to_datetime(t.strip()).date() for t in libur_input.split(",") if t.strip()]

# =========================
# UPLOAD FILE
# =========================
uploaded_file = st.file_uploader("📂 Upload File Excel", type=["xlsx"])

if uploaded_file is None:
    st.warning("Silakan upload file Excel terlebih dahulu")
    st.stop()

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(uploaded_file)
df.columns = df.columns.str.strip()

# =========================
# NORMALISASI KOLOM
# =========================
mapping_kolom = {
    'Pengaduan ID': 'pengaduan_id',
    'Response': 'response',
    'Response Sebelumnya': 'response_sebelumnya',
    'Respon yang Dihitung': 'response_yang_dihitung'
}
df = df.rename(columns=mapping_kolom)

# =========================
# VALIDASI KOLOM
# =========================
required_cols = ['pengaduan_id', 'response', 'response_sebelumnya']
missing_cols = [col for col in required_cols if col not in df.columns]

if missing_cols:
    st.error(f"Kolom berikut tidak ditemukan: {missing_cols}")
    st.stop()

st.success("✅ File berhasil diupload")
st.write("Preview Data:")
st.dataframe(df.head())

# =========================
# FILTER
# =========================
if 'response_yang_dihitung' in df.columns:
    df = df[df['response_yang_dihitung'].fillna('').str.upper() == 'YA']

st.info(f"Jumlah data setelah filter: {len(df)}")

# =========================
# DATETIME
# =========================
df['response'] = pd.to_datetime(df['response'], errors='coerce')
df['response_sebelumnya'] = pd.to_datetime(df['response_sebelumnya'], errors='coerce')
df = df.dropna(subset=['response', 'response_sebelumnya'])

# =========================
# FUNCTION
# =========================
def is_workday(date):
    return date.weekday() < 5 and date not in libur_nasional

def workday_hours_diff(start, end):
    if end <= start:
        return 0
    total_hours = 0
    current = start
    while current.date() <= end.date():
        if is_workday(current.date()):
            day_start = datetime.combine(current.date(), time(8,0))
            day_end = datetime.combine(current.date(), time(17,0))
            s = max(current, day_start)
            e = min(end, day_end)
            if s < e:
                total_hours += (e - s).total_seconds() / 3600
        current += timedelta(days=1)
        current = datetime.combine(current.date(), time(8,0))
    return total_hours

def calculate_working_time(start, end):
    total = timedelta(0)
    current = start
    while current < end:
        current_date = current.date()
        if current.weekday() < 5 and current_date not in libur_nasional:
            next_point = min(end, datetime.combine(current_date + timedelta(days=1), datetime.min.time()))
            total += next_point - current
        current = datetime.combine(current.date() + timedelta(days=1), datetime.min.time())
    return total

def hours_to_hhmmss(hours):
    total_seconds = int(hours * 3600)
    hh = total_seconds // 3600
    mm = (total_seconds % 3600) // 60
    ss = total_seconds % 60
    return f"{hh:02}:{mm:02}:{ss:02}"

# =========================
# FORMAT HARI JAM
# =========================
def hours_to_days_hours(hours):
    total_hours = int(hours)
    days = total_hours // 24
    remaining_hours = total_hours % 24
    return f"{days} hari {remaining_hours} jam"

# =========================
# TOMBOL PROSES
# =========================
if st.button("⚙️ Proses Hitung SLA"):

    results = []

    with st.spinner("Menghitung SLA..."):
        for _, row in df.iterrows():
            start = row['response_sebelumnya']
            end = row['response']

            jam_kerja = workday_hours_diff(start, end)
            waktu_kerja = calculate_working_time(start, end)
            waktu_kerja_jam = waktu_kerja.total_seconds() / 3600

            results.append({
                'pengaduan_id': row['pengaduan_id'],
                'response_time_hours': jam_kerja,
                'response_time_workingtime': waktu_kerja_jam
            })

    df_response = pd.DataFrame(results)

    avg_response = df_response.groupby('pengaduan_id').mean().reset_index()

    avg_response.rename(columns={
        'response_time_hours': 'avg_response_hours',
        'response_time_workingtime': 'avg_working_time_hours'
    }, inplace=True)

    avg_response['avg_working_time_hhmmss'] = avg_response['avg_working_time_hours'].apply(hours_to_hhmmss)

    # RATA-RATA TOTAL (WORKING TIME)
    overall_avg_working = avg_response['avg_working_time_hours'].mean()

    total_row = pd.DataFrame([{
        'pengaduan_id': 'RATA-RATA KESELURUHAN',
        'avg_working_time_hours': overall_avg_working,
        'avg_working_time_hhmmss': hours_to_hhmmss(overall_avg_working)
    }])

    avg_response = pd.concat([avg_response, total_row], ignore_index=True)

    # =========================
    # OUTPUT
    # =========================
    st.success("✅ Perhitungan selesai")

    col1, col2 = st.columns(2)

    col1.metric(
        "Rata-Rata SLA (Jam)",
        hours_to_hhmmss(overall_avg_working)
    )

    col2.metric(
        "Rata-Rata SLA (Hari & Jam)",
        hours_to_days_hours(overall_avg_working)
    )

    st.dataframe(avg_response[['pengaduan_id','avg_working_time_hours','avg_working_time_hhmmss']])

    # =========================
    # DOWNLOAD
    # =========================
    buffer = io.BytesIO()
    avg_response.to_excel(buffer, index=False)
    buffer.seek(0)

    st.download_button(
        label="⬇️ Download Hasil",
        data=buffer,
        file_name="hasil_sla.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )