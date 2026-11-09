# -*- coding: utf-8 -*-
"""
PHT-Tracker - Mandaya Royal Hospital Puri | Departemen Radiologi
  1) Form Digital Monitoring Suhu dan Kelembapan   (grafik titik SUHU & KELEMBAPAN per ruangan, 3 shift P/S/M)
  2) Form Digital Daily Quality Control            (daftar periksa harian per modalitas: KEGIATAN + PARAMETER)
"""
import base64
import calendar
import hmac
import datetime as dt
import io
import json
import math
import os
import textwrap
import time
import urllib.parse
import urllib.request
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Cadangan tema: pastikan warna utama biru tua (bukan merah bawaan Streamlit) walau .streamlit/config.toml belum terbaca.
try:
    for _k, _v in {"base": "light", "primaryColor": "#073763", "backgroundColor": "#ffffff",
                   "secondaryBackgroundColor": "#f2f7fc", "textColor": "#073763"}.items():
        st._config.set_option("theme." + _k, _v)
except Exception:
    pass

st.set_page_config(
    page_title="PHT-Tracker",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"Get help": None, "Report a bug": None, "About": None},
)

# =====================================================================
# 0. KONFIGURASI UTAMA (aman diubah)
# =====================================================================
APP_TITLE = "PHT-Tracker"
HOSPITAL = "Mandaya Royal Hospital Puri"
SITE = "PURI"
DEPT_LINE = "DEPARTEMEN RADIOLOGI MANDAYA ROYAL HOSPITAL PURI"
TZ = dt.timezone(dt.timedelta(hours=7))  # WIB. Streamlit Cloud berjalan di UTC, jadi zona waktu dipaksa ke WIB
QC_DUE_HOUR = 21                         # checklist QC hari ini dianggap "jatuh tempo" setelah jam ini (ada langkah "Mematikan Pesawat" di malam hari)
SUHU_NORMAL = (18.0, 23.0)               # "Target Temperatur (18 - 23)°C" di form suhu, berlaku untuk SEMUA ruangan
HUM_NORMAL = (40.0, 60.0)                # "Target kelembapan ( 40 - 60 % )" di form suhu

DATA_DIR = "data_mandaya"
ENTRY_FILE = os.path.join(DATA_DIR, "entries.csv")
QC_FILE = os.path.join(DATA_DIR, "qc_entries.csv")
CFG_FILE = os.path.join(DATA_DIR, "config.json")
LOGO_FILE = "Mandaya Logo.png"
FONT_REG = "Lexend-Regular.ttf"
FONT_BOLD = "Lexend-Bold.ttf"
os.makedirs(DATA_DIR, exist_ok=True)

FONT_URLS = {
    FONT_REG: [
        "https://github.com/google/fonts/raw/main/ofl/lexend/static/Lexend-Regular.ttf",
        "https://raw.githubusercontent.com/google/fonts/main/ofl/lexend/static/Lexend-Regular.ttf",
    ],
    FONT_BOLD: [
        "https://github.com/google/fonts/raw/main/ofl/lexend/static/Lexend-Bold.ttf",
        "https://raw.githubusercontent.com/google/fonts/main/ofl/lexend/static/Lexend-Bold.ttf",
    ],
}

BLN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
SHIFT_NAME = {"P": "Pagi", "S": "Siang", "M": "Malam"}
# Jam pengisian per shift. Diambil dari kebiasaan pengisian di data lama (puncak jam 07, 13, dan 20).
SHIFT_HOUR = {"P": 7, "S": 13, "M": 20}
SHIFT_LABEL = {k: "{} - {} ({:02d}.00)".format(k, SHIFT_NAME[k], SHIFT_HOUR[k]) for k in "PSM"}
SHIFT_TIMES = " · ".join("{:02d}.00".format(SHIFT_HOUR[k]) for k in "PSM")

# Ruangan untuk form suhu & kelembapan (nama persis seperti kolom "RUANG :" di form).
_ROOMS = ["Digital Diagnost C50", "Combi Diagnost R90", "Amulet Innovality", "USG Philips Affiniti 70",
          "Iqon Spectral CT", "R. Teknik CT Scan", "MRI Ambition X", "R. Teknik MRI", "Lemari ALKES", "Lemari CD"]


def _dev(modalitas, sections, ruang=None, qr_ruang=None):
    """Satu form Daily QC.  sections = [(judul seksi, [(KEGIATAN, PARAMETER), ...]), ...]  (persis urutan di form).

    ruang    = ruangan suhu yang QR-nya dipakai untuk mengisi QC alat ini.
    qr_ruang = dipakai bila alat tidak punya ruangan sendiri (mis. Mobile X-Ray memakai QR ruangan lain).
    """
    secs, items = [], []
    for title, rows in sections:
        its = []
        for n, (keg, par) in enumerate(rows, 1):
            key = "{}::{}".format(title, keg)   # kunci unik per alat -> disimpan di database
            its.append(dict(key=key, no=str(n), kegiatan=keg, param=par))
            items.append(key)
        secs.append(dict(title=title, items=its))
    return dict(nama=modalitas, ruang=ruang, qr_ruang=qr_ruang, info=[["Modalitas", modalitas], ["Lokasi", HOSPITAL]],
                sections=secs, items=items)


# Daftar periksa harian (disalin persis dari "Form Digital Daily Quality Control Tahun 2026.xlsx").
# Salah ketik di form asli (mis. "Keberihan", "Pesawat ON 10 selama menit") sengaja dibiarkan; ubah di sini bila mau dikoreksi.
_DEVICES = [
    # --- form "QC Konvensional DR" ---
    _dev("Digital Diagnost C50", ruang="Digital Diagnost C50", sections=[
        ("Pengecekan Pesawat X - Ray", [
            ("Panel ON", "Lampu indikator menyala"),
            ("Pesawat ON", "Lampu indikator menyala"),
            ("Cek parameter kV", "Display akan muncul"),
            ("Cek parameter mAs", "Display akan muncul"),
            ("Cek lampu indikator", "Indikator menyala"),
            ("Cek kedua diafragma / bukaan kolimator", "Dapat dibuka-tutup"),
            ("Cek FFD / Pergerakan mekanik", "Dapat dinaik turunkan"),
            ("Cek rotasi tube", "Dapat dilakukan penyudutan"),
            ("Cek tube jalan kanan & kiri", "Tabung X-Ray dapat bergerak dengan mulus tanpa tersendat"),
            ("Cek bucky table", "Bisa digerakkan ke atas dan ke bawah"),
            ("Cek fungsi pergeseran meja", "Bisa digerakkan ke seluruh arah"),
            ("Jumlah grid", "Seluruh jumlah grip tersedia"),
        ]),
        ("Warming Up Pesawat", [
            ("Tutup kedua diafragma", "Pastikan tombol dapat diputar"),
            ("Pesawat ON 10 selama menit", "Pastikan alat menyala tanpa ada notifikasi error"),
            ("Lakukan eksposi", "Pastikan alat berhasil ekspos tanpa ada notifikasi error dan noise"),
        ]),
        ("Kondisi Ruangan", [
            ("Kebersihan ruangan", "Tidak terlihat sampah"),
            ("Keberihan meja kontrol & meja komputer", "Tidak berdebu & bernoda"),
            ("Bantal", "1 buah"),
            ("Suhu ruangan", "18 - 22 C"),
            ("Kelembapan Udara", "40 - 60 %"),
        ]),
        ("Mematikan Pesawat", [
            ("Turunkan semua parameter kV & mAs", ""),
            ("Posisikan tube vertikal", ""),
            ("Tekan tombol off pada kontrol pesawat", ""),
            ("Matikan panel listrik", ""),
        ]),
        ("Kondisi Changing Room No 3 dan No 4", [
            ("Kelengkapan baju dan celana pasien", "baju dan celana tersedia untuk seluruh ukuran"),
            ("Kebersihan Loker", "kondisi loker pasien bersih dan tidak terdapat sampah"),
            ("Kondisi smart lock loker", "Kondisi loker bisa dibuka dan ditutup menggunakan gelang"),
            ("Kondisi pintu loker", "Kondisi pintu loker bisa terkunci dan tertutup rapat"),
            ("Kelengkapan jumlah gelang smart lock", "Satu changing room memiliki 4 buah smart lock loker"),
        ]),
    ]),
    # --- form "QC Fluoroskopi" ---
    _dev("Combi Diagnost R90", ruang="Combi Diagnost R90", sections=[
        ("Cek Pesawat", [
            ("Panel ON", "Lampu indikator menyala"),
            ("Pesawat ON", "Lampu indikator menyala"),
            ("Cek parameter kV", "Display akan muncul"),
            ("Cek parameter mAs", "Display akan muncul"),
            ("Cek lampu indikator", "Menyala"),
            ("Cek kedua diafragma", "Dapat dibuka-tutup"),
            ("Cek FFD", "Dapat dinaik turunkan"),
            ("Cek rotasi meja", "Dapat dilakukan penyudutan"),
            ("Cek tube jalan kanan & kiri", "Jalan rel mudah / seret"),
            ("Cek bucky table", "Jalan / tidak"),
            ("Cek fungsi pergeseran meja", "Berfungsi / tidak"),
        ]),
        ("Warming Up Pesawat", [
            ("Tutup kedua diafragma", "Putar tombol"),
            ("Pesawat ON 10 menit", ""),
            ("Lakukan eksposi", ""),
        ]),
        ("Kondisi Ruangan", [
            ("Kebersihan ruangan", "Tidak terlihat sampah"),
            ("Keberihan meja kontrol & meja komputer", "Tidak berdebu & bernoda"),
            ("Bantal", "1 buah"),
            ("Suhu ruangan", "18 - 22 C"),
            ("Kelembapan Udara", "40 - 60 %"),
        ]),
        ("Mematikan Pesawat", [
            ("Turunkan semua parameter kV & mAs", ""),
            ("Posisikan tube vertikal", ""),
            ("Tekan tombol off pada kontrol pesawat", ""),
            ("Matikan panel listrik", ""),
        ]),
    ]),
    # --- form "QC Mammografi" ---
    _dev("Amulet Innovality", ruang="Amulet Innovality", sections=[
        ("Pengecekan Pesawat", [
            ("Control box/ panel listrik", "Lampu indikator menyala"),
            ("Pesawat ON", "Tidak terdapat notifikasi error"),
            ("Pengecekan lampu kolimasi", "Pastikan lampu kolimasi tepat pada area bucky"),
            ("X-Ray Ready", "Pastikan status expose sudah dalam keadaan ready, apabila tidak restart alat dan perhatikan suhu ruangan"),
            ("Cek pergerakan gantry mammografi", "Dapat dinaik-turunkan dan melakukan rotasi"),
            ("Pengecekan tombol ekspose dan tombol emergency", "Tidak terdapat kerusakan / tidak berfungsi"),
            ("Pengecekan Paddle Compression", "Lakukan penekanan terhadap objek dan tidak terdapat notifikasi error"),
        ]),
        ("Warming Up Pesawat", [
            ("Pesawat ON 10 menit", "Tunggu autocalibration dilakukan oleh system dan pastikan daily calibration passed"),
            ("Lakukan eksposi", "Lakukan eksposi terhadap obek dan pastikan tidak terdapat gambaran yang terpotong atau noise"),
        ]),
        ("Kondisi Ruangan", [
            ("Kebersihan ruangan", "Tidak terlihat sampah"),
            ("Keberihan meja kontrol & meja komputer", "Tidak berdebu & bernoda"),
            ("Suhu ruangan", "18 - 22 C"),
            ("Kelembapan Udara", "40 - 60 %"),
        ]),
        ("Mematikan Pesawat", [
            ("Turunkan semua parameter kV & mAs", "Pastikan tidak terdapat notifikasi error"),
            ("Tekan tombol off pada kontrol pesawat", ""),
            ("Matikan panel listrik", "Tekan tombol merah pada panel"),
        ]),
    ]),
    # --- form "QC USG" ---
    _dev("USG Philips Affiniti 70", ruang="USG Philips Affiniti 70", sections=[
        ("Cek Kelengkapan & Peralatan Penunjang", [
            ("Control box/ panel listrik", "Lampu indikator menyala"),
            ("AC", "AC Berfungsi Baik"),
            ("Pintu ruangan", "Pintu bisa dikunci"),
            ("Cek Fungsi Display USG", "Pastikan display USG muncul pada monitor TV"),
            ("Cek fungsi tombol", "Pastikan seluruh tombol (termasuk trackball) dalam kondisi baik"),
            ("Pengecekan UPS", "Cek kondisi baterai UPS, pastikan dalam kondisi normal dan tidak terdapat notifikasi error"),
            ("Cek wheel & locks", "Pastikan pergerakan roda dan mesin USG dapat bergerak dengan lancar"),
            ("Cek scanning housing", "Cek kebersihan dan fisik"),
        ]),
        ("Warming Up Pesawat", [
            ("USG dalam Kondisi ON", "Lakukan Restart Sistem Setiap Pagi"),
        ]),
        ("Kondisi Ruangan", [
            ("Kebersihan ruangan", "Tidak terlihat sampah"),
            ("Kebersihan meja control", "Tidak berdebu & bernoda"),
            ("Suhu ruangan", "18 - 22 C"),
            ("Kelembapan udara", "40 - 60 %"),
            ("Check tissue, sabun wastafel & gel USG", "Tersedia dengan baik dan lengkap"),
            ("Fungsi Gel Warmer", "Pastikan gel warmer berfungsi dengan baik"),
        ]),
        ("Mematikan Pesawat", [
            ("Tekan Shut down pada Monitor", "Pastikan tidak terdapat notifikasi error"),
            ("Mematikan Gel Warmer", "Tekan switch off pada gel wamer"),
        ]),
    ]),
    # --- form "QC CT-Scan" ---
    _dev("Iqon Spectral CT", ruang="Iqon Spectral CT", sections=[
        ("Cek Kelengkapan & Peralatan Penunjang", [
            ("Control box / panel listrik", "Lampu indikator menyala"),
            ("AC", "Menyala / Berfungsi"),
            ("Pintu ruangan", "Dapat terkunci rapat"),
            ("Aksesoris Pemeriksaan", "Seluruh perlengkapan dalam kondisi yang lengkap"),
        ]),
        ("Persiapan Pesawat CT - Scan", [
            ("Pesawat & gantry ON", "Tidak terdapat notifikasi error"),
            ("Monitor workstation ON", "Booting dalam waktu yang wajar"),
            ("Fungsi Intercom", "Suara dari dalam terdengar keluar dan sebaliknya"),
            ("Warm up / Short tube Conditioning", "100 % Passed"),
            ("Pengecekan UPS", "Cek kondisi baterai UPS, pastikan dalam kondisi normal dan tidak terdapat notifikasi error"),
        ]),
        ("Pemeriksaan Injektor", [
            ("Injektor ON", "Injektor dan CT Scan terhubung tanpa kendala koneksi"),
            ("Cek Fungsi Syringe", "Dapat mengisi & membuang cairan untuk kedua channel syringe"),
            ("Tes publik sistem", "Suara Sistem Injector Terdengar secara jelas"),
        ]),
        ("Kondisi Ruangan", [
            ("Kebersihan ruangan", "Ruangan bersih dan tidak terdapat sampah"),
            ("Kebersihan meja operator", "Tidak terdapat debu dan noda"),
            ("Suhu ruangan", "18 - 22 C"),
            ("Kelembapan Udara", "40 - 60 %"),
            ("Ketersediaan linen, bantal dan tissue gulung", "Tersedia secara lengkap"),
            ("Kelengkapan ALKES untuk tindakan", "Tersedia secara lengkap"),
            ("Jumlah standar infus", "1 buah"),
            ("Selimut", "1 buah"),
        ]),
        ("Mematikan CT - Scan", [
            ("Mematikan CT Scan", "Pilih Menu Shutdown Pada Konsol"),
            ("Matikan Gantry", "Putar kunci gantry dan pastikan lampu indikator pada gantry telah mati"),
        ]),
    ]),
    # --- form "QC Mobile" ---
    _dev("Mobile Diagnost wDR", qr_ruang="Digital Diagnost C50", sections=[
        ("Cek Pesawat", [
            ("Panel ON", "Lampu indikator menyala"),
            ("Pesawat ON", "Lampu indikator menyala"),
            ("Cek parameter kV", "Display akan muncul"),
            ("Cek parameter mAs", "Display akan muncul"),
            ("Cek lampu indikator", "Menyala"),
            ("Cek kedua diafragma", "Dapat dibuka-tutup"),
            ("Cek FFD", "Dapat dinaik turunkan"),
            ("Cek rotasi tube", "Dapat dilakukan penyudutan"),
            ("Cek tube jalan kanan & kiri", "Jalan rel mudah / seret"),
            ("Cek fungsi pergeseran roda", "Berfungsi / tidak"),
            ("Jumlah grid", "Lengkap"),
        ]),
        ("Warming Up Pesawat", [
            ("Tutup kedua diafragma", "Putar tombol"),
            ("Pesawat ON 10 menit", ""),
            ("Lakukan eksposi", ""),
        ]),
        ("Ruangan", [
            ("Kebersihan ruangan", "Tidak terlihat sampah"),
            ("Keberihan meja kontrol & meja komputer", "Tidak berdebu & bernoda"),
            ("Bantal", "1 buah"),
            ("Suhu ruangan", "18 - 22 C"),
            ("Kelembapan Udara", "40 - 60 %"),
        ]),
        ("Mematikan Pesawat", [
            ("Turunkan semua parameter kV & mAs", ""),
            ("Posisikan tube vertikal", ""),
            ("Tekan tombol off pada kontrol pesawat", ""),
            ("Matikan panel listrik", ""),
        ]),
    ]),
    # --- form "QC MRI" ---
    _dev("MRI Ambition X", ruang="MRI Ambition X", sections=[
        ("Cek Kelengkapan & Peralatan Penunjang", [
            ("Control box / panel listrik", "Lampu indikator menyala"),
            ("Oksigen central", "Menyala / Berfungsi"),
            ("AC", "Menyala / Berfungsi"),
            ("Pintu ruangan", "Dapat terkunci rapat"),
            ("Suhu Chiller", "Menyala dan suhu ada pada batas aman"),
            ("Pengecekan UPS", "Cek kondisi baterai UPS, pastikan dalam kondisi normal dan tidak terdapat notifikasi error"),
        ]),
        ("Pemeriksaan Injektor", [
            ("Injektor ON", "Injektor dan CT Scan terhubung tanpa kendala koneksi"),
            ("Cek Fungsi Syringe", "Dapat mengisi & membuang cairan untuk kedua channel syringe"),
            ("Tes publik sistem", "Terdengar secara jelas"),
        ]),
        ("Warming Up MRI", [
            ("Pesawat & gantry ON", "Tidak terdapat notifikasi error pastikan terdapat suara kompressor"),
            ("Check up / Warming up Scan Phantom Botol", "Hasil scan normal dan tidak ada artefak"),
            ("Monitor worktation dan ISP", "Menyala / berfungsi"),
        ]),
        ("Kondisi Ruangan", [
            ("Kebersihan ruangan", "Ruangan bersih dan tidak terdapat sampah"),
            ("Kebersihan meja control & meja komputer", "Tidak terdapat debu & noda"),
            ("Suhu ruangan", "18 - 22 C"),
            ("Kelembapan Udara", "40 - 60 %"),
            ("Ketersediaan linen, bantal dan tissue gulung", "Tersedia secara lengkap"),
            ("Kelengkapan ALKES untuk tindakan", "Tersedia secara lengkap"),
            ("Selimut", "1 buah"),
            ("Bantal", "1 buah"),
        ]),
        ("Mematikan Pesawat dan Ambient Experience", [
            ("Mematikan MRI", "Pilih menu shutdown dan matikan sistem setiap malam hari"),
            ("Mematikan sistem ambient MRI", "Pilih menu shutdown dan matikan sistem setiap malam hari"),
        ]),
        ("Kondisi Changing Room No 1 dan No 2", [
            ("Kelengkapan baju dan celana pasien", "baju dan celana tersedia untuk seluruh ukuran"),
            ("Kebersihan Loker", "kondisi loker pasien bersih dan tidak terdapat sampah"),
            ("Kondisi smart lock loker", "Kondisi loker bisa dibuka dan ditutup menggunakan gelang"),
            ("Kondisi pintu loker", "Kondisi pintu loker bisa terkunci dan tertutup rapat"),
            ("Kelengkapan jumlah gelang smart lock", "Satu changing room memiliki 4 buah smart lock loker"),
        ]),
    ]),
]

# Penandatangan di bawah form QC (sesuai Excel): (judul kolom, nama, jabatan). Tanda tangan dibiarkan kosong untuk ditandatangani basah.
SIGNERS = [("Disiapkan Oleh", "Rheinner Nicholaus, Amd. Rad", "PIC. Fasilitas"),
           ("Mengetahui", "Joko Harjanto S.Tr.Rad", "Koordinator Radiologi"),
           ("Mengetahui", "dr Christopher Silman Sp.Rad, Ph.D", "Ho. Dept Radiologi")]

# Nama lengkap staff untuk pesan "Sudah diisi oleh ...". Inisial yang belum ada di sini ditampilkan apa adanya.
# Contoh:  STAFF_NAMES = {"AG": "Nama Lengkap", "HR": "Nama Lengkap"}
STAFF_NAMES = {}


def staff_label(ini):
    ini = (ini or "").strip()
    return STAFF_NAMES.get(ini, ini) or "-"


DEFAULT_CFG = {
    "rooms": [dict(nama=n, tekanan=False, t_lo=SUHU_NORMAL[0], t_hi=SUHU_NORMAL[1], h_lo=HUM_NORMAL[0], h_hi=HUM_NORMAL[1]) for n in _ROOMS],
    "staff": ["AG", "BA", "DT", "HR", "JK", "NA", "ND", "PR", "RH", "SS", "WN"],
    "devices": _DEVICES,
}




def now():
    return dt.datetime.now(TZ)


def TODAY():
    return now().date()


# =====================================================================
# 1. AUTO-DOWNLOAD FONT LEXEND
# =====================================================================
def ensure_lexend_fonts():
    reg_ok = os.path.exists(FONT_REG)
    bold_ok = os.path.exists(FONT_BOLD)
    for path, urls in FONT_URLS.items():
        if os.path.exists(path):
            continue
        for url in urls:
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=20) as r, open(path, "wb") as f:
                    f.write(r.read())
                if path == FONT_REG:
                    reg_ok = True
                if path == FONT_BOLD:
                    bold_ok = True
                break
            except Exception:
                continue
    return reg_ok, bold_ok


# =====================================================================
# 2. PENYIMPANAN (CSV + JSON)
# =====================================================================
ECOLS = ["id", "dibuat", "ruang", "tanggal", "shift", "suhu", "kelembapan", "petugas", "status", "temuan", "tindakan"]
QCOLS = ["id", "dibuat", "alat", "tanggal", "petugas", "hasil", "catatan"]


# ---------------------------------------------------------------------
# PENYIMPANAN
#  - Bila SUPABASE_URL dan SUPABASE_KEY diisi di Secrets  -> data disimpan di database Supabase (PERMANEN).
#  - Bila tidak diisi -> CSV lokal di folder data_phpk (HILANG saat Streamlit Cloud restart; hanya untuk uji coba).
# ---------------------------------------------------------------------
DB_CACHE_TTL = 20      # detik: isian dari HP lain muncul di dashboard paling lambat sekian detik
_STORAGE_ERRS = []     # (di-reset tiap run) agar pesan error tidak berulang


class StorageError(Exception):
    pass


def _secret(name, default=""):
    try:
        v = st.secrets.get(name)
    except Exception:
        v = None
    return str(v).strip() if v else os.environ.get(name, default)


def db_on():
    return bool(_secret("SUPABASE_URL") and _secret("SUPABASE_KEY"))


def _sb(method, path, params=None, body=None, prefer=None):
    """Panggilan REST ke Supabase (PostgREST). Kunci rahasia tidak pernah ikut di pesan error."""
    import requests
    key = _secret("SUPABASE_KEY")
    headers = {"apikey": key, "Content-Type": "application/json"}
    if key.startswith("eyJ"):   # kunci lama (JWT) boleh lewat Authorization; kunci baru sb_secret_... HANYA lewat apikey
        headers["Authorization"] = "Bearer " + key
    if prefer:
        headers["Prefer"] = prefer
    url = _secret("SUPABASE_URL").rstrip("/") + "/rest/v1/" + path
    try:
        r = requests.request(method, url, params=params, json=body, headers=headers, timeout=20)
    except Exception as e:
        raise StorageError("Tidak bisa menghubungi database ({})".format(type(e).__name__))
    if r.status_code >= 300:
        body = r.text.lstrip()
        if body[:1] == "<":   # HTML = bukan API Supabase -> SUPABASE_URL salah
            raise StorageError("SUPABASE_URL salah: server membalas halaman web, bukan API Supabase. "
                               "Isi dengan https://<project-ref>.supabase.co (tanpa /dashboard dan tanpa /rest/v1). Kode {}".format(r.status_code))
        raise StorageError("Database menolak permintaan ({}): {}".format(r.status_code, body[:200]))
    return r


def _ym(x):
    """Normalisasi ke 'YYYY-MM' dari str / date / (tahun, bulan)."""
    if x is None:
        return None
    if isinstance(x, str):
        return x[:7]
    if isinstance(x, (dt.date, dt.datetime)):
        return "{:04d}-{:02d}".format(x.year, x.month)
    return "{:04d}-{:02d}".format(int(x[0]), int(x[1]))


def _next_ym(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return "{:04d}-{:02d}".format(y, m)


@st.cache_data(ttl=DB_CACHE_TTL, show_spinner=False)
def _sb_fetch(table, ym=None, recent=0):
    """Ambil baris dari tabel (dibatasi 1 bulan bila ym diisi). Hasil di-cache beberapa detik agar cepat."""
    if recent:
        return _sb("GET", table, [("select", "*"), ("order", "id.desc"), ("limit", str(recent))]).json()
    rows, off = [], 0
    while True:   # PostgREST membatasi 1000 baris per permintaan -> ambil bertahap
        params = [("select", "*"), ("order", "id.asc"), ("limit", "1000"), ("offset", str(off))]
        if ym:
            params += [("tanggal", "gte." + ym + "-01"), ("tanggal", "lt." + _next_ym(ym) + "-01")]
        chunk = _sb("GET", table, params).json()
        rows += chunk
        if len(chunk) < 1000:
            return rows
        off += 1000


@st.cache_data(ttl=60, show_spinner=False)
def _sb_cfg():
    r = _sb("GET", "app_config", [("select", "value"), ("key", "eq.config")]).json()
    return r[0]["value"] if r else ""


def _warn_storage(msg):
    if msg not in _STORAGE_ERRS:
        _STORAGE_ERRS.append(msg)
        st.error("Penyimpanan data bermasalah. Hubungi admin." + ((" Detail: " + msg) if is_author() else ""))


def _to_df(rows, cols):
    df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols].fillna("").astype(str)


def _read(path, cols):
    if os.path.exists(path):
        try:
            df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
        except Exception:
            df = pd.DataFrame(columns=cols)
    else:
        df = pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


def _next_id(df):
    return str(1 + max([int(x) for x in df["id"] if str(x).isdigit()] + [0]))


def _load(table, path, cols, ym=None, recent=0):
    ym = _ym(ym)
    if db_on():
        try:
            return _to_df(_sb_fetch(table, ym, recent), cols)
        except StorageError as e:
            _warn_storage(str(e))
            return pd.DataFrame(columns=cols)
    df = _read(path, cols)
    if ym:
        df = df[df["tanggal"].str.startswith(ym)]
    if recent:
        df = df.tail(recent)
    return df


def load_entries(ym=None, recent=0):
    """ym = bulan yang dibutuhkan (date / (tahun, bulan) / 'YYYY-MM'); recent = N baris terbaru."""
    return _load("entries", ENTRY_FILE, ECOLS, ym, recent)


def load_qc(ym=None, recent=0):
    return _load("qc_entries", QC_FILE, QCOLS, ym, recent)


def _cfg_raw():
    try:
        if db_on():
            txt = _sb_cfg()
            return json.loads(txt) if txt else {}
        if os.path.exists(CFG_FILE):
            with open(CFG_FILE, encoding="utf-8") as f:
                return json.load(f)
    except StorageError as e:
        _warn_storage(str(e))
    except Exception:
        pass
    return {}


def save_cfg(c):
    """Simpan ruang, staff, alamat aplikasi. Daftar peralatan (devices) tetap dari kode, kecuali sudah ada manual di config."""
    out = {k: c[k] for k in ("rooms", "staff", "base_url") if k in c}
    old = _cfg_raw()
    if old.get("devices"):
        out["devices"] = old["devices"]
    if db_on():
        _sb("POST", "app_config", {"on_conflict": "key"}, [{"key": "config", "value": json.dumps(out, ensure_ascii=False)}],
            "resolution=merge-duplicates,return=minimal")
        _sb_cfg.clear()
        return
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CFG_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def load_cfg():
    cfg = json.loads(json.dumps(DEFAULT_CFG))
    c = _cfg_raw()
    for k in ("rooms", "staff", "devices", "base_url"):
        if c.get(k):
            cfg[k] = c[k]
    for r in cfg["rooms"]:   # batas suhu normal sama untuk semua ruangan (menimpa nilai lama yang mungkin tersimpan)
        r["t_lo"], r["t_hi"] = SUHU_NORMAL
    return cfg


def try_save_cfg(cfg):
    try:
        save_cfg(cfg)
        return True
    except StorageError as e:
        st.error("Gagal menyimpan: {}".format(e))
        return False


def rooms():
    return load_cfg()["rooms"]


def room_names():
    return [r["nama"] for r in rooms()]


def room_std(name):
    for r in rooms():
        if r["nama"] == name:
            return r
    return dict(nama=name, tekanan=False, t_lo=SUHU_NORMAL[0], t_hi=SUHU_NORMAL[1], h_lo=40.0, h_hi=60.0)


def devices():
    return load_cfg()["devices"]


def dev_by_name(name):
    for d in devices():
        if d["nama"] == name:
            return d
    return devices()[0]


def save_entry(**kw):
    """Simpan 1 pengukuran. Data dengan ruang+tanggal+shift yang sama ditimpa."""
    kw["dibuat"] = now().strftime("%Y-%m-%d %H:%M:%S")
    if db_on():
        row = {c: str(kw.get(c, "")) for c in ECOLS if c != "id"}
        _sb("POST", "entries", {"on_conflict": "ruang,tanggal,shift"}, [row], "resolution=merge-duplicates,return=minimal")
        _sb_fetch.clear()
        return
    df = _read(ENTRY_FILE, ECOLS)
    if len(df):
        m = (df["ruang"] == kw["ruang"]) & (df["tanggal"] == kw["tanggal"]) & (df["shift"] == kw["shift"])
        df = df[~m]
    row = {c: "" for c in ECOLS}
    for k, v in kw.items():
        row[k] = str(v)
    row["id"] = _next_id(df)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_csv(ENTRY_FILE, index=False)


def delete_entry(eid):
    if db_on():
        _sb("DELETE", "entries", {"id": "eq." + str(eid)})
        _sb_fetch.clear()
        return
    df = _read(ENTRY_FILE, ECOLS)
    df[df["id"] != str(eid)].to_csv(ENTRY_FILE, index=False)


def save_qc(alat, tanggal, petugas, hasil, catatan):
    """Simpan checklist QC 1 alat 1 hari. Data alat+tanggal yang sama ditimpa."""
    dibuat = now().strftime("%Y-%m-%d %H:%M:%S")
    hasil_txt = json.dumps(hasil, ensure_ascii=False)
    if db_on():
        _sb("POST", "qc_entries", {"on_conflict": "alat,tanggal"},
            [dict(dibuat=dibuat, alat=alat, tanggal=tanggal, petugas=petugas, hasil=hasil_txt, catatan=catatan)],
            "resolution=merge-duplicates,return=minimal")
        _sb_fetch.clear()
        return
    df = _read(QC_FILE, QCOLS)
    if len(df):
        df = df[~((df["alat"] == alat) & (df["tanggal"] == tanggal))]
    row = dict(id=_next_id(df), dibuat=dibuat, alat=alat, tanggal=tanggal, petugas=petugas, hasil=hasil_txt, catatan=catatan)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_csv(QC_FILE, index=False)


def delete_qc(eid):
    if db_on():
        _sb("DELETE", "qc_entries", {"id": "eq." + str(eid)})
        _sb_fetch.clear()
        return
    df = _read(QC_FILE, QCOLS)
    df[df["id"] != str(eid)].to_csv(QC_FILE, index=False)


def replace_table(table, path, df, cols, keys):
    """Restore dari CSV: mengganti SELURUH isi tabel."""
    df = df[cols].drop_duplicates(subset=keys, keep="last")
    if db_on():
        rows = [{c: str(v) for c, v in r.items() if c != "id"} for r in df.to_dict("records")]
        _sb("DELETE", table, {"id": "gt.0"})
        for i in range(0, len(rows), 500):
            _sb("POST", table, None, rows[i:i + 500], "return=minimal")
        _sb_fetch.clear()
        return len(rows)
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_csv(path, index=False)
    return len(df)


def storage_test():
    """(ok, pesan) untuk kotak status penyimpanan di sidebar author."""
    if not db_on():
        return False, "Mode CSV lokal - data HILANG saat aplikasi restart. Isi SUPABASE_URL dan SUPABASE_KEY di Secrets."
    try:
        for t in ("entries", "qc_entries", "app_config"):
            _sb("GET", t, [("select", "*"), ("limit", "1")])
        return True, "Terhubung ke Supabase - 3 tabel siap. Data tersimpan permanen."
    except StorageError as e:
        return False, str(e)


def qc_get(alat, tgl):
    df = load_qc(ym=tgl)
    df = df[(df["alat"] == alat) & (df["tanggal"] == tgl.isoformat())]
    if not len(df):
        return None
    r = df.iloc[-1].to_dict()
    try:
        hasil = json.loads(r["hasil"]) if r["hasil"] else {}
    except Exception:
        hasil = {}
    return dict(petugas=r["petugas"], done=set(k for k, v in hasil.items() if v), catatan=r["catatan"], id=r["id"])


# =====================================================================
# 2b. QR CODE PER RUANGAN
# =====================================================================
def default_base_url():
    cfg = load_cfg()
    if cfg.get("base_url"):
        return cfg["base_url"]
    try:   # alamat aplikasi yang sedang dibuka (Streamlit >= 1.37)
        host = st.context.headers.get("Host")
        if host:
            return ("http://" if host.startswith(("localhost", "127.")) else "https://") + host
    except Exception:
        pass
    return "https://NAMA-APP.streamlit.app"


def room_url(base, ruang):
    return "{}/ruang?ruang={}".format(base.rstrip("/"), urllib.parse.quote(ruang))


def room_devices(ruang):
    """Peralatan yang checklist QC-nya diisi lewat QR ruangan ini."""
    return [d for d in devices() if (d.get("ruang") or d.get("qr_ruang")) == ruang]


def shift_now():
    h = now().hour
    return "P" if h < 11 else ("S" if h < 18 else "M")


def qr_matrix(text):
    from reportlab.graphics.barcode.qrencoder import QRCode, QRErrorCorrectLevel
    q = QRCode(None, QRErrorCorrectLevel.M)
    q.addData(text)
    q.make()
    n = q.getModuleCount()
    return [[bool(q.isDark(r, c)) for c in range(n)] for r in range(n)]


def fnum(v):
    return "%g" % v


# =====================================================================
# 3. DATA BULANAN + ANALISA : SUHU & KELEMBABAN
# =====================================================================
def slot_due(d, s):
    n = now()
    if d < n.date():
        return True
    if d > n.date():
        return False
    return n.hour >= SHIFT_HOUR[s]


def month_data(ruang, y, m):
    """{(tanggal, 'P'|'S'|'M'): dict(suhu, hum, petugas, ket, id)}"""
    out = {}
    df = load_entries(ym=(y, m))
    df = df[df["ruang"] == ruang]
    for r in df.to_dict("records"):
        try:
            td = dt.date.fromisoformat(r["tanggal"])
            if td.year != y or td.month != m:
                continue
            ket = r["temuan"] + (" (" + r["tindakan"] + ")" if r["tindakan"] and r["temuan"] else "")
            out[(td.day, r["shift"][:1])] = dict(suhu=float(r["suhu"]), hum=float(r["kelembapan"]),
                                                 petugas=r["petugas"], ket=ket, id=r["id"])
        except Exception:
            continue
    return out


def classify(v, lo, hi):
    return "ok" if lo <= v <= hi else ("high" if v > hi else "low")


def analyze(ruang, y, m):
    std = room_std(ruang)
    data = month_data(ruang, y, m)
    n = calendar.monthrange(y, m)[1]
    slots, empty_by_day = {}, {}
    high, low, hum_bad = [], [], []
    due = filled = filled_due = ok = hum_ok = 0
    for d in range(1, n + 1):
        for s in "PSM":
            is_due = slot_due(dt.date(y, m, d), s)
            e = data.get((d, s))
            if e is None:
                slots[(d, s)] = dict(state="empty" if is_due else "future")
                if is_due:
                    due += 1
                    empty_by_day.setdefault(d, []).append(s)
                continue
            st_ = classify(e["suhu"], std["t_lo"], std["t_hi"])
            hs = classify(e["hum"], std["h_lo"], std["h_hi"])
            slots[(d, s)] = dict(state=st_, t=e["suhu"], h=e["hum"], hstate=hs, petugas=e["petugas"])
            if is_due:
                due += 1
                filled_due += 1
            filled += 1
            ok += 1 if st_ == "ok" else 0
            hum_ok += 1 if hs == "ok" else 0
            if st_ == "high":
                high.append((d, s, e["suhu"]))
            elif st_ == "low":
                low.append((d, s, e["suhu"]))
            if hs != "ok":
                hum_bad.append((d, s, e["hum"], hs))
    full_empty = [d for d, v in empty_by_day.items() if len(v) == 3]
    return dict(std=std, n=n, slots=slots, empty_by_day=empty_by_day, high=high, low=low, hum_bad=hum_bad, due=due, filled=filled,
                filled_due=filled_due, ok=ok, hum_ok=hum_ok, pct_fill=(filled_due / due * 100 if due else 0.0), pct_ok=(ok / filled * 100 if filled else 0.0),
                pct_hum=(hum_ok / filled * 100 if filled else 0.0), full_empty=sorted(full_empty), n_empty=sum(len(v) for v in empty_by_day.values()))


def abnormal_today():
    """Pengukuran HARI INI (semua ruang) yang di luar batas -> perlu hubungi teknisi."""
    df = load_entries(ym=TODAY())
    df = df[df["tanggal"] == TODAY().isoformat()]
    bad = []
    for r in df.to_dict("records"):
        try:
            std = room_std(r["ruang"])
            sv, hv = float(r["suhu"]), float(r["kelembapan"])
        except Exception:
            continue
        ts, hs = classify(sv, std["t_lo"], std["t_hi"]), classify(hv, std["h_lo"], std["h_hi"])
        if ts != "ok" or hs != "ok":
            parts = []
            if ts != "ok":
                parts.append("suhu {:g}°C ({})".format(sv, "di atas" if ts == "high" else "di bawah"))
            if hs != "ok":
                parts.append("kelembapan {:g}% ({})".format(hv, "di atas" if hs == "high" else "di bawah"))
            bad.append("<b>{}</b> - {}: {}".format(escape(r["ruang"]), SHIFT_NAME.get(r["shift"][:1], r["shift"]), ", ".join(parts)))
    return bad


# =====================================================================
# 4. DATA BULANAN + ANALISA : QC PERALATAN
# =====================================================================
def is_daily(item):
    return "mingguan" not in item.lower()


def qc_month(dev, y, m):
    """{tanggal: dict(petugas, done=set(item), catatan)}"""
    out = {}
    df = load_qc(ym=(y, m))
    df = df[df["alat"] == dev["nama"]]
    for r in df.to_dict("records"):
        try:
            td = dt.date.fromisoformat(r["tanggal"])
            if td.year != y or td.month != m:
                continue
            hasil = json.loads(r["hasil"]) if r["hasil"] else {}
            out[td.day] = dict(petugas=r["petugas"], done=set(k for k, v in hasil.items() if v), catatan=r["catatan"], id=r["id"])
        except Exception:
            continue
    return out


def qc_analyze(dev, y, m):
    data = qc_month(dev, y, m)
    n = calendar.monthrange(y, m)[1]
    daily = [i for i in dev["items"] if is_daily(i)]
    n_now = now()
    days, missing = {}, []
    counted = full = filled = item_done = item_total = 0
    for d in range(1, n + 1):
        dd = dt.date(y, m, d)
        is_due = dd < n_now.date() or (dd == n_now.date() and n_now.hour >= QC_DUE_HOUR)
        e = data.get(d)
        miss = [i for i in daily if not (e and i in e["done"])]
        dn = len(daily) - len(miss)
        if dn == len(daily):
            state = "full"
        elif dn > 0:
            state = "part"
        else:
            state = "empty" if is_due else "future"
        days[d] = dict(state=state, done=dn, total=len(daily))
        if is_due or dn > 0:
            counted += 1
            item_total += len(daily)
            item_done += dn
            if state == "full":
                full += 1
            if dn > 0:
                filled += 1
            if state in ("part", "empty"):
                missing.append((d, state, miss))
    return dict(days=days, n=n, counted=counted, full=full, filled=filled, item_done=item_done, item_total=item_total,
                pct_full=(full / counted * 100 if counted else 0.0), pct_item=(item_done / item_total * 100 if item_total else 0.0),
                missing=missing, n_empty=sum(1 for _d, s, _m in missing if s == "empty"),
                n_part=sum(1 for _d, s, _m in missing if s == "part"))


# =====================================================================
# 5. GAMBAR FORMULIR (vektor: dipakai untuk preview SVG & PDF)
# =====================================================================
BLACK, RED, WHITE = "#000000", "#ff0000", "#ffffff"
NAVY = "#073763"        # warna bar gelap di form suhu (sama dengan Excel)
C_GREY = "#d9d9d9"      # pita judul "SUHU" / "KELEMBAPAN" & kotak keterangan
C_DAY = "#ccffff"       # kolom tanggal genap (form suhu)
C_OUT = "#f4cccc"       # label angka di luar target (form suhu)
C_NAME = "#c9daf8"      # baris NAMA (form suhu)
C_HEAD = "#efefef"      # header NO / KEGIATAN / PARAMETER / tanggal (form QC)
C_BAND = "#9cc2e5"      # pita judul seksi (form QC)
C_NAMEQC = "#cccccc"    # baris NAMA PEKERJA RADIASI (form QC)


def _png_dimensions(path):
    try:
        with open(path, "rb") as f:
            head = f.read(24)
        if len(head) < 24 or head[:8] != b"\x89PNG\r\n\x1a\n":
            return None, None
        return int.from_bytes(head[16:20], "big"), int.from_bytes(head[20:24], "big")
    except Exception:
        return None, None


def _logo_b64(path=LOGO_FILE):
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


def _fit_box(src_w, src_h, dst_w, dst_h):
    if not src_w or not src_h:
        return dst_w, dst_h
    scale = min(dst_w / src_w, dst_h / src_h)
    return src_w * scale, src_h * scale


class Sheet:
    """Daftar operasi gambar dalam 'unit' (y ke bawah). width/height diisi builder."""

    def __init__(self, width=1000.0, height=700.0):
        self.ops = []
        self.width = width
        self.height = height

    def rect(self, x, y, w, h, fill=None, stroke=None, lw=1.0):
        self.ops.append(("rect", x, y, w, h, fill, stroke, lw))

    def line(self, x1, y1, x2, y2, color=BLACK, lw=1.0):
        self.ops.append(("line", x1, y1, x2, y2, color, lw))

    def text(self, x, y, s, size=8.0, bold=False, color=BLACK, anchor="start", rot=0):
        self.ops.append(("text", x, y, str(s), size, bold, color, anchor, rot))

    def circle(self, cx, cy, r, fill=None, stroke=None, lw=1.0):
        self.ops.append(("circle", cx, cy, r, fill, stroke, lw))

    def image(self, x, y, w, h, data_b64):
        self.ops.append(("image", x, y, w, h, data_b64))

    def ctext(self, x, y, w, h, s, size=8.0, bold=False, color=BLACK):
        lines = str(s).split("\n")
        lh = size * 1.18
        base = y + (h - lh * len(lines)) / 2 + size * 0.82
        for i, ln in enumerate(lines):
            self.text(x + w / 2, base + i * lh, ln, size, bold, color, "middle")

    def ltext(self, x, y, h, s, size=8.0, bold=False, color=BLACK):
        """Teks rata kiri, rata tengah secara vertikal dalam kotak setinggi h."""
        lines = str(s).split("\n")
        lh = size * 1.18
        base = y + (h - lh * len(lines)) / 2 + size * 0.82
        for i, ln in enumerate(lines):
            self.text(x, base + i * lh, ln, size, bold, color, "start")

    def check(self, cx, cy, color=BLACK, lw=1.2):
        self.line(cx - 3.4, cy + 0.2, cx - 1.1, cy + 2.9, color, lw)
        self.line(cx - 1.1, cy + 2.9, cx + 3.6, cy - 3.0, color, lw)


def _logo_into(sh, x, y, box_w, box_h, path=LOGO_FILE, align="left"):
    """Tempel logo (rasio asli dijaga) di dalam kotak box_w x box_h."""
    logo = _logo_b64(path)
    if logo:
        sw, shh = _png_dimensions(path)
        fw, fh = _fit_box(sw, shh, box_w, box_h)
        ox = x if align == "left" else (x + box_w - fw if align == "right" else x + (box_w - fw) / 2)
        sh.image(ox, y + (box_h - fh) / 2, fw, fh, logo)
    elif path == LOGO_FILE:   # cadangan bila file logo belum ada
        sh.text(x, y + box_h * 0.55, "Mandaya", 15, True, "#19558c")
        sh.text(x, y + box_h * 0.95, "Royal Hospital PURI", 9, True, "#19558c")


def _tw(s, size, bold=False):
    """Perkiraan lebar teks (unit gambar) memakai metrik Lexend bila ada; bila tidak, Helvetica x 1,18 (Lexend lebih lebar)."""
    try:
        from reportlab.pdfbase import pdfmetrics
        reg_ok, bold_ok = _register_lexend()
        if reg_ok:
            return pdfmetrics.stringWidth(s, "Lexend-Bold" if (bold and bold_ok) else "Lexend", size)
        return pdfmetrics.stringWidth(s, "Helvetica-Bold" if bold else "Helvetica", size) * 1.18
    except Exception:
        return len(s) * size * 0.62


def _wrap(s, size, maxw, bold=False):
    """Pecah teks menjadi beberapa baris agar muat selebar maxw."""
    lines, cur = [], ""
    for w in str(s).split():
        t = (cur + " " + w).strip()
        if cur and _tw(t, size, bold) > maxw:
            lines.append(cur)
            cur = w
        else:
            cur = t
    if cur:
        lines.append(cur)
    return lines or [""]


# ---------- FORM 1: FORM DIGITAL MONITORING SUHU DAN KELEMBAPAN ----------
# Mengikuti sheet per-ruangan di "FORM MONITORING SUHU 2026.xlsx": tabel SUHU (30..15 °C), tabel KELEMBAPAN (65..30 %),
# 31 tanggal x 3 shift (P/S/M), titik hitam = dalam target, titik merah = di luar target, baris NAMA di bawah.
TH_TEMPS = list(range(30, 14, -1))        # baris suhu: 30 ... 15
TH_HUMS = list(range(65, 29, -5))         # baris kelembapan: 65, 60, ... 30


def build_th_form(ruang, y, m, data):
    std = room_std(ruang)
    tlo, thi, hlo, hhi = std["t_lo"], std["t_hi"], std["h_lo"], std["h_hi"]
    W = 1100.0
    MX, LBL = 22.0, 36.0
    SB = (W - 2 * MX - LBL) / 93.0             # lebar 1 kolom shift (P / S / M)
    DW = SB * 3                                # lebar 1 tanggal
    XD = MX + LBL                              # awal kolom tanggal 1
    XE = XD + 31 * DW                          # akhir kolom tanggal 31
    RH = 17.0                                  # tinggi baris angka
    HB, HT, HD, HBAR, HN = 15.0, 15.0, 13.0, 7.0, 17.0
    sh = Sheet(W, 10)

    # ---- kepala form: logo kiri, judul tengah ----
    _logo_into(sh, MX, 8, 190, 62)
    sh.text(W / 2, 36, "Form Digital Monitoring Suhu dan Kelembapan", 19, True, BLACK, "middle")
    sh.text(W / 2, 60, "Departemen Radiologi Tahun {}".format(y), 19, True, BLACK, "middle")
    sh.text(MX, 92, "RUANG : {}".format(ruang), 11.5, True)
    sh.text(MX, 110, "BULAN : {} {}".format(BLN[m - 1].upper(), y), 10.5, True)

    def header_rows(y0, judul, target):
        """Pita judul + baris Tgl + baris Dinas. Mengembalikan y awal baris angka."""
        xs = XD + 8 * DW                        # pita 'judul' di 8 tanggal pertama, 'target' di sisanya (seperti di Excel)
        sh.rect(MX, y0, xs - MX, HB, C_GREY, BLACK, 0.5)
        sh.ctext(MX, y0, xs - MX, HB, judul, 8.5, True)
        sh.rect(xs, y0, XE - xs, HB, C_GREY, BLACK, 0.5)
        sh.ctext(xs, y0, XE - xs, HB, target, 8.5, True)
        yt, yd = y0 + HB, y0 + HB + HT
        sh.rect(MX, yt, LBL, HT, WHITE, BLACK, 0.5)
        sh.ctext(MX, yt, LBL, HT, "Tgl", 7.5, True)
        sh.rect(MX, yd, LBL, HD, WHITE, BLACK, 0.5)
        sh.ctext(MX, yd, LBL, HD, "Dinas", 7, True)
        for d in range(1, 32):
            x = XD + (d - 1) * DW
            fill = C_DAY if d % 2 == 0 else WHITE
            sh.rect(x, yt, DW, HT, fill, BLACK, 0.5)
            sh.ctext(x, yt, DW, HT, str(d), 7.5, True)
            for k, ch in enumerate("PSM"):
                sh.rect(x + k * SB, yd, SB, HD, fill, BLACK, 0.4)
                sh.ctext(x + k * SB, yd, SB, HD, ch, 6.2, True)
        return yd + HD

    def grid(y0, values, lo, hi):
        """Baris angka: label di kiri (merah muda bila di luar target), sel putih / biru muda per tanggal."""
        for i, v in enumerate(values):
            yy = y0 + i * RH
            out = v < lo or v > hi
            sh.rect(MX, yy, LBL, RH, C_OUT if out else WHITE, BLACK, 0.5)
            sh.ctext(MX, yy, LBL, RH, str(v), 8, True, RED if out else BLACK)
            for d in range(1, 32):
                x = XD + (d - 1) * DW
                fill = C_DAY if d % 2 == 0 else WHITE
                for k in range(3):
                    sh.rect(x + k * SB, yy, SB, RH, fill, BLACK, 0.4)

    def dots(y0, values, key, lo, hi):
        top, step = values[0], values[0] - values[1]
        for (d, s), e in data.items():
            v = e[key]
            k = "PSM".index(s)
            cx = XD + (d - 1) * DW + (k + 0.5) * SB
            cy = y0 + (top - v) / step * RH + RH / 2
            cy = min(max(cy, y0 + 3.5), y0 + len(values) * RH - 3.5)
            sh.circle(cx, cy, 3.4, RED if (v < lo or v > hi) else BLACK)

    # ---- tabel SUHU ----
    y1 = 124.0
    g1 = header_rows(y1, "SUHU", "Target Temperatur ({:g} - {:g})°C".format(tlo, thi))
    grid(g1, TH_TEMPS, tlo, thi)
    dots(g1, TH_TEMPS, "suhu", tlo, thi)
    yb = g1 + len(TH_TEMPS) * RH
    sh.rect(MX, yb, XE - MX, HBAR, NAVY, None, 0)

    # ---- tabel KELEMBAPAN ----
    y2 = yb + HBAR + 2
    g2 = header_rows(y2, "KELEMBAPAN", "Target kelembapan ( {:g} - {:g} % )".format(hlo, hhi))
    grid(g2, TH_HUMS, hlo, hhi)
    dots(g2, TH_HUMS, "hum", hlo, hhi)
    yn = g2 + len(TH_HUMS) * RH

    # ---- baris NAMA (inisial petugas per shift) ----
    sh.rect(MX, yn, LBL, HN, C_NAME, BLACK, 0.5)
    sh.ctext(MX, yn, LBL, HN, "NAMA", 7.5, True)
    for d in range(1, 32):
        for k, s in enumerate("PSM"):
            x = XD + (d - 1) * DW + k * SB
            sh.rect(x, yn, SB, HN, C_NAME, BLACK, 0.4)
            e = data.get((d, s))
            if e and e["petugas"]:
                p = e["petugas"]
                sh.ctext(x, yn, SB, HN, p, 5.6 if len(p) <= 2 else 4.2, True)
    ye = yn + HN
    sh.rect(MX, ye, XE - MX, 12.0, NAVY, None, 0)
    yend = ye + 12.0

    # ---- keterangan temuan (hanya muncul bila ada temuan di luar target) ----
    notes = []
    for (d, s), e in sorted(data.items()):
        if e.get("ket"):
            notes.extend(_wrap("Tgl {} ({}): {}".format(d, SHIFT_NAME[s], e["ket"]), 7.5, XE - MX - 30))
    if notes:
        yy = yend + 16
        sh.text(MX, yy, "Keterangan temuan:", 8, True)
        for ln in notes[:8]:
            yy += 10
            sh.text(MX + 8, yy, ln, 7.5)
        yend = yy
    sh.height = yend + 14.0
    return sh


# ---------- FORM 2: FORM DIGITAL DAILY QUALITY CONTROL ----------
# Mengikuti sheet "QC ..." di "Form Digital Daily Quality Control Tahun 2026.xlsx": NO | KEGIATAN | PARAMETER | tanggal 1..31,
# pita biru per seksi, baris NAMA PEKERJA RADIASI, kotak keterangan, dan tabel tanda tangan.
def dev_sections(dev):
    """Seksi form QC. Bila konfigurasi alat tersimpan dari versi lama (tanpa 'sections'), jadikan satu seksi tanpa judul."""
    if dev.get("sections"):
        return dev["sections"]
    return [dict(title="", items=[dict(key=i, no=str(n), kegiatan=i, param="") for n, i in enumerate(dev["items"], 1)])]


def item_label(key):
    """Kunci item 'Seksi::Kegiatan' -> teks yang enak dibaca."""
    return key.replace("::", " / ") if "::" in key else key


def build_qc_form(dev, y, m, data):
    X0, NOW, KEGW, PARW, DAYW = 14.0, 28.0, 172.0, 218.0, 15.5
    XK = X0 + NOW                  # awal kolom KEGIATAN
    XP = XK + KEGW                 # awal kolom PARAMETER
    XD = XP + PARW                 # awal kolom tanggal 1
    XE = XD + 31 * DAYW
    W = XE + 14.0
    FS, LH, PAD = 7.2, 8.8, 3.6    # ukuran huruf isi, tinggi baris teks, padding vertikal
    sh = Sheet(W, 10)

    # ---- kepala form ----
    _logo_into(sh, X0, 8, 190, 62)
    sh.text(W / 2, 36, "Form Digital Daily Quality Control", 18, True, BLACK, "middle")
    sh.text(W / 2, 59, "Departemen Radiologi Tahun {}".format(y), 18, True, BLACK, "middle")
    sh.text(X0, 90, "Modalitas : {}".format(dev["info"][0][1]), 10.5, True)
    sh.text(X0, 106, "BULAN : {} {}".format(BLN[m - 1].upper(), y), 10.5, True)

    # ---- header tabel ----
    yh, HH = 114.0, 19.0
    for x, w, t in ((X0, NOW, "NO"), (XK, KEGW, "KEGIATAN"), (XP, PARW, "PARAMETER")):
        sh.rect(x, yh, w, HH, C_HEAD, BLACK, 0.6)
        sh.ctext(x, yh, w, HH, t, 8.5, True)
    for d in range(1, 32):
        x = XD + (d - 1) * DAYW
        sh.rect(x, yh, DAYW, HH, C_HEAD, BLACK, 0.6)
        sh.ctext(x, yh, DAYW, HH, str(d), 7.4, True)

    # ---- seksi + butir ----
    yy = yh + HH
    BANDH = 15.0
    for sec in dev_sections(dev):
        if sec["title"]:
            sh.rect(X0, yy, XE - X0, BANDH, C_BAND, BLACK, 0.6)
            sh.ctext(X0, yy, XE - X0, BANDH, sec["title"].strip(), 8.2, True)
            yy += BANDH
        for it in sec["items"]:
            kl = _wrap(it["kegiatan"], FS, KEGW - 8)
            pl = _wrap(it["param"], FS, PARW - 8) if it["param"] else [""]
            rh = max(15.0, max(len(kl), len(pl)) * LH + PAD * 2)
            sh.rect(X0, yy, NOW, rh, None, BLACK, 0.6)
            sh.ctext(X0, yy, NOW, rh, it["no"], FS + 0.4)
            sh.rect(XK, yy, KEGW, rh, None, BLACK, 0.6)
            sh.ltext(XK + 4, yy, rh, "\n".join(kl), FS)
            sh.rect(XP, yy, PARW, rh, None, BLACK, 0.6)
            if it["param"]:
                sh.ltext(XP + 4, yy, rh, "\n".join(pl), FS)
            for d in range(1, 32):
                x = XD + (d - 1) * DAYW
                sh.rect(x, yy, DAYW, rh, None, BLACK, 0.6)
                e = data.get(d)
                if e and it["key"] in e["done"]:
                    sh.check(x + DAYW / 2, yy + rh / 2)
            yy += rh

    # ---- NAMA PEKERJA RADIASI ----
    NH = 19.0
    sh.rect(X0, yy, XD - X0, NH, C_NAMEQC, BLACK, 0.6)
    sh.ctext(X0, yy, XD - X0, NH, "NAMA PEKERJA RADIASI", 8.5, True)
    for d in range(1, 32):
        x = XD + (d - 1) * DAYW
        sh.rect(x, yy, DAYW, NH, C_NAMEQC, BLACK, 0.6)
        e = data.get(d)
        if e and e["petugas"]:
            p = e["petugas"]
            sh.ctext(x, yy, DAYW, NH, p, 7.4 if len(p) <= 2 else 5.2, True)
    yy += NH + 14

    # ---- catatan harian (bila ada) ----
    notes = []
    for d in sorted(data):
        if data[d].get("catatan"):
            notes.extend(_wrap("Tgl {}: {}".format(d, data[d]["catatan"]), 7.5, XE - X0 - 20))
    if notes:
        sh.text(X0, yy, "Catatan:", 8, True)
        yy += 10
        for ln in notes[:6]:
            sh.text(X0 + 8, yy, ln, 7.5)
            yy += 10
        yy += 6

    # ---- kotak keterangan (kiri) + tabel tanda tangan (kanan) ----
    ybox = yy
    LGW, LGH = XD - X0 + 20, 58.0
    sh.rect(X0, ybox, LGW, LGH, C_GREY, BLACK, 0.6)
    sh.text(X0 + 6, ybox + 14, "Keterangan :", 7.8, False)
    sh.check(X0 + 12, ybox + 28)
    sh.text(X0 + 22, ybox + 30, ": Memenuhi / Lulus Uji / Lengkap", 7.8)
    sh.text(X0 + 8, ybox + 44, "X", 8.2, True)
    sh.text(X0 + 22, ybox + 44, ": Tidak Memenuhi / Tidak Lulus Uji / Tidak Lengkap", 7.8)

    SGX = XD + 22.0                              # tabel tanda tangan dimulai di kolom tanggal, lebar 3 blok sama besar
    SGW = (XE - SGX) / 3.0
    HS, HSIG, HNM, HRL = 15.0, 48.0, 15.0, 15.0
    i = 0
    while i < len(SIGNERS):                      # judul kolom yang sama berurutan digabung (mis. 2 kolom "Mengetahui")
        j = i
        while j + 1 < len(SIGNERS) and SIGNERS[j + 1][0] == SIGNERS[i][0]:
            j += 1
        x = SGX + i * SGW
        w = (j - i + 1) * SGW
        sh.rect(x, ybox, w, HS, C_GREY, BLACK, 0.6)
        sh.ctext(x, ybox, w, HS, SIGNERS[i][0], 8, True)
        i = j + 1
    for i, (_jud, nama, role) in enumerate(SIGNERS):
        x = SGX + i * SGW
        sh.rect(x, ybox + HS, SGW, HSIG, None, BLACK, 0.6)
        sh.rect(x, ybox + HS + HSIG, SGW, HNM, None, BLACK, 0.6)
        fs = 7.4
        while _tw(nama, fs) > SGW - 8 and fs > 5.2:
            fs -= 0.2
        sh.ctext(x, ybox + HS + HSIG, SGW, HNM, nama, fs)
        sh.rect(x, ybox + HS + HSIG + HNM, SGW, HRL, None, BLACK, 0.6)
        sh.ctext(x, ybox + HS + HSIG + HNM, SGW, HRL, role, 7.4)
    sh.height = ybox + max(LGH, HS + HSIG + HNM + HRL) + 14.0
    return sh


# ---------- LEMBAR QR CODE RUANGAN (stiker, 6 per halaman A4 portrait) ----------
def build_qr_sheets(base):
    CW, CH = 240.0, 282.0
    cards = []
    for rn in room_names():
        devs = [d["nama"] for d in room_devices(rn)]
        cards.append((rn, room_url(base, rn), devs))
    sheets = []
    for p0 in range(0, len(cards), 6):
        chunk = cards[p0:p0 + 6]
        rows = max(1, math.ceil(len(chunk) / 2))
        sh = Sheet(2 * CW, rows * CH)
        for i, (rn, url, devs) in enumerate(chunk):
            x, y = (i % 2) * CW, (i // 2) * CH
            sh.rect(x + 3, y + 3, CW - 6, CH - 6, None, "#9e9e9e", 0.6)
            _logo_into(sh, x + CW / 2 - 60, y + 8, 120, 38, LOGO_FILE, "center")
            sh.text(x + CW / 2, y + 56, "DEPARTEMEN RADIOLOGI", 8, True, "#555555", "middle")
            nm = _wrap(rn, 19, CW - 24, True)
            if len(nm) == 1:
                sh.text(x + CW / 2, y + 79, nm[0], 19 if _tw(nm[0], 19, True) <= CW - 24 else 14, True, BLACK, "middle")
            else:   # nama panjang -> 2 baris lebih kecil
                nm = _wrap(rn, 14, CW - 24, True)[:2]
                for k, ln in enumerate(nm):
                    sh.text(x + CW / 2, y + (79 if len(nm) == 1 else 74) + k * 15, ln, 14, True, BLACK, "middle")
            mat = qr_matrix(url)
            n = len(mat)
            size = 128.0
            qx, qy = x + (CW - size) / 2, y + 92
            sh.rect(qx - 7, qy - 7, size + 14, size + 14, "#ffffff", "#bdbdbd", 0.5)
            cell = size / n
            for r in range(n):   # gabungkan modul gelap berurutan jadi satu persegi panjang
                c = 0
                while c < n:
                    if mat[r][c]:
                        c2 = c
                        while c2 + 1 < n and mat[r][c2 + 1]:
                            c2 += 1
                        sh.rect(qx + c * cell, qy + r * cell, (c2 - c + 1) * cell, cell, BLACK, None, 0)
                        c = c2 + 1
                    else:
                        c += 1
            sh.text(x + CW / 2, y + 243, "SCAN UNTUK MENGISI", 9.5, True, BLACK, "middle")
            sub = "Suhu & Kelembapan" + (" + QC " + ", ".join(devs) if devs else "")
            for k, ln in enumerate(_wrap(sub, 7, CW - 24)[:2]):
                sh.text(x + CW / 2, y + 254 + k * 8.5, ln, 7, False, "#333333", "middle")
            sh.text(x + CW / 2, y + 276, url if len(url) < 70 else url[:67] + "...", 4.8, False, "#888888", "middle")
        sheets.append(sh)
    return sheets



# ---------- konverter: SVG (preview) & PDF ----------
def _font_css():
    css = ""
    for path, wgt in ((FONT_REG, 400), (FONT_BOLD, 700)):
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    b = base64.b64encode(f.read()).decode()
                css += "@font-face{{font-family:'LexendEmbed';font-style:normal;font-weight:{};src:url(data:font/ttf;base64,{}) format('truetype');}}".format(wgt, b)
            except Exception:
                pass
    return css


def sheet_to_svg(sh, width_px=1400):
    W, H = sh.width, sh.height
    p = ["<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 {:.1f} {:.1f}' width='{}' height='{}' font-family=\"LexendEmbed, Lexend, Helvetica, Arial, sans-serif\">".format(
        W, H, width_px, int(width_px * H / W))]
    css = _font_css()
    if css:
        p.append("<defs><style type='text/css'>{}</style></defs>".format(css))
    p.append("<rect x='0' y='0' width='{:.1f}' height='{:.1f}' fill='#ffffff'/>".format(W, H))
    for op in sh.ops:
        k = op[0]
        if k == "rect":
            _, x, y, w, h, fill, stroke, lw = op
            p.append("<rect x='{:.2f}' y='{:.2f}' width='{:.2f}' height='{:.2f}' fill='{}' stroke='{}' stroke-width='{}'/>".format(
                x, y, w, h, fill or "none", stroke or "none", lw))
        elif k == "line":
            _, x1, y1, x2, y2, c, lw = op
            p.append("<line x1='{:.2f}' y1='{:.2f}' x2='{:.2f}' y2='{:.2f}' stroke='{}' stroke-width='{}' stroke-linecap='round'/>".format(x1, y1, x2, y2, c, lw))
        elif k == "circle":
            _, cx, cy, r, fill, stroke, lw = op
            p.append("<circle cx='{:.2f}' cy='{:.2f}' r='{}' fill='{}' stroke='{}' stroke-width='{}'/>".format(cx, cy, r, fill or "none", stroke or "none", lw))
        elif k == "image":
            _, x, y, w, h, b64 = op
            p.append("<image x='{:.2f}' y='{:.2f}' width='{:.2f}' height='{:.2f}' preserveAspectRatio='xMidYMid meet' href='data:image/png;base64,{}'/>".format(x, y, w, h, b64))
        else:
            _, x, y, s, size, bold, color, anchor, rot = op
            tf = " transform='rotate(-{} {:.2f} {:.2f})'".format(rot, x, y) if rot else ""
            p.append("<text x='{:.2f}' y='{:.2f}' font-size='{}' font-weight='{}' fill='{}' text-anchor='{}'{}>{}</text>".format(
                x, y, size, "bold" if bold else "normal", color, anchor, tf, escape(s)))
    return "".join(p) + "</svg>"


LEXEND_STATUS = {"reg": False, "bold": False, "tried": False}


def _register_lexend():
    if LEXEND_STATUS["tried"]:
        return LEXEND_STATUS["reg"], LEXEND_STATUS["bold"]
    ensure_lexend_fonts()
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    ok = {}
    for key, name, path in (("reg", "Lexend", FONT_REG), ("bold", "Lexend-Bold", FONT_BOLD)):
        ok[key] = False
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path))
                ok[key] = True
            except Exception:
                pass
    LEXEND_STATUS.update(reg=ok["reg"], bold=ok["bold"], tried=True)
    return ok["reg"], ok["bold"]


def sheets_to_pdf(sheets, title="Formulir Radiologi - " + HOSPITAL, portrait=None):
    """portrait=None -> tiap halaman memilih sendiri A4 landscape / portrait mana yang membuat formulir tercetak lebih besar."""
    from reportlab.lib.colors import HexColor
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.utils import ImageReader

    reg_ok, bold_ok = _register_lexend()
    fn_reg = "Lexend" if reg_ok else "Helvetica"
    fn_bold = "Lexend-Bold" if bold_ok else "Helvetica-Bold"
    buf = io.BytesIO()
    MG = 10.0
    c = Canvas(buf, pagesize=landscape(A4))
    c.setTitle(title)
    c.setAuthor(HOSPITAL + " - Departemen Radiologi")

    for sh in sheets:
        mg = getattr(sh, "margin", MG)
        if portrait is None:
            s_l = min((landscape(A4)[0] - 2 * mg) / sh.width, (landscape(A4)[1] - 2 * mg) / sh.height)
            s_p = min((A4[0] - 2 * mg) / sh.width, (A4[1] - 2 * mg) / sh.height)
            pw, ph = A4 if s_p > s_l * 1.05 else landscape(A4)
        else:
            pw, ph = A4 if portrait else landscape(A4)
        c.setPageSize((pw, ph))
        s = min((pw - 2 * mg) / sh.width, (ph - 2 * mg) / sh.height)
        ox = (pw - sh.width * s) / 2
        tx = lambda x, ox=ox, s=s: ox + x * s
        ty = lambda y, mg=mg, s=s: ph - (mg + y * s)
        for op in sh.ops:
            k = op[0]
            if k == "rect":
                _, x, y, w, h, fill, stroke, lw = op
                if not fill and not stroke:
                    continue
                if fill:
                    c.setFillColor(HexColor(fill))
                if stroke:
                    c.setStrokeColor(HexColor(stroke))
                    c.setLineWidth(max(lw * s, 0.1))
                c.rect(tx(x), ty(y + h), w * s, h * s, stroke=1 if stroke else 0, fill=1 if fill else 0)
            elif k == "line":
                _, x1, y1, x2, y2, col, lw = op
                c.setStrokeColor(HexColor(col))
                c.setLineWidth(max(lw * s, 0.1))
                c.setLineCap(1)
                c.line(tx(x1), ty(y1), tx(x2), ty(y2))
            elif k == "circle":
                _, cx, cy, r, fill, stroke, lw = op
                if fill:
                    c.setFillColor(HexColor(fill))
                if stroke:
                    c.setStrokeColor(HexColor(stroke))
                    c.setLineWidth(max(lw * s, 0.1))
                c.circle(tx(cx), ty(cy), r * s, stroke=1 if stroke else 0, fill=1 if fill else 0)
            elif k == "image":
                _, x, y, w, h, b64 = op
                try:
                    img = ImageReader(io.BytesIO(base64.b64decode(b64)))
                    c.drawImage(img, tx(x), ty(y + h), w * s, h * s, mask="auto", preserveAspectRatio=True)
                except Exception:
                    pass
            else:
                _, x, y, txt, size, bold, color, anchor, rot = op
                c.saveState()
                c.setFillColor(HexColor(color))
                c.setFont(fn_bold if bold else fn_reg, size * s)
                c.translate(tx(x), ty(y))
                if rot:
                    c.rotate(rot)
                if anchor == "middle":
                    c.drawCentredString(0, 0, txt)
                elif anchor == "end":
                    c.drawRightString(0, 0, txt)
                else:
                    c.drawString(0, 0, txt)
                c.restoreState()
        c.showPage()
    c.save()
    return buf.getvalue()


# =====================================================================
# 6. STYLE & KOMPONEN UI
# =====================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;500;600;700;800&display=swap');

:root{--navy:#073763;--blue:#2a6aa5;--ink:#073763;--muted:#5c7691;--line:#d5e1ee;--soft:#f2f7fc;--ok:#2e9e5b;--bad:#d64545;--warn:#f0a81c;
--shadow:0 6px 22px rgba(7,55,99,.10);--shadow-hi:0 14px 34px rgba(7,55,99,.20);
    --primary-color:#073763 !important;--primary:#073763 !important}
html{accent-color:#073763}
:root, html, body { color-scheme: light !important; }
html{scroll-behavior:smooth}
html, body, [class*="css"], .stApp, .stMarkdown, .stButton, .stTextInput,
.stSelectbox, .stDateInput, .stNumberInput, .stRadio, .stCheckbox, .stDataFrame,
[data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stAppViewContainer"] {
    font-family: 'Lexend', 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
    color-scheme: light !important;
}
.stApp{background:#ffffff !important;color:var(--ink)}
.stApp h1,.stApp h2,.stApp h3,.stApp h4{color:var(--navy) !important;font-weight:700;letter-spacing:.2px}
.stApp [data-testid="stWidgetLabel"] *{color:var(--ink) !important;font-weight:500}
.stApp [data-testid="stCaptionContainer"], .stApp [data-testid="stCaptionContainer"] *{color:var(--muted) !important}
[data-testid="stHeader"]{background:transparent !important}
.block-container{padding-top:1.4rem;max-width:100%;animation:fadeIn .5s ease both}
[data-testid="stDecoration"], [data-testid="stStatusWidget"], [data-testid="stAppDeployButton"], [data-testid="stToolbarActions"],
.stAppDeployButton, #MainMenu, footer, [data-testid="stMainMenu"]{display:none !important;visibility:hidden !important}
/* Toolbar (Share/GitHub/edit) disembunyikan dengan visibility, BUKAN display, supaya tombol pembuka sidebar yang bisa berada di dalamnya tetap bisa ditampilkan */
[data-testid="stToolbar"]{visibility:hidden !important}
[data-testid="stExpandSidebarButton"], [data-testid="stExpandSidebarButton"] *, [data-testid="stSidebarCollapsedControl"], [data-testid="stSidebarCollapsedControl"] *{visibility:visible !important}
/* Layar lebar: sidebar tidak bisa ditutup (tombol "<<" disembunyikan) agar tidak pernah "hilang" */
@media (min-width: 992px){[data-testid="stSidebarCollapseButton"]{display:none !important}}
::-webkit-scrollbar{width:10px;height:10px}::-webkit-scrollbar-thumb{background:#bfd2e8;border-radius:8px}::-webkit-scrollbar-thumb:hover{background:#8fb2d8}

/* ---------- animasi ---------- */
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
@keyframes floaty{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}
@keyframes shimmer{from{background-position:0% 0}to{background-position:300% 0}}
@keyframes blob{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(26px,18px) scale(1.12)}}
@keyframes dropIn{from{opacity:0;transform:translateY(-8px) scale(.98)}to{opacity:1;transform:none}}
@keyframes pulseRing{0%{box-shadow:0 0 0 0 rgba(214,69,69,.35)}70%{box-shadow:0 0 0 14px rgba(214,69,69,0)}100%{box-shadow:0 0 0 0 rgba(214,69,69,0)}}
@property --n{syntax:'<integer>';initial-value:0;inherits:false}
@keyframes cntup{from{--n:0}}
.cnt{counter-reset:num var(--n);animation:cntup 1.2s cubic-bezier(.2,.8,.2,1) both}
.cnt::after{content:counter(num) attr(data-s)}

/* ---------- header (latar biru Mandaya, logo di kapsul putih, teks putih) ---------- */
.logo-chip{display:inline-flex;align-items:center;background:#ffffff;border-radius:14px;padding:8px 16px;box-shadow:0 4px 14px rgba(0,0,0,.14)}
.logo-chip img{height:48px;width:auto;display:block}
.logo-chip.sm{padding:5px 10px;border-radius:10px}.logo-chip.sm img{height:30px}
.hero{position:relative;overflow:hidden;background:#19558c;border-radius:22px;padding:22px 34px 28px;margin-bottom:22px;text-align:center;box-shadow:var(--shadow);animation:fadeUp .6s ease both}
.hero .blob{position:absolute;border-radius:50%;filter:blur(46px);opacity:.18;pointer-events:none;background:#ffffff}
.hero .b1{width:240px;height:240px;left:-70px;top:-100px;animation:blob 9s ease-in-out infinite}
.hero .b2{width:220px;height:220px;right:-60px;bottom:-110px;animation:blob 11s ease-in-out infinite reverse}
.hero > *:not(.blob){position:relative;z-index:1}
.hero .top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px}
.wm{text-align:left;line-height:1}
.wm b{display:block;font-size:26px;font-weight:700;letter-spacing:2px;color:#ffffff !important}
.wm span{display:block;font-size:8.5px;font-weight:600;letter-spacing:7px;margin-top:5px;color:#ffffff !important}
.hero .hfy{font-size:21px;font-weight:300;color:#ffffff !important;letter-spacing:.3px}
.hero .ttl{font-size:clamp(22px,3.1vw,36px);font-weight:800;color:#ffffff !important;line-height:1.25;letter-spacing:.3px;margin:0 auto 8px;max-width:980px}
.hero .sub{font-size:clamp(12px,1.4vw,15px);color:rgba(255,255,255,.88) !important;letter-spacing:1.2px;margin-bottom:14px}
.hero .code{display:inline-block;font-size:12px;color:#ffffff !important;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.38);padding:5px 14px;border-radius:999px}
.hero-c{display:flex;align-items:center;gap:16px;background:#19558c;border-radius:16px;padding:12px 18px;margin-bottom:14px;box-shadow:var(--shadow);animation:fadeUp .5s ease both}
.hero-c .wm b{font-size:17px;letter-spacing:1.5px}.hero-c .wm span{font-size:6px;letter-spacing:5px;margin-top:3px}
.hero-c .ttl{font-size:14px;font-weight:800;color:#ffffff !important;line-height:1.3}
.hero-c .sub{font-size:11px;color:rgba(255,255,255,.85) !important;margin-top:2px}

/* ---------- kartu KPI (melayang) ---------- */
.kp{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin:8px 0 18px}
.kc{--c:#073763;background:#fff;border:1px solid var(--line);border-top:4px solid var(--c);border-radius:16px;padding:14px 18px;box-shadow:var(--shadow);
transition:transform .25s cubic-bezier(.2,.8,.2,1),box-shadow .25s;animation:fadeUp .55s ease both}
.kc:nth-child(2){animation-delay:.06s}.kc:nth-child(3){animation-delay:.12s}.kc:nth-child(4){animation-delay:.18s}.kc:nth-child(5){animation-delay:.24s}.kc:nth-child(6){animation-delay:.3s}
.kc:hover{transform:translateY(-6px);box-shadow:var(--shadow-hi)}
.kc .t{font-size:12px;color:#5c7691;font-weight:500}.kc .v{font-size:32px;font-weight:800;margin:4px 0 2px;color:#073763;line-height:1.1}.kc .s{font-size:11px;color:#5c7691}

/* ---------- kalender ---------- */
.cal{border-collapse:separate;border-spacing:2px;font-size:11px}
.cal td,.cal th{text-align:center;min-width:26px;height:26px;padding:0 2px;border-radius:6px;transition:transform .15s,box-shadow .15s}
.cal th{background:#dbe8f6;color:#073763;font-weight:600}
.cal td.l{text-align:left;font-weight:700;background:#f2f7fc;color:#073763;padding:0 10px;min-width:96px;white-space:nowrap;border-radius:8px}
.cal td:not(.l):hover{transform:scale(1.35);box-shadow:0 6px 14px rgba(7,55,99,.28);position:relative;z-index:2}
.s-ok{background:#2e9e5b;color:#fff}.s-high{background:#d64545;color:#fff;font-weight:700}.s-low{background:#f0a81c;color:#2b2000;font-weight:700}
.s-empty{background:#9aa5b1;color:#fff}.s-future{background:#eef1f8;color:#bbb}
.lg2{display:flex;gap:16px;font-size:12px;margin:10px 0;flex-wrap:wrap;color:#073763}.lg2 i{display:inline-block;width:14px;height:14px;border-radius:4px;margin-right:6px;vertical-align:-3px}

/* ---------- kotak, peringatan, floating ---------- */
.box{background:#fff;border:1px solid var(--line);border-radius:16px;padding:14px 18px;box-shadow:var(--shadow);margin-bottom:12px;font-size:13px;line-height:1.75;color:#073763;
transition:transform .25s cubic-bezier(.2,.8,.2,1),box-shadow .25s;animation:fadeUp .5s ease both}
.box:hover{transform:translateY(-3px);box-shadow:var(--shadow-hi)}
.alert{background:#fff4f4;border:1px solid #f5c6c6;border-left:6px solid #d64545;border-radius:14px;padding:12px 18px;margin:4px 0 14px;font-size:13px;line-height:1.75;color:#7a1f1f;
animation:fadeUp .5s ease both,pulseRing 2.4s ease-out infinite}
.hint{font-size:12px;margin:-8px 0 8px 2px}.hint.ok{color:#2e9e5b}.hint.bad{color:#d64545}
.empty{text-align:center;padding:46px 20px;background:#f2f7fc;border:1.5px dashed #bfd2e8;border-radius:22px;color:#073763;animation:fadeUp .5s ease both}
.empty .ic{font-size:56px;display:inline-block;animation:floaty 3.2s ease-in-out infinite}
.empty .h{font-size:20px;font-weight:700;margin:8px 0 4px}.empty .p{font-size:13px;color:#5c7691}
.fab{position:fixed;right:18px;bottom:76px;z-index:999;background:linear-gradient(135deg,#073763,#2a6aa5);color:#fff;border-radius:999px;padding:10px 18px;font-size:13px;font-weight:600;
box-shadow:0 10px 28px rgba(7,55,99,.38);animation:floaty 3.4s ease-in-out infinite;white-space:nowrap}
.fab small{opacity:.85;font-weight:400;margin-left:6px}

/* ---------- sidebar ---------- */
[data-testid="stSidebar"]{background:#ffffff !important;border-right:1px solid var(--line)}
[data-testid="stSidebarNavLink"], [data-testid="stSidebarNav"] a{border-radius:12px !important;margin:3px 8px !important;transition:background .18s,transform .18s,box-shadow .18s}
[data-testid="stSidebarNavLink"] *, [data-testid="stSidebarNav"] a *{color:#073763 !important;font-weight:500}
[data-testid="stSidebarNavLink"]:hover, [data-testid="stSidebarNav"] a:hover{background:#e6eff9 !important;transform:translateX(5px)}
[data-testid="stSidebarNavLink"][aria-current="page"], [data-testid="stSidebarNav"] a[aria-current="page"]{background:linear-gradient(90deg,#dbe8f6,#f2f7fc) !important;box-shadow:inset 4px 0 0 #073763}
[data-testid="stSidebarNavLink"][aria-current="page"] *{font-weight:700}

/* ---------- tombol ---------- */
.stButton>button, .stDownloadButton>button{border-radius:999px;border:1.5px solid var(--line);background:#fff;color:#073763;font-weight:600;padding:.45rem 1.2rem;
transition:transform .2s,box-shadow .2s,border-color .2s,color .2s;box-shadow:0 1px 2px rgba(7,55,99,.08)}
.stButton>button:hover, .stDownloadButton>button:hover{transform:translateY(-2px);box-shadow:0 10px 22px rgba(7,55,99,.18);border-color:#2a6aa5;color:#2a6aa5}
.stButton>button:active, .stDownloadButton>button:active{transform:translateY(0) scale(.98)}
.stButton>button[kind="primary"], [data-testid="stBaseButton-primary"]{background:linear-gradient(135deg,#073763,#2a6aa5) !important;border:none !important;color:#fff !important}
.stButton>button[kind="primary"] *, [data-testid="stBaseButton-primary"] *{color:#fff !important}
.stButton>button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover{box-shadow:0 12px 26px rgba(42,106,165,.45)}
.stButton>button:disabled, .stDownloadButton>button:disabled{opacity:.45;transform:none;box-shadow:none}

/* ---------- FIX DROPDOWN & TEXT INPUT (LIGHT THEME) ---------- */
div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"] > div {
    border-radius: 12px !important;
    background-color: #ffffff !important;
    border: 1.5px solid var(--line) !important;
    transition: border-color .18s, box-shadow .18s;
}
div[data-baseweb="input"] input,
div[data-baseweb="select"] input,
div[data-baseweb="select"] > div {
    color: #073763 !important;
    -webkit-text-fill-color: #073763 !important;
}
div[data-baseweb="select"] svg { fill: #073763 !important; }
div[data-baseweb="input"]:hover, div[data-baseweb="textarea"]:hover, div[data-baseweb="select"] > div:hover { border-color:#8fb2d8 !important; }
div[data-baseweb="input"]:focus-within, div[data-baseweb="textarea"]:focus-within, div[data-baseweb="select"] > div:focus-within { border-color:#2a6aa5 !important; box-shadow:0 0 0 4px rgba(42,106,165,.16); }
div[data-baseweb="popover"] ul[role="listbox"] { background-color: #ffffff !important; border: 1px solid var(--line); border-radius:14px; box-shadow:0 16px 38px rgba(7,55,99,.20); padding:6px; }
ul[role="listbox"] { list-style:none !important; margin:0; }
li[role="option"] { background-color: #ffffff !important; color: #073763 !important; border-radius:10px; margin:2px 0; transition:background .12s,padding-left .12s; }
li[role="option"]:hover, li[role="option"][aria-selected="true"] { background-color: #e6eff9 !important; color: #073763 !important; padding-left:18px; }
li[role="option"][aria-selected="true"]{font-weight:700;}
label[data-baseweb="checkbox"]:has(input:checked) > span:first-child{background-color:#073763 !important;border-color:#073763 !important}

/* ---------- TAB: BIRU TUA (menimpa merah bawaan Streamlit) ---------- */
[data-baseweb="tab-highlight"], [data-baseweb="tab-border"],
[role="tablist"] [data-baseweb="tab-highlight"], [role="tablist"] [data-baseweb="tab-border"],
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {
    display: none !important;
    background: transparent !important;
    height: 0 !important;
    opacity: 0 !important;
    border: none !important;
}
[role="tab"], button[data-baseweb="tab"] {
    background-color: #f2f7fc !important;
    border: 1.5px solid #bfd2e8 !important;
    border-radius: 12px !important;
    padding: 8px 20px !important;
    margin-right: 10px !important;
    transition: all 0.2s ease !important;
    outline: none !important;
    box-shadow: none !important;
}
[role="tab"] *, button[data-baseweb="tab"] * {
    color: #073763 !important;
    -webkit-text-fill-color: #073763 !important;
    font-weight: 700 !important;
    fill: #073763 !important;
}
[role="tab"]:hover, button[data-baseweb="tab"]:hover { background-color: #e6eff9 !important; }
[role="tab"][aria-selected="true"], button[data-baseweb="tab"][aria-selected="true"] {
    background: linear-gradient(90deg, #dbe8f6, #f2f7fc) !important;
    box-shadow: inset 4px 0 0 #073763 !important;
    border: 1.5px solid #bfd2e8 !important;
}
[role="tab"]:focus, [role="tab"]:active, [role="tab"]:focus-visible,
button[data-baseweb="tab"]:focus, button[data-baseweb="tab"]:active, button[data-baseweb="tab"]:focus-visible {
    outline: none !important;
    outline-color: transparent !important;
    box-shadow: inset 4px 0 0 #073763 !important;
    border-color: #bfd2e8 !important;
}

/* PERBAIKAN WARNA EXPANDER */
[data-testid="stExpander"]{border:1px solid var(--line) !important;border-radius:16px !important;background:#fff;box-shadow:var(--shadow);transition:box-shadow .25s}
[data-testid="stExpander"]:hover{box-shadow:var(--shadow-hi)}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] summary * { font-weight:600 !important; color: var(--navy) !important; }
[data-testid="stExpander"] summary svg { fill: var(--navy) !important; }

[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:14px;overflow:hidden;box-shadow:var(--shadow)}
[data-testid="stAlert"]{border-radius:14px}
@media (prefers-reduced-motion: reduce){*{animation:none !important;transition:none !important}}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def _logo_uri():
    b = _logo_b64()
    return ("data:image/png;base64," + b) if b else ""


def _brand_html(small=False):
    """Logo Mandaya di kapsul putih (logo berwarna biru tua, jadi perlu latar putih di atas header biru)."""
    uri = _logo_uri()
    if uri:
        return "<div class='logo-chip{}'><img src='{}' alt='{}'></div>".format(" sm" if small else "", uri, escape(HOSPITAL))
    return "<div class='wm'><b>MANDAYA</b><span>ROYAL HOSPITAL PURI</span></div>"   # cadangan bila file logo belum ada


def header(compact=False):
    """Header dashboard: latar biru Mandaya, logo di kapsul putih, teks putih."""
    wm = _brand_html(small=compact)
    if compact:   # versi ringkas untuk halaman QR (layar HP)
        st.markdown("<div class='hero-c'>{}<div><div class='ttl'>{}</div><div class='sub'>Suhu &middot; Kelembapan &middot; Daily QC &middot; {}</div></div></div>".format(
            wm, escape(APP_TITLE), escape(SITE.title())), unsafe_allow_html=True)
        return
    st.markdown("<div class='hero'><span class='blob b1'></span><span class='blob b2'></span><div class='top'>{}<div class='hfy'>Departemen Radiologi</div></div>"
                "<div class='ttl'>{}</div><div class='sub'>Monitoring Suhu &amp; Kelembapan &middot; Daily Quality Control</div>"
                "<div class='code'>{}</div></div>".format(wm, escape(APP_TITLE), escape(HOSPITAL)), unsafe_allow_html=True)


def month_year_bar(prefix, cols=(1.6, 1)):
    t = TODAY()
    c2, c3 = st.columns(list(cols))
    bln = c2.selectbox("Bulan", list(range(1, 13)), format_func=lambda x: BLN[x - 1], index=t.month - 1, key=prefix + "_bln")
    yrs = list(range(t.year - 2, t.year + 2))
    thn = c3.selectbox("Tahun", yrs, index=yrs.index(t.year), key=prefix + "_thn")
    return int(bln), int(thn)


def filter_bar(prefix, options, label):
    c1, c2, c3 = st.columns([2.2, 1.6, 1])
    pilih = c1.selectbox(label, options, key=prefix + "_pilih")
    t = TODAY()
    bln = c2.selectbox("Bulan", list(range(1, 13)), format_func=lambda x: BLN[x - 1], index=t.month - 1, key=prefix + "_bln")
    yrs = list(range(t.year - 2, t.year + 2))
    thn = c3.selectbox("Tahun", yrs, index=yrs.index(t.year), key=prefix + "_thn")
    return pilih, int(bln), int(thn)


def days_text(days):
    return ", ".join(str(d) for d in days) if days else "-"


# =====================================================================
# 6b. MODE AUTHOR (password di Streamlit Secrets: AUTHOR_PASSWORD)
# =====================================================================
def _author_secret():
    """Password author HANYA dari Streamlit Secrets (AUTHOR_PASSWORD) atau environment. Tidak ada password bawaan.

    Cara mengatur di Streamlit Cloud: app -> Settings -> Secrets, lalu tambahkan:
        AUTHOR_PASSWORD = "password_rahasia_anda"
    """
    return _secret("AUTHOR_PASSWORD")


def is_author():
    return bool(st.session_state.get("is_author"))


def author_box():
    """Kotak login author di sidebar. Tanpa AUTHOR_PASSWORD di Secrets, tidak ada yang bisa masuk."""
    with st.sidebar.expander("🔐 Author"):
        if is_author():
            st.success("Mode author aktif")
            if st.button("Keluar", key="auth_out"):
                st.session_state["is_author"] = False
                st.rerun()
            return
        pw = st.text_input("Password author", type="password", key="auth_pw")
        if st.button("Masuk", key="auth_btn"):
            secret = _author_secret()
            if secret and hmac.compare_digest(pw.encode("utf-8"), secret.encode("utf-8")):
                st.session_state["is_author"] = True
                st.rerun()
            else:
                time.sleep(1)
                st.error("Password salah." if secret else "Password author belum diatur di Secrets aplikasi.")


# =====================================================================
# 7. HALAMAN
# =====================================================================
def dash_th():
    st.caption("Rangkuman pencapaian pemeriksaan suhu & kelembapan (3x sehari: {})".format(SHIFT_TIMES))
    ruang, m, y = filter_bar("db", room_names(), "Ruangan")
    an = analyze(ruang, y, m)
    std = an["std"]
    if an["due"] == 0:
        st.info("Belum ada slot pengukuran yang jatuh tempo pada bulan ini.")
    st.markdown("""
    <div class='kp'>
      <div class='kc'><div class='t'>Persentase pengisian</div><div class='v'><span class='cnt' style='--n:{pf:.0f}' data-s='%'></span></div><div class='s'>{fd} data terisi dari {d} slot jatuh tempo</div></div>
      <div class='kc' style='border-top-color:#2e9e5b'><div class='t'>Suhu normal ({lo:g}&ndash;{hi:g} &deg;C)</div><div class='v' style='color:#2e9e5b'><span class='cnt' style='--n:{po:.0f}' data-s='%'></span></div><div class='s'>{o} dari {f} data terisi</div></div>
      <div class='kc' style='border-top-color:#2e9e5b'><div class='t'>Kelembapan dalam batas ({hl:g}&ndash;{hh:g} %)</div><div class='v' style='color:#2e9e5b'><span class='cnt' style='--n:{ph:.0f}' data-s='%'></span></div><div class='s'>{ho} dari {f} data terisi</div></div>
      <div class='kc' style='border-top-color:#d64545'><div class='t'>Suhu DI ATAS batas (&gt; {hi:g} &deg;C)</div><div class='v' style='color:#d64545'><span class='cnt' style='--n:{hi_n}' data-s=''></span></div><div class='s'>slot pengukuran tidak normal</div></div>
      <div class='kc' style='border-top-color:#f0a81c'><div class='t'>Suhu di bawah batas (&lt; {lo:g} &deg;C)</div><div class='v' style='color:#d98e00'><span class='cnt' style='--n:{lo_n}' data-s=''></span></div><div class='s'>slot pengukuran</div></div>
      <div class='kc' style='border-top-color:#9aa5b1'><div class='t'>Jumlah tanggal yang belum diisi</div><div class='v' style='color:#667'><span class='cnt' style='--n:{e}' data-s=''></span></div><div class='s'>{fe} tanggal yang belum diisi</div></div>
    </div>""".format(pf=an["pct_fill"], f=an["filled"], fd=an["filled_due"], d=an["due"], lo=std["t_lo"], hi=std["t_hi"], po=an["pct_ok"], o=an["ok"],
                     hl=std["h_lo"], hh=std["h_hi"], ph=an["pct_hum"], ho=an["hum_ok"],
                     hi_n=len(an["high"]), lo_n=len(an["low"]), e=an["n_empty"], fe=len(an["full_empty"])), unsafe_allow_html=True)

    head = "<tr><th></th>" + "".join("<th>{}</th>".format(d) for d in range(1, an["n"] + 1)) + "</tr>"
    body = ""
    for s in "PSM":
        body += "<tr><td class='l'>Shift {}</td>".format(SHIFT_NAME[s])
        for d in range(1, an["n"] + 1):
            c = an["slots"][(d, s)]
            stt = c["state"]
            tip = {"ok": "{:g}°C normal", "high": "{:g}°C DI ATAS batas", "low": "{:g}°C di bawah batas"}.get(stt, "")
            tip = "" if stt == "future" else "Tgl {} {}: ".format(d, SHIFT_NAME[s]) + (tip.format(c["t"]) if tip else "BELUM DIISI")
            txt = {"ok": "&#10003;", "high": "&#9650;", "low": "&#9660;", "empty": "&ndash;", "future": ""}[stt]
            body += "<td class='s-{}' title='{}'>{}</td>".format(stt, tip, txt)
        body += "</tr>"
    st.markdown("<div style='overflow-x:auto'><table class='cal'>{}{}</table></div>".format(head, body), unsafe_allow_html=True)
    st.markdown("<div class='lg2'><span><i class='s-ok'></i>Normal ({lo:g}&ndash;{hi:g} &deg;C)</span><span><i class='s-high'></i>&#9650; Di atas batas (tidak normal)</span><span><i class='s-low'></i>&#9660; Di bawah batas</span>"
                "<span><i class='s-empty'></i>Belum diisi</span></div>".format(lo=std["t_lo"], hi=std["t_hi"]), unsafe_allow_html=True)

    data = month_data(ruang, y, m)
    if data:
        per_day = {}
        for (d, _s), e in data.items():
            per_day.setdefault(d, []).append(e)
        dfc = pd.DataFrame({"Tanggal": sorted(per_day),
                            "Suhu rata-rata (°C)": [sum(x["suhu"] for x in per_day[d]) / len(per_day[d]) for d in sorted(per_day)],
                            "Kelembapan rata-rata (%)": [sum(x["hum"] for x in per_day[d]) / len(per_day[d]) for d in sorted(per_day)]}).set_index("Tanggal")
        g1, g2 = st.columns(2)
        g1.markdown("**Tren suhu harian (rata-rata)**")
        g1.line_chart(dfc[["Suhu rata-rata (°C)"]])
        g2.markdown("**Tren kelembapan harian (rata-rata)**")
        g2.line_chart(dfc[["Kelembapan rata-rata (%)"]])

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Suhu di atas batas**")
        st.dataframe(pd.DataFrame([dict(Tanggal=d, Shift=SHIFT_NAME[s], **{"Suhu (°C)": v, "Batas atas": std["t_hi"], "Selisih": round(v - std["t_hi"], 1)}) for d, s, v in an["high"]]
                                  or [dict(Tanggal="-", Shift="-", **{"Suhu (°C)": "-", "Batas atas": "-", "Selisih": "-"})]), hide_index=True)
        if an["low"]:
            st.markdown("**Suhu di bawah batas**")
            st.dataframe(pd.DataFrame([dict(Tanggal=d, Shift=SHIFT_NAME[s], **{"Suhu (°C)": v, "Batas bawah": std["t_lo"], "Selisih": round(v - std["t_lo"], 1)}) for d, s, v in an["low"]]), hide_index=True)
    with c2:
        st.markdown("**Riwayat pengisian**")
        st.dataframe(pd.DataFrame([dict(Tanggal=d, **{"Shift belum diisi": ", ".join(SHIFT_NAME[x] for x in v), "Status": "Belum diisi" if len(v) == 3 else "Sebagian"}) for d, v in sorted(an["empty_by_day"].items())]
                                  or [dict(Tanggal="-", **{"Shift belum diisi": "Semua terisi", "Status": "-"})]), hide_index=True)
    with c3:
        st.markdown("**Kelembapan di luar batas ({:g}–{:g} %)**".format(std["h_lo"], std["h_hi"]))
        st.dataframe(pd.DataFrame([dict(Tanggal=d, Shift=SHIFT_NAME[s], **{"Kelembapan (%)": v, "Keterangan": "Di atas" if k == "high" else "Di bawah"}) for d, s, v, k in an["hum_bad"]]
                                  or [dict(Tanggal="-", Shift="-", **{"Kelembapan (%)": "-", "Keterangan": "-"})]), hide_index=True)

    st.markdown("**Rekap semua ruangan - {} {}**".format(BLN[m - 1], y))
    rows = []
    for rn in room_names():
        a = analyze(rn, y, m)
        rows.append({"Ruangan": rn, "% Pengisian": round(a["pct_fill"]), "% Suhu normal": round(a["pct_ok"]), "% Kelembapan dalam batas": round(a["pct_hum"]),
                     "Suhu di atas batas": len(a["high"]), "Suhu di bawah batas": len(a["low"]), "Tanggal yang kosong": a["n_empty"],
                     "Tanggal yang belum diisi": days_text(a["full_empty"])})
    st.dataframe(pd.DataFrame(rows), hide_index=True)


def dash_qc():
    st.caption("Rangkuman Daily Quality Control peralatan radiologi")
    m, y = month_year_bar("qd")
    devs = devices()
    ans = {d["nama"]: qc_analyze(d, y, m) for d in devs}
    tot_counted = sum(a["counted"] for a in ans.values())
    tot_full = sum(a["full"] for a in ans.values())
    tot_idone = sum(a["item_done"] for a in ans.values())
    tot_itotal = sum(a["item_total"] for a in ans.values())
    tot_empty = sum(a["n_empty"] for a in ans.values())
    tot_part = sum(a["n_part"] for a in ans.values())
    if tot_counted == 0:
        st.info("Belum ada hari yang jatuh tempo / terisi pada bulan ini.")
    st.markdown("""
    <div class='kp'>
      <div class='kc'><div class='t'>Hari checklist lengkap</div><div class='v'><span class='cnt' style='--n:{pf:.0f}' data-s='%'></span></div><div class='s'>{full} dari {cnt} hari-alat jatuh tempo</div></div>
      <div class='kc' style='border-top-color:#2e9e5b'><div class='t'>Kegiatan terselesaikan</div><div class='v' style='color:#2e9e5b'><span class='cnt' style='--n:{pi:.0f}' data-s='%'></span></div><div class='s'>{idn} dari {itot} kegiatan harian</div></div>
      <div class='kc' style='border-top-color:#f0a81c'><div class='t'>Hanya sebagian</div><div class='v' style='color:#d98e00'><span class='cnt' style='--n:{part}' data-s=''></span></div><div class='s'>hari-alat belum lengkap</div></div>
      <div class='kc' style='border-top-color:#9aa5b1'><div class='t'>Hari belum diisi</div><div class='v' style='color:#667'><span class='cnt' style='--n:{emp}' data-s=''></span></div><div class='s'>hari-alat belum diisi</div></div>
    </div>""".format(pf=(tot_full / tot_counted * 100 if tot_counted else 0), full=tot_full, cnt=tot_counted,
                     pi=(tot_idone / tot_itotal * 100 if tot_itotal else 0), idn=tot_idone, itot=tot_itotal, part=tot_part, emp=tot_empty),
                unsafe_allow_html=True)

    n = calendar.monthrange(y, m)[1]
    head = "<tr><th></th>" + "".join("<th>{}</th>".format(d) for d in range(1, n + 1)) + "</tr>"
    body = ""
    cls = {"full": "ok", "part": "low", "empty": "empty", "future": "future"}
    txt = {"full": "&#10003;", "part": "&#189;", "empty": "&ndash;", "future": ""}
    for dev in devs:
        a = ans[dev["nama"]]
        body += "<tr><td class='l'>{}</td>".format(escape(dev["nama"]))
        for d in range(1, n + 1):
            c = a["days"][d]
            tip = "{} tgl {}: {}/{} kegiatan".format(escape(dev["nama"]), d, c["done"], c["total"]) if c["state"] != "future" else ""
            body += "<td class='s-{}' title='{}'>{}</td>".format(cls[c["state"]], tip, txt[c["state"]])
        body += "</tr>"
    st.markdown("<div style='overflow-x:auto'><table class='cal'>{}{}</table></div>".format(head, body), unsafe_allow_html=True)
    st.markdown("<div class='lg2'><span><i class='s-ok'></i>&#10003; Lengkap</span><span><i class='s-low'></i>&#189; Sebagian</span>"
                "<span><i class='s-empty'></i>Belum diisi</span></div>", unsafe_allow_html=True)

    st.markdown("**Rekap semua peralatan - {} {}**".format(BLN[m - 1], y))
    st.dataframe(pd.DataFrame([{"Peralatan": d["nama"], "Hari lengkap": ans[d["nama"]]["full"], "Hari jatuh tempo": ans[d["nama"]]["counted"],
                                "% Hari lengkap": round(ans[d["nama"]]["pct_full"]), "% Kegiatan selesai": round(ans[d["nama"]]["pct_item"]),
                                "Hanya sebagian": ans[d["nama"]]["n_part"], "Hari belum diisi": ans[d["nama"]]["n_empty"]} for d in devs]), hide_index=True)


def _status_text(e, std):
    ts = classify(e["suhu"], std["t_lo"], std["t_hi"])
    hs = classify(e["hum"], std["h_lo"], std["h_hi"])
    parts = []
    if ts != "ok":
        parts.append("🔴 Suhu di atas batas" if ts == "high" else "🟠 Suhu di bawah batas")
    if hs != "ok":
        parts.append("🔴 Kelembapan di atas batas" if hs == "high" else "🟠 Kelembapan di bawah batas")
    return " · ".join(parts) if parts else "✅ Normal"


def history_tab():
    """Riwayat angka hasil pengisian (terbaru di atas). Data lama tetap bisa dibuka dengan memilih bulan/tahun."""
    st.caption("Riwayat angka hasil pengisian. Pilih ruangan/peralatan, bulan, dan tahun; data bulan-bulan sebelumnya tetap tersimpan.")
    jenis = st.radio("Jenis data", ["Suhu & Kelembapan", "QC Peralatan"], horizontal=True, key="rw_jenis")
    if jenis == "Suhu & Kelembapan":
        pilih, m, y = filter_bar("rw", ["Semua ruangan"] + room_names(), "Ruangan")
        targets = room_names() if pilih == "Semua ruangan" else [pilih]
        rows, ts_all, hs_all = [], [], []
        for rn in targets:
            std = room_std(rn)
            for (d, s), e in month_data(rn, y, m).items():
                ts_all.append(e["suhu"])
                hs_all.append(e["hum"])
                rows.append((d, "PSM".index(s), {"Tanggal": dt.date(y, m, d).strftime("%d/%m/%Y"), "Shift": SHIFT_NAME[s], "Ruangan": rn,
                                                  "Suhu (°C)": e["suhu"], "Kelembapan (%)": e["hum"], "Status": _status_text(e, std),
                                                  "Staff": staff_label(e["petugas"]), "Temuan": e["ket"]}))
        if not rows:
            st.info("Belum ada data pada {} {} untuk {}.".format(BLN[m - 1], y, pilih))
            return
        rows.sort(key=lambda r: (r[0], r[1]), reverse=True)   # terbaru di atas
        cards = [("Jumlah data", str(len(rows))), ("Suhu rata-rata", "{:.1f} °C".format(sum(ts_all) / len(ts_all))),
                 ("Suhu terendah – tertinggi", "{:g} – {:g} °C".format(min(ts_all), max(ts_all))),
                 ("Kelembapan rata-rata", "{:.0f} %".format(sum(hs_all) / len(hs_all))),
                 ("Kelembapan terendah – tertinggi", "{:g} – {:g} %".format(min(hs_all), max(hs_all)))]
        st.markdown("<div class='kp'>" + "".join("<div class='kc'><div class='t'>{}</div><div class='v' style='font-size:22px'>{}</div></div>".format(t, v)
                                                 for t, v in cards) + "</div>", unsafe_allow_html=True)
        df = pd.DataFrame([r[2] for r in rows])
        if pilih != "Semua ruangan":
            df = df.drop(columns=["Ruangan"])
        st.dataframe(df, hide_index=True)
        st.download_button("⬇️ Unduh tabel ini (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                           "Riwayat_Suhu_{}_{}_{:02d}.csv".format(pilih.replace(" ", "_"), y, m), "text/csv", key="rw_csv")
    else:
        alat, m, y = filter_bar("rq", [d["nama"] for d in devices()], "Peralatan")
        dev = dev_by_name(alat)
        qd = qc_month(dev, y, m)
        daily = [i for i in dev["items"] if is_daily(i)]
        rows = []
        for d in sorted(qd, reverse=True):   # terbaru di atas
            e = qd[d]
            n_done = len([i for i in daily if i in e["done"]])
            rows.append({"Tanggal": dt.date(y, m, d).strftime("%d/%m/%Y"), "Kegiatan selesai": "{}/{}".format(n_done, len(daily)),
                         "Status": "✅ Lengkap" if n_done == len(daily) else "½ Sebagian", "Staff": staff_label(e["petugas"]) if e["petugas"] else "-",
                         "Catatan": e["catatan"]})
        if not rows:
            st.info("Belum ada data QC pada {} {} untuk {}.".format(BLN[m - 1], y, alat))
            return
        df = pd.DataFrame(rows)
        st.dataframe(df, hide_index=True)
        st.download_button("⬇️ Unduh tabel ini (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                           "Riwayat_QC_{}_{}_{:02d}.csv".format(alat.replace(" ", "_"), y, m), "text/csv", key="rq_csv")


def page_dashboard():
    header()
    if is_author() and not db_on():
        st.warning("Data masih disimpan di CSV lokal dan akan HILANG saat aplikasi restart. Hubungkan database Supabase (lihat README).")
    bad = abnormal_today()
    if bad:
        st.markdown("<div class='alert'><b>&#9888; Hubungi teknisi / maintenance</b> - pengukuran hari ini di luar batas normal:<br>{}</div>".format("<br>".join(bad)), unsafe_allow_html=True)
    t1, t2, t3 = st.tabs(["🌡️ Suhu & Kelembapan", "🛠️ QC Peralatan", "📋 Riwayat Data"])
    with t1:
        dash_th()
    with t2:
        dash_qc()
    with t3:
        history_tab()


HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


def _show_msg(key):
    msg = st.session_state.pop(key, None)
    if msg:
        (st.success if msg[0] == "success" else st.error)(msg[1])
        if msg[0] == "success":
            st.toast("Data berhasil disimpan", icon="✅")


def _save_th(p, ruang, tgl, shift_key, petugas):
    ss = st.session_state
    std = room_std(ruang)
    suhu, hum = ss.get(p + "_suhu"), ss.get(p + "_hum")
    temuan = (ss.get(p + "_temuan") or "").strip()
    tindakan = ss.get(p + "_tindakan") or ""
    if suhu is None or hum is None:
        ss[p + "_msg"] = ("error", "Suhu dan kelembapan wajib diisi.")
        return
    abn = classify(suhu, std["t_lo"], std["t_hi"]) != "ok" or classify(hum, std["h_lo"], std["h_hi"]) != "ok"
    if abn and not temuan:
        ss[p + "_msg"] = ("error", "Kolom Temuan wajib diisi karena suhu/kelembapan di luar batas standar.")
        return
    try:
        save_entry(ruang=ruang, tanggal=tgl.isoformat(), shift=shift_key, suhu=suhu, kelembapan=hum, petugas=petugas,
                   status="Abnormal" if abn else "Normal", temuan=temuan, tindakan=tindakan if abn else "")
    except StorageError as e:
        ss[p + "_msg"] = ("error", "DATA BELUM TERSIMPAN - {}. Coba lagi atau hubungi admin.".format(e))
        return
    ss[p + "_msg"] = ("success", "Tersimpan: {} - {} - {} ({} °C / {} %).{}".format(
        ruang, tgl.strftime("%d/%m/%Y"), SHIFT_NAME[shift_key], fnum(suhu), fnum(hum), " Segera hubungi teknisi." if abn else ""))
    for k in ("_suhu", "_hum", "_temuan", "_timpa"):   # kosongkan isian agar tidak tersimpan dobel di jam berikutnya
        ss.pop(p + k, None)


def th_entry_form(ruang, tgl, shift_key, petugas, p):
    """Isian suhu & kelembapan (dipakai halaman QR dan halaman manual)."""
    std = room_std(ruang)
    _show_msg(p + "_msg")
    c3, c4 = st.columns(2)
    suhu = c3.number_input("Suhu (°C) [normal {:g}–{:g}] *".format(std["t_lo"], std["t_hi"]), value=None, min_value=0.0, max_value=60.0, step=0.1,
                           format="%.1f", placeholder="hasil baca termometer", key=p + "_suhu")
    hum = c4.number_input("Kelembapan (%) [standar {:g}–{:g}] *".format(std["h_lo"], std["h_hi"]), value=None, min_value=0.0, max_value=100.0, step=1.0,
                          format="%.0f", placeholder="hasil baca higrometer", key=p + "_hum")
    ts = classify(suhu, std["t_lo"], std["t_hi"]) if suhu is not None else None
    hs = classify(hum, std["h_lo"], std["h_hi"]) if hum is not None else None
    if ts:
        c3.markdown("<div class='hint {}'>{}</div>".format("ok" if ts == "ok" else "bad", "✓ Suhu normal" if ts == "ok" else "⚠ Tidak normal: {} batas ({:g}–{:g} °C)".format("DI ATAS" if ts == "high" else "DI BAWAH", std["t_lo"], std["t_hi"])), unsafe_allow_html=True)
    if hs:
        c4.markdown("<div class='hint {}'>{}</div>".format("ok" if hs == "ok" else "bad", "✓ Dalam batas standar" if hs == "ok" else "⚠ {} batas standar ({:g}–{:g} %)".format("DI ATAS" if hs == "high" else "DI BAWAH", std["h_lo"], std["h_hi"])), unsafe_allow_html=True)
    abn = (ts is not None and ts != "ok") or (hs is not None and hs != "ok")
    if abn:
        st.error("Di luar batas normal - segera hubungi teknisi / maintenance.")
        st.text_area("Temuan ketidaksesuaian * (wajib karena di luar batas standar)", key=p + "_temuan")
        st.selectbox("Tindakan korektif *", ["Lapor Teknisi AC / Maintenance", "Pengaturan AC / Dehumidifier", "Lainnya"], key=p + "_tindakan")
    ada = load_entries(ym=tgl)
    same = ada[(ada["ruang"] == ruang) & (ada["tanggal"] == tgl.isoformat()) & (ada["shift"] == shift_key)]
    dup = len(same) > 0
    timpa = True
    if dup:
        st.warning("Pengisian Suhu Ruangan {} pada tanggal {} Sudah diisi oleh {}".format(
            ruang, tgl.strftime("%d/%m/%Y"), staff_label(same.iloc[-1]["petugas"])))
        timpa = st.checkbox("Timpa data yang sudah ada", key=p + "_timpa")
    st.button("Simpan suhu & kelembapan", type="primary", key=p + "_save", disabled=(dup and not timpa),
              on_click=_save_th, args=(p, ruang, tgl, shift_key, petugas))


def _set_all(keys, val):
    for k in keys:
        st.session_state[k] = val


def _save_qc_cb(pk, alat, tgl, petugas, pairs):
    ss = st.session_state
    hasil = {it: True for k, it in pairs if ss.get(k)}
    catatan = (ss.get(pk + "_cat") or "").strip()
    if not hasil and not catatan:
        ss[pk + "_msg"] = ("error", "Belum ada kegiatan yang dicentang.")
        return
    try:
        save_qc(alat, tgl.isoformat(), petugas, hasil, catatan)
    except StorageError as e:
        ss[pk + "_msg"] = ("error", "DATA BELUM TERSIMPAN - {}. Coba lagi atau hubungi admin.".format(e))
        return
    ss[pk + "_msg"] = ("success", "Checklist tersimpan: {} - {} ({} kegiatan dicentang).".format(alat, tgl.strftime("%d/%m/%Y"), len(hasil)))


def qc_checklist_ui(dev, tgl, petugas, p):
    """Daftar centang QC 1 alat 1 hari (dipakai halaman QR dan halaman manual)."""
    alat = dev["nama"]
    pk = "{}_{}_{}".format(p, alat, tgl.isoformat())
    _show_msg(pk + "_msg")
    existing = qc_get(alat, tgl)
    preset = existing["done"] if existing else set()
    if existing:
        st.info("Checklist {} tanggal {} sudah ada (staff {}). Menyimpan akan memperbarui data tersebut.".format(alat, tgl.strftime("%d/%m/%Y"), staff_label(existing["petugas"])))
    kidx = {it: i for i, it in enumerate(dev["items"])}
    pairs = [("{}_i{}".format(pk, i), it) for i, it in enumerate(dev["items"])]
    keys = [k for k, _ in pairs]
    b1, b2, _sp = st.columns([1, 1, 2])
    b1.button("Centang semua", on_click=_set_all, args=(keys, True), key=pk + "_all")
    b2.button("Kosongkan", on_click=_set_all, args=(keys, False), key=pk + "_none")
    for sec in dev_sections(dev):          # tampil per seksi, sama seperti urutan di form
        if sec["title"]:
            st.markdown("**{}**".format(escape(sec["title"].strip())))
        for it in sec["items"]:
            k = "{}_i{}".format(pk, kidx[it["key"]])
            if k not in st.session_state:
                st.session_state[k] = it["key"] in preset
            label = "**{}**".format(it["kegiatan"]) + (" — {}".format(it["param"]) if it["param"] else "")
            st.checkbox(label, key=k)
    for ln in dev.get("ket") or []:
        st.caption(ln)
    st.text_area("Catatan / temuan (opsional)", value=existing["catatan"] if existing else "", key=pk + "_cat")
    st.button("Simpan checklist QC", type="primary", key=pk + "_save", on_click=_save_qc_cb, args=(pk, alat, tgl, petugas, pairs))


def page_ruang():
    """Halaman tujuan QR code: /ruang?ruang=NAMA -> isi suhu & kelembapan + QC untuk ruangan itu."""
    header(compact=True)
    names = room_names()
    qp = st.query_params.get("ruang", "")
    if isinstance(qp, list):
        qp = qp[0] if qp else ""
    ss = st.session_state
    if qp in names:
        ss["locked_room"] = qp    # scan QR -> ruangan terkunci selama sesi (tetap terkunci walau pindah menu lalu kembali)
    locked = ss.get("locked_room")
    if is_author():               # hanya author yang boleh memilih ruangan secara manual
        ruang = st.selectbox("Ruangan (mode author)", names, index=names.index(locked) if locked in names else 0, key="qr_pilih")
    elif qp and qp not in names:
        st.error("QR code tidak dikenali (ruangan '{}'). Scan ulang QR code di ruangan.".format(qp))
        return
    elif locked in names:
        ruang = locked
    else:
        st.markdown("<div class='empty'><div class='ic'>🔳</div><div class='h'>Scan QR code di ruangan</div>"
                    "<div class='p'>Scan QR code pada setiap ruangan yang tersedia</div></div>", unsafe_allow_html=True)
        return
    n = now()
    tgl = n.date()
    data = month_data(ruang, tgl.year, tgl.month)
    badges = " &nbsp; ".join("{} {} ({:02d}.00)".format("✅" if (tgl.day, s) in data else "⬜", SHIFT_NAME[s], SHIFT_HOUR[s]) for s in "PSM")
    st.markdown("<div class='box'><span style='font-size:12px;color:#667'>Ruangan</span><br><span style='font-size:26px;font-weight:700;color:#073763'>{}</span><br>"
                "<span style='font-size:12px;color:#556'>{}, {} {} {} &middot; {} WIB</span><br><span style='font-size:12px'>Sudah diisi hari ini: {}</span>{}</div>".format(
                    escape(ruang), HARI[tgl.weekday()], tgl.day, BLN[tgl.month - 1], tgl.year, n.strftime("%H:%M"), badges,
                    "" if is_author() else "<br><span style='font-size:11px;color:#667'>&#128274; Terkunci ke ruangan ini. Untuk ruangan lain, scan QR code ruangan tersebut.</span>"), unsafe_allow_html=True)
    petugas = st.selectbox("Nama staff (inisial) *", load_cfg()["staff"], key="qr_petugas")
    st.markdown("### 1. Suhu & Kelembapan")
    sk = shift_now()
    shift_key = st.radio("Jam pemantauan", list(SHIFT_LABEL.keys()), index=list("PSM").index(sk), format_func=lambda k: SHIFT_LABEL[k],
                         horizontal=True, key="qr_shift_" + sk)
    st.markdown("<div class='fab'>{} {} <small>{}</small></div>".format("🔓" if is_author() else "🔒", escape(ruang), SHIFT_NAME[shift_key]), unsafe_allow_html=True)
    th_entry_form(ruang, tgl, shift_key, petugas, "qr")
    st.markdown("### 2. Daily Quality Control Peralatan")
    devs = room_devices(ruang)
    if not devs:
        st.caption("Tidak ada daftar periksa peralatan untuk ruangan ini.")
    for dev in devs:
        with st.expander("QC {}".format(dev["nama"]), expanded=(len(devs) == 1)):
            qc_checklist_ui(dev, tgl, petugas, "qr")


def YESTERDAY():
    return TODAY() - dt.timedelta(days=1)


def backdate_th():
    st.caption("Backdate suhu & kelembapan - hanya tanggal sebelum hari ini. Pengisian real-time lewat QR code di ruangan.")
    cfg = load_cfg()
    left, right = st.columns([2.2, 1], gap="large")
    with left:
        ruang = st.selectbox("Ruangan *", room_names(), key="in_ruang")
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal backdate *", value=YESTERDAY(), max_value=YESTERDAY(), key="in_tgl")
        shift_key = c2.radio("Jam pemantauan *", list(SHIFT_LABEL.keys()), format_func=lambda k: SHIFT_LABEL[k], horizontal=True, key="in_shift")
        petugas = st.selectbox("Nama staff (inisial) *", cfg["staff"], key="in_petugas")
        th_entry_form(ruang, tgl, shift_key, petugas, "in")
    with right:
        st.markdown("<div class='box'><b>Panduan pengisian</b><br>Menu ini <b>hanya untuk backdate</b> (tanggal sebelum hari ini). Pengisian real-time dilakukan lewat <b>scan QR code di ruangan</b>.<br><br>"
                    "1. Pilih ruangan, tanggal, dan jam ({}).<br>2. Isi suhu & kelembapan sesuai hasil baca alat.<br>3. Pilih inisial staff.<br>"
                    "4. Bila di luar batas standar, temuan & tindakan korektif wajib diisi dan segera hubungi teknisi.</div>".format(SHIFT_TIMES), unsafe_allow_html=True)
    st.markdown("**Data terakhir - {}**".format(ruang))
    df = load_entries(recent=300)
    df = df[df["ruang"] == ruang]
    if len(df):
        df = df.sort_values("id", key=lambda s: s.astype(int), ascending=False).head(15)
        st.dataframe(df[["id", "tanggal", "shift", "suhu", "kelembapan", "petugas", "status", "temuan"]], hide_index=True)
        cc1, cc2 = st.columns([1, 3])
        did = cc1.selectbox("Hapus data ID", list(df["id"]), key="in_del")
        if cc2.button("Hapus data terpilih", key="in_del_btn"):
            try:
                delete_entry(did)
                st.success("Data ID {} dihapus. Muat ulang halaman untuk memperbarui tabel.".format(did))
            except StorageError as e:
                st.error("Gagal menghapus: {}".format(e))
    else:
        st.caption("Belum ada data yang diinput untuk ruangan ini.")


def backdate_qc():
    st.caption("Backdate daftar periksa harian - hanya tanggal sebelum hari ini. Pengisian real-time lewat QR code di ruangan.")
    cfg = load_cfg()
    left, right = st.columns([2.2, 1], gap="large")
    with left:
        alat = st.selectbox("Peralatan *", [d["nama"] for d in devices()], key="qi_alat")
        dev = dev_by_name(alat)
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal backdate *", value=YESTERDAY(), max_value=YESTERDAY(), key="qi_tgl")
        petugas = c2.selectbox("Nama staff (inisial) *", cfg["staff"], key="qi_petugas")
        st.caption(" · ".join("{}: {}".format(a, b) for a, b in dev["info"][:3]))
        qc_checklist_ui(dev, tgl, petugas, "qi")
    with right:
        st.markdown("<div class='box'><b>Panduan pengisian</b><br>Menu ini <b>hanya untuk backdate</b> (tanggal sebelum hari ini). Pengisian real-time lewat <b>scan QR code di ruangan</b>.<br><br>1. Pilih peralatan dan tanggal.<br>"
                    "2. Centang kegiatan yang sudah dikerjakan (atau <i>Centang semua</i>).<br>"
                    "3. Catatan tampil di bawah formulir.</div>", unsafe_allow_html=True)
    st.markdown("**Data terakhir - {}**".format(alat))
    df = load_qc(recent=300)
    df = df[df["alat"] == alat]
    if len(df):
        df = df.sort_values("id", key=lambda s: s.astype(int), ascending=False).head(15).copy()

        def _n(h):
            try:
                return len([1 for v in json.loads(h).values() if v])
            except Exception:
                return 0
        df["jumlah dicentang"] = df["hasil"].map(_n)
        st.dataframe(df[["id", "tanggal", "petugas", "jumlah dicentang", "catatan"]], hide_index=True)
        cc1, cc2 = st.columns([1, 3])
        did = cc1.selectbox("Hapus data ID", list(df["id"]), key="qi_del")
        if cc2.button("Hapus data terpilih", key="qi_del_btn"):
            try:
                delete_qc(did)
                st.success("Data ID {} dihapus. Muat ulang halaman untuk memperbarui tabel.".format(did))
            except StorageError as e:
                st.error("Gagal menghapus: {}".format(e))
    else:
        st.caption("Belum ada checklist yang diinput untuk peralatan ini.")


def page_backdate():
    header()
    if not is_author():
        st.error("Halaman ini hanya untuk author.")
        return
    t1, t2 = st.tabs(["🌡️ Suhu & Kelembapan", "🛠️ QC Peralatan"])
    with t1:
        backdate_th()
    with t2:
        backdate_qc()


def _preview(sh, base_px=1400, scale=0.85):
    """Preview formulir dengan ukuran tetap (85%)."""
    wpx = int(base_px * scale)
    components.html("<div style='overflow:auto;background:#fff;border:1px solid #d5dbe3'>{}</div>".format(sheet_to_svg(sh, wpx)),
                    height=int(wpx * sh.height / sh.width) + 30, scrolling=True)


def page_download():
    header()
    names = ["🌡️ Form Suhu & Kelembapan", "🛠️ Form Daily QC"]
    if is_author():   # generator QR hanya untuk author
        names.append("🔳 QR Code Ruangan")
    tabs = st.tabs(names)
    t1, t2 = tabs[0], tabs[1]
    with t1:
        ruang, m, y = filter_bar("dl", room_names(), "Ruangan")
        sh = build_th_form(ruang, y, m, month_data(ruang, y, m))
        _preview(sh)
        fname = "Form_Suhu_Kelembapan_{}_{}_{:02d}".format(ruang.replace(" ", "_"), y, m)
        c1, c2, c3 = st.columns(3)
        try:
            c1.download_button("⬇️ PDF (ruangan ini)", sheets_to_pdf([sh], "Pemeriksaan Suhu & Kelembapan - " + ruang), fname + ".pdf", "application/pdf", key="dl_pdf")
            allsh = [build_th_form(rn, y, m, month_data(rn, y, m)) for rn in room_names()]
            c2.download_button("⬇️ PDF (semua ruangan)", sheets_to_pdf(allsh, "Pemeriksaan Suhu & Kelembapan - Semua Ruangan"),
                               "Form_Suhu_Kelembapan_SEMUA_{}_{:02d}.pdf".format(y, m), "application/pdf", key="dl_pdf_all")
        except ImportError:
            st.error("Library reportlab belum terpasang. Jalankan:  pip install reportlab")
        rows = [{"Ruang": ruang, "Tanggal": dt.date(y, m, d).isoformat(), "Jam": "{:02d}.00".format(SHIFT_HOUR[s]), "Shift": SHIFT_NAME[s],
                 "Suhu (°C)": e["suhu"], "Kelembapan (%)": e["hum"], "Staff": e["petugas"], "Temuan": e["ket"]}
                for (d, s), e in sorted(month_data(ruang, y, m).items())]
        c3.download_button("⬇️ Data (CSV)", pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig"), fname + ".csv", "text/csv", key="dl_csv")
    with t2:
        alat, m2, y2 = filter_bar("dq", [d["nama"] for d in devices()], "Peralatan")
        dev = dev_by_name(alat)
        sh2 = build_qc_form(dev, y2, m2, qc_month(dev, y2, m2))
        _preview(sh2)
        fname2 = "Daily_QC_{}_{}_{:02d}".format(alat.replace(" ", "_"), y2, m2)
        q1, q2, q3 = st.columns(3)
        try:
            q1.download_button("⬇️ PDF (peralatan ini)", sheets_to_pdf([sh2], "Form Daily QC - " + alat), fname2 + ".pdf", "application/pdf", key="dq_pdf")
            allq = [build_qc_form(d, y2, m2, qc_month(d, y2, m2)) for d in devices()]
            q2.download_button("⬇️ PDF (semua peralatan)", sheets_to_pdf(allq, "Form Daily QC - Semua Peralatan"),
                               "Daily_QC_SEMUA_{}_{:02d}.pdf".format(y2, m2), "application/pdf", key="dq_pdf_all")
        except ImportError:
            st.error("Library reportlab belum terpasang. Jalankan:  pip install reportlab")
        qd = qc_month(dev, y2, m2)
        rows2 = []
        for d in sorted(qd):
            e = qd[d]
            rows2.append({"Alat": alat, "Tanggal": dt.date(y2, m2, d).isoformat(), "Staff": e["petugas"],
                          "Kegiatan selesai": len([i for i in dev["items"] if i in e["done"]]), "Total kegiatan": len(dev["items"]),
                          "QC yang belum dikerjakan": "; ".join(item_label(i) for i in dev["items"] if i not in e["done"] and is_daily(i)), "Catatan": e["catatan"]})
        q3.download_button("⬇️ Data (CSV)", pd.DataFrame(rows2).to_csv(index=False).encode("utf-8-sig"), fname2 + ".csv", "text/csv", key="dq_csv")
        st.caption("Format mengikuti Form Digital Daily Quality Control - centang = kegiatan selesai.")
    if is_author():
        with tabs[2]:
            download_qr_tab()


def download_qr_tab():
    if not is_author():
        return
    st.markdown("**QR Code per ruangan** - scan membuka halaman pengisian Suhu & Kelembapan + Daily QC untuk ruangan tersebut ({} ruangan).".format(len(room_names())))
    base = st.text_input("Alamat aplikasi (URL Streamlit)", value=default_base_url(), key="qr_base").strip()
    if "NAMA-APP" in base or not base.startswith("http"):
        st.warning("Ganti dengan alamat asli aplikasi Anda, mis. https://pht-tracker.streamlit.app, lalu cetak QR. Jangan cetak sebelum alamatnya benar.")
    if st.button("Simpan alamat ini", key="qr_base_save"):
        cfg = load_cfg()
        cfg["base_url"] = base.rstrip("/")
        if try_save_cfg(cfg):
            st.success("Alamat disimpan.")
    sheets = build_qr_sheets(base)
    _preview(sheets[0], 900)
    st.download_button("⬇️ PDF stiker QR (semua ruangan, A4)", sheets_to_pdf(sheets, "QR Code Ruangan - " + HOSPITAL, portrait=True),
                       "QR_Ruangan_Mandaya.pdf", "application/pdf", key="qr_pdf")
    st.dataframe(pd.DataFrame([{"Ruangan": r, "Checklist QC": ", ".join(d["nama"] for d in room_devices(r)) or "-", "Link": room_url(base, r)} for r in room_names()]),
                 hide_index=True)
    st.caption("Cetak, laminating, tempel di dekat termometer/higrometer. Link bisa diuji dengan membuka salah satu alamat di tabel.")


# =====================================================================
# 8. NAVIGASI & SIDEBAR
# =====================================================================
_pages = [(page_dashboard, "Dashboard", "📊", "dashboard", True),
          (page_ruang, "Isi Ruangan (QR)", "🔳", "ruang", False)]
if is_author():   # backdate manual (tanggal sebelum hari ini) hanya untuk author
    _pages.append((page_backdate, "Input Backdate (manual)", "🕘", "backdate", False))
_pages.append((page_download, "Download & QR" if is_author() else "Download Formulir", "⬇️", "download", False))
pg = st.navigation([st.Page(f, title="{}. {}".format(i + 1, t), icon=ic, url_path=u, default=d) for i, (f, t, ic, u, d) in enumerate(_pages)])

with st.sidebar:
    if is_author():
        with st.expander("💾 Penyimpanan data", expanded=not db_on()):
            if st.button("Tes koneksi database", key="db_test"):
                ok, msg = storage_test()
                (st.success if ok else st.error)(msg)
            else:
                st.caption("Supabase: terhubung" if db_on() else "Mode CSV lokal (TIDAK permanen)")
        with st.expander("Pengaturan ruang & staff"):
            cfg = load_cfg()
            nm = st.text_input("Nama ruang baru / ruang yang diubah", key="cfg_nm")
            tl, th = SUHU_NORMAL
            st.caption("Suhu normal {:g}–{:g} °C berlaku sama untuk semua ruangan (diatur di SUHU_NORMAL pada kode).".format(tl, th))
            c1, c2 = st.columns(2)
            hl = c1.number_input("Kelembapan min (%)", value=40.0, step=1.0, key="cfg_hl")
            hh = c2.number_input("Kelembapan maks (%)", value=60.0, step=1.0, key="cfg_hh")
            if st.button("Simpan ruang", key="cfg_add"):
                if nm.strip():
                    cfg["rooms"] = [r for r in cfg["rooms"] if r["nama"] != nm.strip()] + [dict(nama=nm.strip(), t_lo=tl, t_hi=th, h_lo=hl, h_hi=hh, tekanan=False)]
                    if try_save_cfg(cfg):
                        st.success("Ruang disimpan.")
            stf = st.text_input("Inisial staff (pisahkan koma)", value=",".join(cfg["staff"]), key="cfg_staff")
            if st.button("Simpan staff", key="cfg_staff_btn"):
                cfg["staff"] = [x.strip().upper() for x in stf.split(",") if x.strip()] or cfg["staff"]
                if try_save_cfg(cfg):
                    st.success("Staff disimpan.")
        with st.expander("Backup & restore data"):
            if db_on():
                st.caption("Data tersimpan permanen di Supabase. Backup CSV tetap berguna sebagai salinan tambahan (mis. akhir bulan).")
            else:
                st.caption("PERINGATAN: mode CSV lokal, data bisa hilang saat aplikasi restart. Hubungkan Supabase, dan backup rutin.")
            for label, table, path, cols, keys, k in (("Suhu & kelembapan", "entries", ENTRY_FILE, ECOLS, ["ruang", "tanggal", "shift"], "th"),
                                                      ("QC peralatan", "qc_entries", QC_FILE, QCOLS, ["alat", "tanggal"], "qc")):
                full = load_entries() if k == "th" else load_qc()
                st.download_button("⬇️ Backup {} (CSV)".format(label), full.to_csv(index=False).encode("utf-8-sig"),
                                   os.path.basename(path), "text/csv", key="bk_" + k)
                up = st.file_uploader("Restore {}".format(label), type="csv", key="up_" + k)
                if up is not None and st.button("Restore {} (MENGGANTI seluruh data saat ini)".format(label), key="rs_" + k):
                    try:
                        new = pd.read_csv(up, dtype=str, encoding="utf-8-sig").fillna("")
                        miss = [c for c in cols if c not in new.columns]
                        if miss:
                            st.error("Kolom tidak lengkap: " + ", ".join(miss))
                        else:
                            n_rows = replace_table(table, path, new, cols, keys)
                            st.success("Data {} dipulihkan ({} baris).".format(label, n_rows))
                    except StorageError as ex:
                        st.error("Gagal restore: {}".format(ex))
                    except Exception as ex:
                        st.error("Gagal restore: {}".format(ex))
    author_box()

pg.run()
