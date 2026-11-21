"""Aplikasi Buku Induk Siswa (web) — Flask + SQLite + openpyxl.

Fitur: login, dashboard, data siswa, nilai, export/import Excel,
buku induk siap cetak/PDF, user aplikasi, pengaturan, profil.

Jalan: .venv/Scripts/python.exe app.py  ->  http://127.0.0.1:5000
"""
from __future__ import annotations

import io
import os
import re
import sqlite3
import tempfile
from datetime import date, datetime
from functools import wraps
from pathlib import Path

from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   send_file, session, url_for)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.security import check_password_hash, generate_password_hash

BASE = Path(__file__).resolve().parent
_DB_BUNDLED = BASE / "data" / "buku_induk.db"


def _db_path() -> Path:
    """Lokal: data/buku_induk.db. Serverless (VERCEL): /tmp karena FS lain read-only."""
    env = os.environ.get("BUKU_INDUK_DB")
    if env:
        return Path(env)
    if os.environ.get("VERCEL"):
        tmp = Path(tempfile.gettempdir()) / "buku_induk.db"
        if not tmp.exists() and _DB_BUNDLED.exists():
            tmp.write_bytes(_DB_BUNDLED.read_bytes())
        return tmp
    return _DB_BUNDLED


DB_PATH = _db_path()
DEFAULT_MAPEL = "Pendidikan Pancasila,Bahasa Indonesia,Matematika,IPA,IPS,Bahasa Inggris"

app = Flask(__name__)


def _secret_key() -> bytes:
    """Secret persisten di data/secret.key supaya sesi tidak putus tiap restart."""
    env = os.environ.get("SECRET_KEY")
    if env:
        return env.encode()
    f = BASE / "data" / "secret.key"
    if f.exists():
        return f.read_bytes()
    k = os.urandom(32)
    try:
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_bytes(k)
    except OSError:
        # FS read-only (Vercel): set env SECRET_KEY agar sesi tidak putus antar instance.
        pass
    return k


app.secret_key = _secret_key()

SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT UNIQUE NOT NULL,
  nama TEXT NOT NULL,
  password TEXT NOT NULL,
  role TEXT NOT NULL DEFAULT 'operator'
);
CREATE TABLE IF NOT EXISTS settings(
  id INTEGER PRIMARY KEY CHECK (id = 1),
  nama_sekolah TEXT, npsn TEXT, alamat TEXT, kota TEXT,
  kepala_sekolah TEXT, tahun_ajaran TEXT, mapel TEXT
);
CREATE TABLE IF NOT EXISTS siswa(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  nis TEXT UNIQUE NOT NULL,
  nisn TEXT, nama TEXT NOT NULL,
  jenis_kelamin TEXT, tempat_lahir TEXT, tanggal_lahir TEXT,
  agama TEXT, alamat TEXT, no_hp TEXT,
  nama_ayah TEXT, pekerjaan_ayah TEXT, nama_ibu TEXT, pekerjaan_ibu TEXT,
  kelas TEXT, tahun_masuk TEXT, status TEXT DEFAULT 'Aktif', catatan TEXT
);
CREATE TABLE IF NOT EXISTS nilai(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  siswa_id INTEGER NOT NULL REFERENCES siswa(id) ON DELETE CASCADE,
  mapel TEXT NOT NULL, semester TEXT NOT NULL, nilai REAL,
  UNIQUE(siswa_id, mapel, semester)
);
"""

SISWA_FIELDS = [
    ("nis", "NIS"),
    ("nisn", "NISN"),
    ("nama", "NAMA"),
    ("jenis_kelamin", "JENIS KELAMIN"),
    ("tempat_lahir", "TEMPAT LAHIR"),
    ("tanggal_lahir", "TANGGAL LAHIR"),
    ("agama", "AGAMA"),
    ("alamat", "ALAMAT"),
    ("no_hp", "NO HP"),
    ("nama_ayah", "NAMA AYAH"),
    ("pekerjaan_ayah", "PEKERJAAN AYAH"),
    ("nama_ibu", "NAMA IBU"),
    ("pekerjaan_ibu", "PEKERJAAN IBU"),
    ("kelas", "KELAS"),
    ("tahun_masuk", "TAHUN MASUK"),
    ("status", "STATUS"),
    ("catatan", "CATATAN"),
]
SISWA_KEYS = [k for k, _ in SISWA_FIELDS]
SISWA_LABELS = dict(SISWA_FIELDS)
NILAI_COLS = ["NIS", "NAMA", "MAPEL", "SEMESTER", "NILAI"]


# ---------------------------------------------------------------- database
def db() -> sqlite3.Connection:
    if "db" not in g:
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def _close_db(_exc):
    conn = g.pop("db", None)
    if conn:
        conn.close()


def init_db() -> None:
    conn = db()
    conn.executescript(SCHEMA)
    if not conn.execute("SELECT 1 FROM users LIMIT 1").fetchone():
        conn.execute(
            "INSERT INTO users(username, nama, password, role) VALUES (?,?,?,?)",
            ("admin", "Administrator", generate_password_hash("admin123"), "admin"),
        )
    if not conn.execute("SELECT 1 FROM settings LIMIT 1").fetchone():
        conn.execute(
            "INSERT INTO settings(id, nama_sekolah, npsn, alamat, kota, kepala_sekolah,"
            " tahun_ajaran, mapel) VALUES (1,?,?,?,?,?,?,?)",
            ("Sekolah Dunia Mimpi", "", "", "", "", "2025/2026", DEFAULT_MAPEL),
        )
    conn.commit()


if os.environ.get("VERCEL"):  # serverless: tidak ada `python app.py` yang memanggil init_db
    with app.app_context():
        init_db()


def settings_map() -> dict:
    row = db().execute("SELECT * FROM settings WHERE id = 1").fetchone()
    return dict(row) if row else {}


def mapel_list() -> list[str]:
    raw = settings_map().get("mapel") or DEFAULT_MAPEL
    return [m.strip() for m in raw.split(",") if m.strip()]


@app.context_processor
def inject_globals():
    return {"S": settings_map(), "me": session.get("nama"), "role": session.get("role")}


BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]


@app.template_filter("tgl")
def _tgl(v):
    """2011-04-12 -> 12 April 2011 (teks lain dilewatkan apa adanya)."""
    if not v:
        return ""
    try:
        d = datetime.strptime(str(v)[:10], "%Y-%m-%d")
    except ValueError:
        return v
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


# ---------------------------------------------------------------- auth
def login_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if not session.get("uid"):
            return redirect(url_for("login", next=request.path))
        return fn(*a, **kw)

    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "")
        row = db().execute("SELECT * FROM users WHERE username = ?", (u,)).fetchone()
        if row and check_password_hash(row["password"], p):
            session.update(uid=row["id"], nama=row["nama"], role=row["role"])
            return redirect(request.args.get("next") or url_for("dashboard"))
        flash("Username atau password salah.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------- dashboard
@app.route("/")
@login_required
def dashboard():
    conn = db()
    stat = {
        "siswa": conn.execute("SELECT COUNT(*) c FROM siswa").fetchone()["c"],
        "kelas": conn.execute("SELECT COUNT(DISTINCT kelas) c FROM siswa WHERE kelas<>''").fetchone()["c"],
        "nilai": conn.execute("SELECT COUNT(*) c FROM nilai").fetchone()["c"],
        "rata": conn.execute("SELECT AVG(nilai) a FROM nilai").fetchone()["a"],
        "laki": conn.execute("SELECT COUNT(*) c FROM siswa WHERE jenis_kelamin LIKE 'L%'").fetchone()["c"],
        "perempuan": conn.execute("SELECT COUNT(*) c FROM siswa WHERE jenis_kelamin LIKE 'P%'").fetchone()["c"],
    }
    kelas = conn.execute(
        "SELECT kelas, COUNT(*) c FROM siswa GROUP BY kelas ORDER BY kelas"
    ).fetchall()
    terbaru = conn.execute("SELECT * FROM siswa ORDER BY id DESC LIMIT 6").fetchall()
    return render_template("dashboard.html", stat=stat, kelas=kelas, terbaru=terbaru)


# ---------------------------------------------------------------- siswa
@app.route("/siswa")
@login_required
def siswa_list():
    q = request.args.get("q", "").strip()
    kelas = request.args.get("kelas", "").strip()
    sql = "SELECT * FROM siswa WHERE 1=1"
    args: list = []
    if q:
        sql += " AND (nama LIKE ? OR nis LIKE ? OR nisn LIKE ?)"
        args += [f"%{q}%"] * 3
    if kelas:
        sql += " AND kelas = ?"
        args.append(kelas)
    sql += " ORDER BY nis"
    rows = db().execute(sql, args).fetchall()
    semua_kelas = [r["kelas"] for r in db().execute(
        "SELECT DISTINCT kelas FROM siswa WHERE kelas<>'' ORDER BY kelas")]
    return render_template("siswa_list.html", rows=rows, q=q, kelas=kelas, semua_kelas=semua_kelas)


def _form_to_siswa() -> dict:
    data = {k: request.form.get(k, "").strip() for k in SISWA_KEYS}
    data["jenis_kelamin"] = "Laki-laki" if data["jenis_kelamin"].startswith("L") else (
        "Perempuan" if data["jenis_kelamin"].startswith("P") else data["jenis_kelamin"])
    return data


@app.route("/siswa/baru", methods=["GET", "POST"])
@login_required
def siswa_baru():
    if request.method == "POST":
        d = _form_to_siswa()
        if not d["nis"] or not d["nama"]:
            flash("NIS dan Nama wajib diisi.", "error")
        else:
            cols = ", ".join(SISWA_KEYS)
            marks = ", ".join("?" * len(SISWA_KEYS))
            try:
                db().execute(f"INSERT INTO siswa({cols}) VALUES({marks})",
                             [d[k] for k in SISWA_KEYS])
                db().commit()
                flash(f"Siswa {d['nama']} ditambahkan.", "ok")
                return redirect(url_for("siswa_detail", sid=db().execute(
                    "SELECT id FROM siswa WHERE nis=?", (d["nis"],)).fetchone()["id"]))
            except sqlite3.IntegrityError:
                flash("NIS sudah dipakai siswa lain.", "error")
    return render_template("siswa_form.html", siswa=None)


@app.route("/siswa/<int:sid>")
@login_required
def siswa_detail(sid):
    conn = db()
    siswa = conn.execute("SELECT * FROM siswa WHERE id = ?", (sid,)).fetchone()
    if not siswa:
        abort(404)
    nilai = conn.execute(
        "SELECT * FROM nilai WHERE siswa_id = ? ORDER BY semester, mapel", (sid,)
    ).fetchall()
    per_sem = {}
    for n in nilai:
        per_sem.setdefault(n["semester"], []).append(n)
    return render_template("siswa_detail.html", siswa=siswa, per_sem=per_sem,
                           mapel=mapel_list())


@app.route("/siswa/<int:sid>/edit", methods=["GET", "POST"])
@login_required
def siswa_edit(sid):
    conn = db()
    siswa = conn.execute("SELECT * FROM siswa WHERE id = ?", (sid,)).fetchone()
    if not siswa:
        abort(404)
    if request.method == "POST":
        d = _form_to_siswa()
        if not d["nis"] or not d["nama"]:
            flash("NIS dan Nama wajib diisi.", "error")
        else:
            sets = ", ".join(f"{k} = ?" for k in SISWA_KEYS)
            try:
                conn.execute(f"UPDATE siswa SET {sets} WHERE id = ?",
                             [d[k] for k in SISWA_KEYS] + [sid])
                conn.commit()
                flash("Data siswa diperbarui.", "ok")
                return redirect(url_for("siswa_detail", sid=sid))
            except sqlite3.IntegrityError:
                flash("NIS sudah dipakai siswa lain.", "error")
    return render_template("siswa_form.html", siswa=siswa)


@app.route("/siswa/<int:sid>/hapus", methods=["POST"])
@login_required
def siswa_hapus(sid):
    db().execute("DELETE FROM siswa WHERE id = ?", (sid,))
    db().commit()
    flash("Data siswa dihapus.", "ok")
    return redirect(url_for("siswa_list"))


@app.route("/siswa/<int:sid>/nilai", methods=["POST"])
@login_required
def nilai_simpan(sid):
    mapel = request.form.get("mapel", "").strip()
    semester = request.form.get("semester", "1").strip() or "1"
    raw = request.form.get("nilai", "").replace(",", ".").strip()
    if not mapel or not raw:
        flash("Mapel dan nilai wajib diisi.", "error")
        return redirect(url_for("siswa_detail", sid=sid))
    try:
        val = float(raw)
    except ValueError:
        flash("Nilai harus berupa angka.", "error")
        return redirect(url_for("siswa_detail", sid=sid))
    db().execute(
        "INSERT INTO nilai(siswa_id, mapel, semester, nilai) VALUES (?,?,?,?)"
        " ON CONFLICT(siswa_id, mapel, semester) DO UPDATE SET nilai = excluded.nilai",
        (sid, mapel, semester, val),
    )
    db().commit()
    flash(f"Nilai {mapel} semester {semester} disimpan.", "ok")
    return redirect(url_for("siswa_detail", sid=sid))


@app.route("/nilai/<int:nid>/hapus", methods=["POST"])
@login_required
def nilai_hapus(nid):
    row = db().execute("SELECT siswa_id FROM nilai WHERE id = ?", (nid,)).fetchone()
    db().execute("DELETE FROM nilai WHERE id = ?", (nid,))
    db().commit()
    return redirect(url_for("siswa_detail", sid=row["siswa_id"]) if row else url_for("siswa_list"))


# ---------------------------------------------------------------- excel
HDR_FILL = PatternFill("solid", fgColor="1D4ED8")
HDR_FONT = Font(bold=True, color="FFFFFF")
WIDTHS = {"NAMA": 28, "ALAMAT": 34, "CATATAN": 26}


def _sheet(ws, headers: list[str], rows: list[list]):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS.get(headers[c - 1], 16)
    for r in rows:
        ws.append(r)
    ws.freeze_panes = "A2"


def _xlsx(wb: Workbook, name: str):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True, download_name=name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/export/siswa")
@login_required
def export_siswa():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Siswa"
    rows = [[r[k] for k in SISWA_KEYS] for r in
            db().execute("SELECT * FROM siswa ORDER BY nis")]
    _sheet(ws, [lab for _, lab in SISWA_FIELDS], rows)
    return _xlsx(wb, f"data-siswa-{date.today().isoformat()}.xlsx")


@app.route("/export/nilai")
@login_required
def export_nilai():
    semester = request.args.get("semester", "1")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Nilai"
    existing = {(r["siswa_id"], r["mapel"]): r["nilai"] for r in db().execute(
        "SELECT siswa_id, mapel, nilai FROM nilai WHERE semester = ?", (semester,))}
    rows = []
    for s in db().execute("SELECT id, nis, nama FROM siswa ORDER BY nis"):
        for m in mapel_list():
            v = existing.get((s["id"], m))
            rows.append([s["nis"], s["nama"], m, semester, "" if v is None else v])
    _sheet(ws, NILAI_COLS, rows)
    return _xlsx(wb, f"data-nilai-{date.today().isoformat()}.xlsx")


def _norm(h) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(h or "").lower()).strip()


ALIASES = {
    "nis": {"nis", "no induk", "nomor induk", "nomor induk siswa", "no induk siswa"},
    "nisn": {"nisn", "nis n"},
    "nama": {"nama", "nama siswa", "nama lengkap"},
    "jenis_kelamin": {"jenis kelamin", "jk", "l p", "gender"},
    "tempat_lahir": {"tempat lahir", "tmp lahir"},
    "tanggal_lahir": {"tanggal lahir", "tgl lahir", "tanggal"},
    "agama": {"agama"},
    "alamat": {"alamat", "alamat lengkap"},
    "no_hp": {"no hp", "hp", "telepon", "no telp", "whatsapp"},
    "nama_ayah": {"nama ayah", "ayah"},
    "pekerjaan_ayah": {"pekerjaan ayah", "kerja ayah"},
    "nama_ibu": {"nama ibu", "ibu"},
    "pekerjaan_ibu": {"pekerjaan ibu", "kerja ibu"},
    "kelas": {"kelas", "rombel"},
    "tahun_masuk": {"tahun masuk", "angkatan", "tahun"},
    "status": {"status", "status siswa"},
    "catatan": {"catatan", "keterangan"},
}


def _date_str(v) -> str:
    if isinstance(v, (datetime, date)):
        return v.date().isoformat() if isinstance(v, datetime) else v.isoformat()
    s = str(v or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s


def _read_rows(file) -> tuple[list[str], list[list]]:
    wb = load_workbook(file, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(it, [])]
    rows = [[c for c in r] for r in it if any(c not in (None, "") for c in r)]
    return header, rows


@app.route("/import", methods=["GET", "POST"])
@login_required
def import_page():
    if request.method == "POST":
        kind = request.form.get("kind")
        f = request.files.get("file")
        if not f or not f.filename:
            flash("Pilih file Excel (.xlsx) dulu.", "error")
            return redirect(url_for("import_page"))
        try:
            header, rows = _read_rows(f)
        except Exception as e:  # file rusak / bukan xlsx
            flash(f"Gagal membaca file: {e}", "error")
            return redirect(url_for("import_page"))
        cols = {_norm(h): i for i, h in enumerate(header) if h}
        if kind == "siswa":
            flash(*_import_siswa(cols, rows))
        elif kind == "nilai":
            flash(*_import_nilai(cols, rows))
        else:
            flash("Jenis import tidak dikenal.", "error")
        return redirect(url_for("import_page"))
    contoh = {"mapel": mapel_list(), "jumlah": db().execute(
        "SELECT COUNT(*) c FROM siswa").fetchone()["c"]}
    return render_template("import.html", contoh=contoh)


def _col(cols: dict, key: str) -> int | None:
    for alias in ALIASES[key]:
        if alias in cols:
            return cols[alias]
    return None


def _import_siswa(cols: dict, rows: list[list]) -> tuple[str, str]:
    idx = {k: _col(cols, k) for k in SISWA_KEYS}
    if idx["nis"] is None or idx["nama"] is None:
        return ("File tidak punya kolom NIS dan NAMA. Unduh template export dulu.", "error")
    conn = db()
    masuk, update, gagal = 0, 0, 0
    for r in rows:
        def val(k):
            i = idx[k]
            if i is None or i >= len(r) or r[i] is None:
                return ""
            v = r[i]
            if k == "tanggal_lahir":
                return _date_str(v)
            if isinstance(v, float) and v.is_integer():
                return str(int(v))
            return str(v).strip()

        d = {k: val(k) for k in SISWA_KEYS}
        if not d["nis"] or not d["nama"]:
            gagal += 1
            continue
        if not d["jenis_kelamin"]:
            d["jenis_kelamin"] = ""
        elif d["jenis_kelamin"][0].upper() == "L":
            d["jenis_kelamin"] = "Laki-laki"
        elif d["jenis_kelamin"][0].upper() == "P":
            d["jenis_kelamin"] = "Perempuan"
        ada = conn.execute("SELECT 1 FROM siswa WHERE nis = ?", (d["nis"],)).fetchone()
        cols_sql = ", ".join(SISWA_KEYS)
        marks = ", ".join("?" * len(SISWA_KEYS))
        sets = ", ".join(f"{k} = excluded.{k}" for k in SISWA_KEYS if k != "nis")
        conn.execute(
            f"INSERT INTO siswa({cols_sql}) VALUES({marks})"
            f" ON CONFLICT(nis) DO UPDATE SET {sets}",
            [d[k] for k in SISWA_KEYS],
        )
        masuk += 0 if ada else 1
        update += 1 if ada else 0
    conn.commit()
    msg = f"Import data siswa selesai: {masuk} baru, {update} diperbarui"
    if gagal:
        msg += f", {gagal} baris dilewati (NIS/Nama kosong)"
    return (msg + ".", "ok")


def _import_nilai(cols: dict, rows: list[list]) -> tuple[str, str]:
    c_nis, c_mapel, c_sem, c_nilai = (
        _col(cols, "nis"), cols.get("mapel"), cols.get("semester"), cols.get("nilai"))
    if None in (c_nis, c_mapel, c_sem, c_nilai):
        return ("File tidak punya kolom NIS, MAPEL, SEMESTER, NILAI. "
                "Unduh template export nilai dulu.", "error")
    conn = db()
    ok, lewat = 0, 0
    for r in rows:
        def g(i):
            return r[i] if i is not None and i < len(r) else None

        nis = str(g(c_nis) or "").strip()
        if isinstance(g(c_nis), float) and float(g(c_nis)).is_integer():
            nis = str(int(g(c_nis)))
        mapel = str(g(c_mapel) or "").strip()
        sem = str(g(c_sem) or "").strip()
        if isinstance(g(c_sem), float) and float(g(c_sem)).is_integer():
            sem = str(int(g(c_sem)))
        raw = g(c_nilai)
        if not nis or not mapel or raw in (None, ""):
            lewat += 1
            continue
        try:
            nval = float(str(raw).replace(",", "."))
        except ValueError:
            lewat += 1
            continue
        siswa = conn.execute("SELECT id FROM siswa WHERE nis = ?", (nis,)).fetchone()
        if not siswa:
            lewat += 1
            continue
        conn.execute(
            "INSERT INTO nilai(siswa_id, mapel, semester, nilai) VALUES (?,?,?,?)"
            " ON CONFLICT(siswa_id, mapel, semester) DO UPDATE SET nilai = excluded.nilai",
            (siswa["id"], mapel, sem or "1", nval),
        )
        ok += 1
    conn.commit()
    msg = f"Import data nilai selesai: {ok} nilai masuk"
    if lewat:
        msg += f", {lewat} baris dilewati (NIS tidak dikenal / nilai kosong)"
    return (msg + ".", "ok")


# ---------------------------------------------------------------- buku induk
@app.route("/buku-induk")
@login_required
def buku_induk_list():
    kelas = request.args.get("kelas", "")
    sql = "SELECT s.*, (SELECT COUNT(*) FROM nilai n WHERE n.siswa_id = s.id) jml_nilai FROM siswa s"
    args = []
    if kelas:
        sql += " WHERE s.kelas = ?"
        args.append(kelas)
    sql += " ORDER BY CAST(s.nis AS INTEGER), s.nis"
    rows = db().execute(sql, args).fetchall()
    semua_kelas = [r["kelas"] for r in db().execute(
        "SELECT DISTINCT kelas FROM siswa WHERE kelas<>'' ORDER BY kelas")]
    return render_template("buku_induk_list.html", rows=rows, kelas=kelas, semua_kelas=semua_kelas)


@app.route("/buku-induk/<int:sid>")
@login_required
def buku_induk(sid):
    data = _lembar_data(sid)
    return render_template("buku_induk_cetak.html", daftar=[data], satu=True)


@app.route("/buku-induk/cetak")
@login_required
def buku_induk_cetak():
    kelas = request.args.get("kelas", "")
    sql = "SELECT id FROM siswa"
    args = []
    if kelas:
        sql += " WHERE kelas = ?"
        args.append(kelas)
    sql += " ORDER BY CAST(nis AS INTEGER), nis"
    ids = [r["id"] for r in db().execute(sql, args)]
    return render_template("buku_induk_cetak.html",
                           daftar=[_lembar_data(i) for i in ids], satu=False)


def _lembar_data(sid: int) -> dict:
    conn = db()
    siswa = conn.execute("SELECT * FROM siswa WHERE id = ?", (sid,)).fetchone()
    if not siswa:
        abort(404)
    nilai = conn.execute(
        "SELECT * FROM nilai WHERE siswa_id = ? ORDER BY semester, mapel", (sid,)).fetchall()
    per_sem: dict[str, list] = {}
    for n in nilai:
        per_sem.setdefault(n["semester"], []).append(n)
    rata = conn.execute("SELECT AVG(nilai) a FROM nilai WHERE siswa_id = ?",
                        (sid,)).fetchone()["a"]
    return {"s": siswa, "per_sem": per_sem, "rata": rata}


# ---------------------------------------------------------------- pengguna
@app.route("/pengguna", methods=["GET", "POST"])
@login_required
def pengguna():
    conn = db()
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        n = request.form.get("nama", "").strip() or u
        p = request.form.get("password", "")
        r = request.form.get("role", "operator")
        if not u or len(p) < 6:
            flash("Username wajib, password minimal 6 karakter.", "error")
        else:
            try:
                conn.execute("INSERT INTO users(username, nama, password, role) VALUES (?,?,?,?)",
                             (u, n, generate_password_hash(p), r))
                conn.commit()
                flash(f"User {u} ditambahkan.", "ok")
            except sqlite3.IntegrityError:
                flash("Username sudah dipakai.", "error")
        return redirect(url_for("pengguna"))
    users = conn.execute("SELECT * FROM users ORDER BY id").fetchall()
    return render_template("pengguna.html", users=users)


@app.route("/pengguna/<int:uid>/reset", methods=["POST"])
@login_required
def pengguna_reset(uid):
    p = request.form.get("password", "")
    if len(p) < 6:
        flash("Password baru minimal 6 karakter.", "error")
    else:
        db().execute("UPDATE users SET password = ? WHERE id = ?",
                     (generate_password_hash(p), uid))
        db().commit()
        flash("Password direset.", "ok")
    return redirect(url_for("pengguna"))


@app.route("/pengguna/<int:uid>/hapus", methods=["POST"])
@login_required
def pengguna_hapus(uid):
    if uid == session.get("uid"):
        flash("Tidak bisa menghapus akun yang sedang dipakai.", "error")
    elif db().execute("SELECT COUNT(*) c FROM users").fetchone()["c"] <= 1:
        flash("Minimal harus ada satu user.", "error")
    else:
        db().execute("DELETE FROM users WHERE id = ?", (uid,))
        db().commit()
        flash("User dihapus.", "ok")
    return redirect(url_for("pengguna"))


# ---------------------------------------------------------------- pengaturan & profil
@app.route("/pengaturan", methods=["GET", "POST"])
@login_required
def pengaturan():
    if request.method == "POST":
        d = {k: request.form.get(k, "").strip() for k in
             ("nama_sekolah", "npsn", "alamat", "kota", "kepala_sekolah",
              "tahun_ajaran", "mapel")}
        db().execute(
            "UPDATE settings SET nama_sekolah=?, npsn=?, alamat=?, kota=?,"
            " kepala_sekolah=?, tahun_ajaran=?, mapel=? WHERE id=1",
            tuple(d.values()),
        )
        db().commit()
        flash("Pengaturan disimpan.", "ok")
        return redirect(url_for("pengaturan"))
    return render_template("pengaturan.html")


@app.route("/profil", methods=["GET", "POST"])
@login_required
def profil():
    conn = db()
    user = conn.execute("SELECT * FROM users WHERE id = ?", (session["uid"],)).fetchone()
    if request.method == "POST":
        nama = request.form.get("nama", "").strip() or user["nama"]
        lama, baru = request.form.get("lama", ""), request.form.get("baru", "")
        if baru:
            if not check_password_hash(user["password"], lama):
                flash("Password lama salah.", "error")
                return redirect(url_for("profil"))
            if len(baru) < 6:
                flash("Password baru minimal 6 karakter.", "error")
                return redirect(url_for("profil"))
            conn.execute("UPDATE users SET nama=?, password=? WHERE id=?",
                         (nama, generate_password_hash(baru), user["id"]))
            flash("Profil dan password diperbarui.", "ok")
        else:
            conn.execute("UPDATE users SET nama=? WHERE id=?", (nama, user["id"]))
            flash("Profil diperbarui.", "ok")
        conn.commit()
        session["nama"] = nama
        return redirect(url_for("profil"))
    return render_template("profil.html", user=user)


@app.errorhandler(404)
def not_found(_e):
    return render_template("404.html"), 404


def _seed_demo():
    contoh = [
        ("2026001", "0071234567", "Ayu Lestari", "Perempuan", "Ponorogo", "2011-04-12",
         "Islam", "Jl. Diponegoro 12", "081234567801", "Sutrisno", "Petani", "Sriyati", "Ibu Rumah Tangga",
         "7A", "2026", "Aktif", ""),
        ("2026002", "0071234568", "Bagas Pratama", "Laki-laki", "Madiun", "2011-07-03",
         "Islam", "Jl. Ahmad Yani 5", "081234567802", "Wahyu", "Wiraswasta", "Endang", "Guru",
         "7A", "2026", "Aktif", ""),
        ("2026003", "0071234569", "Citra Dewi", "Perempuan", "Ponorogo", "2011-01-22",
         "Kristen", "Jl. Merdeka 44", "081234567803", "Petrus", "PNS", "Maria", "Bidan",
         "7B", "2026", "Aktif", ""),
        ("2026004", "0071234570", "Dimas Saputra", "Laki-laki", "Ngawi", "2010-11-09",
         "Islam", "Jl. Pahlawan 7", "081234567804", "Agus", "Buruh", "Yuni", "Pedagang",
         "8A", "2025", "Aktif", ""),
        ("2026005", "0071234571", "Eka Puspita", "Perempuan", "Ponorogo", "2010-05-30",
         "Hindu", "Jl. Kartini 3", "081234567805", "Made", "Peternak", "Niluh", "Penjahit",
         "8A", "2025", "Aktif", ""),
        ("2026006", "0071234572", "Fajar Nugroho", "Laki-laki", "Ponorogo", "2009-09-14",
         "Islam", "Jl. Sudirman 21", "081234567806", "Bambang", "Pedagang", "Wati", "Ibu Rumah Tangga",
         "9B", "2024", "Aktif", ""),
    ]
    conn = db()
    for c in contoh:
        conn.execute(f"INSERT OR IGNORE INTO siswa({', '.join(SISWA_KEYS)})"
                     f" VALUES({', '.join('?' * len(SISWA_KEYS))})", list(c))
    conn.commit()
    mapel = mapel_list()
    for s in conn.execute("SELECT id FROM siswa").fetchall():
        for i, m in enumerate(mapel):
            conn.execute("INSERT OR IGNORE INTO nilai(siswa_id, mapel, semester, nilai)"
                         " VALUES (?,?,?,?)", (s["id"], m, "1", 78 + (i * 3 + s["id"]) % 18))
    conn.commit()


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=5000)
    ap.add_argument("--demo", action="store_true", help="isi contoh data siswa")
    a = ap.parse_args()
    with app.app_context():
        init_db()
        if a.demo and not db().execute("SELECT 1 FROM siswa LIMIT 1").fetchone():
            _seed_demo()
            print("Contoh data ditambahkan.")
    print(f"Buku Induk Siswa -> http://127.0.0.1:{a.port}  (admin / admin123)")
    app.run(host="127.0.0.1", port=a.port, debug=False)
