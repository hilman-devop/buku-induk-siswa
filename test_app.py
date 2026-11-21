"""Smoke test aplikasi buku induk: python test_app.py"""
import io
import os
import tempfile
import unittest
from pathlib import Path

TMP = tempfile.mkdtemp(prefix="buku-induk-test-")
os.environ["BUKU_INDUK_DB"] = str(Path(TMP) / "test.db")

import app as appmod  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402


class BukuIndukTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        appmod.app.config.update(TESTING=True)
        with appmod.app.app_context():
            appmod.init_db()

    def setUp(self):
        self.c = appmod.app.test_client()
        with appmod.app.app_context():
            conn = appmod.db()
            conn.execute("DELETE FROM nilai")
            conn.execute("DELETE FROM siswa")
            conn.execute("DELETE FROM users WHERE username <> 'admin'")
            conn.commit()
        r = self.c.post("/login", data={"username": "admin", "password": "admin123"})
        self.assertEqual(r.status_code, 302, "login admin gagal")

    # -- auth
    def test_redirect_tanpa_login(self):
        c = appmod.app.test_client()
        r = c.get("/siswa")
        self.assertEqual(r.status_code, 302)
        self.assertIn("/login", r.headers["Location"])

    def test_login_salah(self):
        c = appmod.app.test_client()
        r = c.post("/login", data={"username": "admin", "password": "salah"})
        self.assertIn(b"salah", r.data)

    # -- siswa
    def test_tambah_dan_detail_siswa(self):
        r = self.c.post("/siswa/baru", data={
            "nis": "1001", "nama": "Budi Santoso", "jenis_kelamin": "Laki-laki",
            "kelas": "7A", "tanggal_lahir": "2011-02-03", "agama": "Islam",
            "nama_ayah": "Slamet", "pekerjaan_ayah": "Petani",
        }, follow_redirects=True)
        self.assertEqual(r.status_code, 200)
        self.assertIn("Budi Santoso".encode(), r.data)
        self.assertIn(b"1001", r.data)
        r = self.c.get("/siswa?q=Budi")
        self.assertIn(b"Budi Santoso", r.data)

    def test_nis_duplikat_ditolak(self):
        data = {"nis": "1002", "nama": "Ani"}
        self.c.post("/siswa/baru", data=data)
        r = self.c.post("/siswa/baru", data={**data, "nama": "Ani Lain"},
                        follow_redirects=True)
        self.assertIn("sudah dipakai".encode(), r.data)

    def test_nilai_dan_buku_induk(self):
        self.c.post("/siswa/baru", data={"nis": "1003", "nama": "Citra", "kelas": "8B"})
        with appmod.app.app_context():
            sid = appmod.db().execute("SELECT id FROM siswa WHERE nis='1003'").fetchone()["id"]
        r = self.c.post(f"/siswa/{sid}/nilai",
                        data={"mapel": "Matematika", "semester": "1", "nilai": "88,5"},
                        follow_redirects=True)
        self.assertIn(b"88", r.data)
        r = self.c.get(f"/buku-induk/{sid}")
        self.assertEqual(r.status_code, 200)
        self.assertIn(b"Matematika", r.data)
        self.assertIn(b"Buku Induk Siswa", r.data)
        # upsert nilai, bukan duplikat
        self.c.post(f"/siswa/{sid}/nilai",
                    data={"mapel": "Matematika", "semester": "1", "nilai": "90"})
        with appmod.app.app_context():
            jml = appmod.db().execute(
                "SELECT COUNT(*) c FROM nilai WHERE siswa_id=?", (sid,)).fetchone()["c"]
        self.assertEqual(jml, 1)

    # -- excel
    def test_export_siswa(self):
        self.c.post("/siswa/baru", data={"nis": "2001", "nama": "Dewi", "kelas": "9A"})
        r = self.c.get("/export/siswa")
        self.assertEqual(r.status_code, 200)
        self.assertIn("spreadsheetml", r.headers["Content-Type"])
        wb = load_workbook(io.BytesIO(r.data))
        ws = wb.active
        header = [c.value for c in ws[1]]
        self.assertEqual(header[:3], ["NIS", "NISN", "NAMA"])
        self.assertEqual(ws.cell(row=2, column=1).value, "2001")

    def test_export_nilai_template(self):
        self.c.post("/siswa/baru", data={"nis": "2002", "nama": "Eko"})
        r = self.c.get("/export/nilai?semester=2")
        wb = load_workbook(io.BytesIO(r.data))
        ws = wb.active
        self.assertEqual([c.value for c in ws[1]], ["NIS", "NAMA", "MAPEL", "SEMESTER", "NILAI"])
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        self.assertGreater(len(rows), 1)
        self.assertTrue(all(r[3] == "2" for r in rows))

    def _xlsx(self, header, rows):
        wb = Workbook()
        ws = wb.active
        ws.append(header)
        for r in rows:
            ws.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def test_import_siswa_baru_dan_update(self):
        f = self._xlsx(["NO INDUK", "NAMA", "JK", "KELAS", "TANGGAL LAHIR"],
                       [["3001", "Fajar", "L", "7A", "2011-05-01"],
                        ["3002", "Gita", "P", "7A", "01/06/2011"]])
        r = self.c.post("/import", data={"kind": "siswa", "file": (f, "siswa.xlsx")},
                        content_type="multipart/form-data", follow_redirects=True)
        self.assertIn("2 baru".encode(), r.data)
        with appmod.app.app_context():
            rows = appmod.db().execute(
                "SELECT * FROM siswa ORDER BY nis").fetchall()
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["tanggal_lahir"], "2011-05-01")
        self.assertEqual(rows[1]["tanggal_lahir"], "2011-06-01")
        self.assertEqual(rows[1]["jenis_kelamin"], "Perempuan")

        # import ulang: NIS 3001 update nama, 3003 baru
        f2 = self._xlsx(["NIS", "NAMA", "KELAS"],
                        [["3001", "Fajar Revisi", "7B"], ["3003", "Hadi", "7A"]])
        r = self.c.post("/import", data={"kind": "siswa", "file": (f2, "siswa2.xlsx")},
                        content_type="multipart/form-data", follow_redirects=True)
        self.assertIn("1 baru".encode(), r.data)
        self.assertIn("1 diperbarui".encode(), r.data)
        with appmod.app.app_context():
            conn = appmod.db()
            self.assertEqual(len(conn.execute("SELECT * FROM siswa").fetchall()), 3)
            self.assertEqual(conn.execute(
                "SELECT nama FROM siswa WHERE nis='3001'").fetchone()["nama"], "Fajar Revisi")

    def test_import_nilai(self):
        self.c.post("/siswa/baru", data={"nis": "4001", "nama": "Indah"})
        f = self._xlsx(["NIS", "MAPEL", "SEMESTER", "NILAI"],
                       [["4001", "IPA", 1, 85], ["4001", "IPS", 1, 79],
                        ["9999", "IPA", 1, 90]])
        r = self.c.post("/import", data={"kind": "nilai", "file": (f, "nilai.xlsx")},
                        content_type="multipart/form-data", follow_redirects=True)
        self.assertIn("2 nilai masuk".encode(), r.data)
        self.assertIn("1 baris dilewati".encode(), r.data)

    def test_import_file_rusak(self):
        r = self.c.post("/import", data={"kind": "siswa", "file": (
            io.BytesIO(b"bukan excel"), "x.xlsx")},
            content_type="multipart/form-data", follow_redirects=True)
        self.assertIn("Gagal membaca file".encode(), r.data)

    def test_cetak_semua_urut_nomor_induk(self):
        for nis, nama in [("5002", "Kedua"), ("5001", "Pertama")]:
            self.c.post("/siswa/baru", data={"nis": nis, "nama": nama})
        self.c.get("/")  # habiskan flash agar tidak ikut terhitung
        r = self.c.get("/buku-induk/cetak")
        html = r.data.decode()
        self.assertEqual(r.status_code, 200)
        self.assertLess(html.index("Pertama"), html.index("Kedua"))

    # -- pengaturan, pengguna, profil
    def test_pengaturan_tersimpan(self):
        self.c.post("/siswa/baru", data={"nis": "6001", "nama": "Joko"})
        r = self.c.post("/pengaturan", data={
            "nama_sekolah": "SMP Test", "npsn": "123", "alamat": "Jl. A",
            "kota": "Ponorogo", "kepala_sekolah": "Pak Kepala",
            "tahun_ajaran": "2026/2027", "mapel": "Matematika, IPA",
        }, follow_redirects=True)
        self.assertIn(b"SMP Test", r.data)
        r = self.c.get("/export/nilai")
        wb = load_workbook(io.BytesIO(r.data))
        mapel = {row[2] for row in wb.active.iter_rows(min_row=2, values_only=True)}
        self.assertEqual(mapel, {"Matematika", "IPA"})

    def test_pengguna_crud(self):
        r = self.c.post("/pengguna", data={"username": "op1", "nama": "Operator 1",
                                           "password": "rahasia1", "role": "operator"},
                        follow_redirects=True)
        self.assertIn(b"op1", r.data)
        # login dengan user baru
        c2 = appmod.app.test_client()
        r = c2.post("/login", data={"username": "op1", "password": "rahasia1"})
        self.assertEqual(r.status_code, 302)
        # reset password
        with appmod.app.app_context():
            uid = appmod.db().execute("SELECT id FROM users WHERE username='op1'").fetchone()["id"]
        self.c.post(f"/pengguna/{uid}/reset", data={"password": "baru1234"},
                    follow_redirects=True)
        c3 = appmod.app.test_client()
        r = c3.post("/login", data={"username": "op1", "password": "baru1234"})
        self.assertEqual(r.status_code, 302)

    def test_profil_ganti_password(self):
        r = self.c.post("/profil", data={"nama": "Admin Baru", "lama": "admin123",
                                         "baru": "admin456"}, follow_redirects=True)
        self.assertIn("diperbarui".encode(), r.data)
        c = appmod.app.test_client()
        self.assertEqual(c.post("/login", data={"username": "admin", "password": "admin456"}).status_code, 302)
        # kembalikan
        self.c.post("/profil", data={"nama": "Administrator", "lama": "admin456",
                                     "baru": "admin123"})

    def test_dashboard_dan_halaman_lain(self):
        for url in ["/", "/siswa", "/buku-induk", "/import", "/pengguna",
                    "/pengaturan", "/profil"]:
            r = self.c.get(url)
            self.assertEqual(r.status_code, 200, f"{url} -> {r.status_code}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
