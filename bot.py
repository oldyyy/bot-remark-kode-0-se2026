"""
Bot pemberian remark #0 (NTB) di FASIH SE2026, berdasarkan rekap dashboard
fasih-scraper/output/rasio_ntb/rekap_remark_ntb_kode0_*.xlsx (sheet "Remark #0").

Sasaran (per usaha = assignment_id + index1, kolom `kunci`):
  - remark_dashboard = "#0"     dan di FASIH belum ada "#0" (kolom remark_fasih), atau
  - remark_dashboard = "#DC_00" dan di FASIH belum ada remark sama sekali (jumlah_remark_fasih = 0).

Alur per assignment:
  1. Buka /edit. Terkunci (tanpa tombol Kirim) -> Revoke -> buka ulang /edit (hanya --live).
  2. Untuk tiap usaha sasaran: Blok "SE2026 - L BLOK II" -> kartu usaha (dicocokkan dengan nama usaha)
     -> pertanyaan 27.a "Nilai produksi/pendapatan/penjualan barang dan jasa"
        (31.a bila usaha hanya mengisi bulanan) -> ikon Catatan -> tulis "#0" -> Simpan.
     Bila catatan "#0" sudah ada di rincian itu -> tidak ditulis ulang.
  3. Ada catatan baru -> Kirim (GALAT yang dikenal diperbaiki otomatis; GALAT lain -> dilewati & dicatat).
     Tidak ada yang ditulis -> tidak dikirim.

Pemakaian (lihat README.md):
  buka_chrome.bat                 lalu login SSO di jendela Chrome itu
  py bot.py --dry-run --limit 3   hanya memeriksa, tidak menulis / mengirim
  py bot.py --live --limit 3      eksekusi
  py bot.py --status              ringkasan progres dari log
Opsi: --file XLSX, --port 9340, --shard K/N, --ids a,b, --ulang, --cepat
"""
import argparse
import csv
import difflib
import json
import random
import re
import sys
import time
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)

SURVEY_ID = "fd68e454-ba45-4b85-8205-f3bf777ded24"
FASIH = "https://fasih-sm.bps.go.id"
EDIT_URL = FASIH + "/app/assignment/{sid}/{aid}/edit"
ASSIGNMENT_API = "/app/api/assignment-general/api/assignment/get-by-assignment-id?assignmentId={aid}"
BASE = Path(__file__).resolve().parent
LOG_DIR = BASE / "logs"
INPUT_DIRS = [BASE / "input", BASE.parents[1] / "fasih-scraper" / "output" / "rasio_ntb"]
POLA_FILE = "rekap_remark_ntb_kode0_*.xlsx"

TEKS_REMARK = "#0"
BLOK_USAHA = "SE2026 - L BLOK II"
LABEL_TAHUNAN = re.compile(r"^27\.\s*a\.\s*Nilai\s+(produksi|pendapatan|penjualan)", re.I)
LABEL_BULANAN = re.compile(r"^31\.\s*a\.\s*Nilai\s+(produksi|pendapatan|penjualan)", re.I)
ADA_KODE0 = re.compile(r"(^|[^\w])#0(?!\d)")  # "#0" tapi bukan "#01"/"#DC_00"
NAMA_MIN_SKOR = 0.75
STATUS_FINAL = ("SELESAI", "SUDAH_ADA")  # tidak diulang saat bot dijalankan lagi (kecuali --ulang)

JEDA = {  # detik, (min, maks) - selalu diacak
    "buka": (2.0, 4.5), "navigasi": (1.0, 2.5), "field": (0.8, 2.0), "tempel": (0.5, 1.3),
    "kirim": (1.5, 3.5), "antar": (8.0, 18.0), "istirahat_tiap": (20, 40), "istirahat": (90.0, 240.0),
    "baca": (0.6, 1.4),  # jeda singkat menunggu tampilan (pengganti jeda tetap)
}


def jeda(kind, cepat=False):
    lo, hi = JEDA[kind]
    if cepat and kind in ("antar", "istirahat"):
        lo, hi = lo / 4, hi / 4
    # sesekali (±1 dari 8) jeda lebih panjang, seperti orang yang berhenti sejenak
    if kind in ("navigasi", "field", "antar") and random.random() < 0.12:
        hi *= 1.8
    time.sleep(random.uniform(lo, hi))


def norm(s):
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", str(s or "").upper())).strip()


# ---------------------------------------------------------------- data sasaran
def cari_file(arg):
    if arg:
        return Path(arg)
    for d in INPUT_DIRS:
        files = sorted(d.glob(POLA_FILE), key=lambda p: p.stat().st_mtime) if d.exists() else []
        if files:
            return files[-1]
    raise SystemExit(f"!! File {POLA_FILE} tidak ditemukan. Taruh di folder input/ atau pakai --file.")


def load_targets(path):
    """assignment_id -> [usaha sasaran]"""
    ws = openpyxl.load_workbook(path, read_only=True)["Remark #0"]
    rows = ws.iter_rows(values_only=True)
    head = [str(h or "") for h in next(rows)]
    groups, n_total, n_lewat = OrderedDict(), 0, 0
    for r in rows:
        d = dict(zip(head, r))
        if not d.get("kunci"):
            continue
        n_total += 1
        dash = str(d.get("remark_dashboard") or "").strip().upper()
        fasih = str(d.get("remark_fasih") or "")
        jml = int(d.get("jumlah_remark_fasih") or 0)
        sasaran = (dash == "#0" and not ADA_KODE0.search(fasih)) or (dash == "#DC_00" and jml == 0)
        if not sasaran:
            n_lewat += 1
            continue
        aid, _, idx = str(d["kunci"]).partition("|")
        groups.setdefault(aid.strip(), []).append({
            "kunci": str(d["kunci"]).strip(), "index1": idx.strip(), "nama_usaha": str(d.get("nama_usaha") or ""),
            "kbli": str(d.get("kbli") or ""), "kecamatan": d.get("kecamatan") or "", "desa": d.get("desa") or "",
            "remark_dashboard": dash})
    print(f"File: {path.name} | {n_total} usaha, {n_total - n_lewat} sasaran di {len(groups)} assignment")
    return groups


LEWATI_FILE = LOG_DIR / "lewati_manual.txt"


def baca_lewati():
    """assignment_id yang sengaja tidak diproses bot (mis. halaman edit-nya memutus sesi) -> dikerjakan manual.
    Satu id per baris; teks setelah '#' diabaikan."""
    if not LEWATI_FILE.exists():
        return set()
    return {ln.split("#")[0].strip() for ln in LEWATI_FILE.read_text(encoding="utf-8").splitlines() if ln.split("#")[0].strip()}


def baca_rekap():
    """kunci -> status final dari semua logs/rekap*.csv"""
    done = {}
    for f in sorted(LOG_DIR.glob("rekap*.csv")):
        with open(f, encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                if r.get("status") in STATUS_FINAL and r.get("kunci"):
                    done[r["kunci"]] = r["status"]
    return done


# ---------------------------------------------------------------- halaman
JS_FIELDS = r"""
() => {
  const main = document.querySelector('main') || document.body;
  return [...main.querySelectorAll('input[type=text], textarea')].filter(i => i.offsetParent).map(i => {
    let p = i.parentElement, label = '';
    for (let k = 0; k < 10 && p; k++, p = p.parentElement) {
      const s = (p.innerText || '').trim();
      if (s && s.length < 400) label = s.split('\n')[0]; else if (s) break;
    }
    return {id: i.id, label: label.trim(), value: i.value, disabled: i.disabled || i.readOnly};
  });
}
"""

JS_CARDS = r"""
() => [...document.querySelectorAll('[data-nested-view]')].filter(c => c.offsetParent).map((c, idx) => {
  const t = c.innerText.split('\n').map(s => s.trim()).filter(Boolean);
  return {idx, no: t[0], nama: t[1] || ''};
})
"""


class FormError(Exception):
    pass


class PerluRevoke(FormError):
    pass


class PerluManual(FormError):
    """Tidak bisa dipastikan otomatis (mis. kartu usaha tidak cocok) -> ditangani manual."""


class Berhenti(Exception):
    """Kondisi tak terduga setelah tindakan tulis -> seluruh run dihentikan."""


class SesiHabis(Exception):
    """Layar login tampil & sesi tidak aktif: bot menunggu login ulang MANUAL, lalu melanjutkan."""


class Form:
    def __init__(self, page, cepat):
        self.page = page
        self.cepat = cepat

    def body(self):
        try:
            return self.page.inner_text("body", timeout=5000)
        except Exception:
            return ""

    def wait(self, cond, timeout=30, what="kondisi"):
        end = time.time() + timeout
        while time.time() < end:
            try:
                if cond():
                    return
            except Exception:
                pass
            time.sleep(random.uniform(0.5, 1.0))
        raise FormError(f"timeout menunggu {what}")

    def cek_sesi(self):
        txt = self.body()
        if "Masuk ke akun Anda" in txt or "sso.bps.go.id" in self.page.url:
            raise SesiHabis("layar login tampil")
        if re.search(r"captcha|bukan bot|human visitor|support id is", txt, re.I):
            raise SystemExit("!! Muncul verifikasi/CAPTCHA. Bot berhenti; selesaikan manual lalu jalankan lagi.")

    def snapshot_nilai(self):
        try:
            return {f["id"]: (f["value"] or "").strip() for f in self.page.evaluate(JS_FIELDS) if f.get("id")}
        except Exception:
            return {}

    def buka(self, aid, boleh_revoke=False):
        """Buka /edit. Terkunci + boleh_revoke -> Revoke lalu buka ulang. True bila revoke dilakukan.
        Bila isi berubah setelah revoke (seharusnya hanya status) -> Berhenti."""
        url = EDIT_URL.format(sid=SURVEY_ID, aid=aid)
        self._muat(url)
        if self.bisa_edit():
            jeda("buka")
            return False
        if not boleh_revoke:
            raise PerluRevoke("assignment terkunci (tanpa tombol Kirim); perlu revoke")
        sebelum = self.snapshot_nilai()
        self.revoke()
        self._muat(url)
        if not self.bisa_edit():
            raise PerluRevoke("masih terkunci setelah revoke")
        sesudah = self.snapshot_nilai()
        berubah = [(k, v, sesudah.get(k, v)) for k, v in sebelum.items() if v and sesudah.get(k, v) != v]
        if berubah:
            raise Berhenti("isi berubah setelah revoke: " + "; ".join(f"{k}: '{a}' -> '{b}'" for k, a, b in berubah[:5]))
        jeda("buka")
        return True

    def _muat(self, url):
        # Halaman login FASIH tampil di alamat /edit yang sama. Saat token sesi diperbarui, layar login bisa
        # muncul SEKILAS lalu form tampil sendiri -> jangan langsung dianggap sesi habis: beri waktu, muat
        # ulang sekali, baru berhenti bila halaman login tetap ada.
        login = lambda: self.page.get_by_text(re.compile(r"Masuk ke akun Anda|Lanjutkan dengan SSO")).count() > 0
        form = lambda: self.page.locator("#fasih-form-nav-next-button").count() > 0 and "Loading" not in self.body()
        for coba in range(2):
            self.page.goto(url, wait_until="domcontentloaded")
            self.wait(lambda: login() or form(), 60, "form termuat")
            if not login():
                break
            try:
                self.wait(form, 20, "form setelah layar login sekilas")
                break
            except FormError:
                if coba == 1 or "sso.bps.go.id" in self.page.url:
                    if self.sesi_aktif():
                        # sesi sehat, hanya assignment ini yang menampilkan layar login -> lewati, cek manual
                        raise FormError("halaman edit assignment ini menampilkan layar login padahal sesi aktif; cek manual")
                    raise SesiHabis("layar login tetap tampil")
                time.sleep(random.uniform(3, 6))
        self.cek_sesi()

    def sesi_aktif(self):
        """Buka daftar survei: tampil tanpa layar login = sesi masih aktif."""
        try:
            self.page.goto(FASIH + "/app/surveys", wait_until="domcontentloaded")
            time.sleep(random.uniform(5, 8))
            return (self.page.get_by_text(re.compile(r"Masuk ke akun Anda|Lanjutkan dengan SSO")).count() == 0
                    and "/surveys" in self.page.url)
        except Exception:
            return False
        time.sleep(random.uniform(1.2, 2.6))

    def revoke(self):
        jeda("navigasi")
        tombol = self.page.locator("button[class*='f:inline-flex']").filter(visible=True)
        if tombol.count() < 2:
            raise PerluRevoke("tombol revoke (mengambang kanan bawah) tidak ditemukan")
        tombol.last.click()
        dlg = self.page.locator("[role=dialog], [role=alertdialog]").filter(has_text="Revoke Assignment")
        try:
            self.wait(lambda: dlg.count() > 0, 10, "dialog 'Revoke Assignment?'")
        except FormError:
            self.page.keyboard.press("Escape")
            raise PerluRevoke("tombol mengambang tidak memunculkan dialog 'Revoke Assignment?'")
        time.sleep(random.uniform(0.8, 1.8))
        dlg.first.locator("button", has_text=re.compile(r"^\s*Konfirmasi\s*$")).first.click()
        self.wait(lambda: dlg.count() == 0, 20, "revoke selesai")
        time.sleep(random.uniform(1.5, 3.0))

    def bisa_edit(self):
        return self.page.locator("button", has_text=re.compile(r"^\s*Kirim\s*$")).filter(visible=True).count() > 0

    def ke_blok(self, nama):
        aside = self.page.locator("aside")
        if aside.count() == 0 or not aside.first.is_visible():
            self.page.locator("button", has_text="sidebar-toggle").first.click()
            self.wait(lambda: aside.count() > 0 and aside.first.is_visible(), 10, "sidebar terbuka")
        item = aside.first.get_by_text(nama, exact=True).first
        item.evaluate("e => (e.closest('a,button,[role=button],li,[tabindex]') || e).click()")
        jeda("navigasi")

    def peta_usaha(self, aid):
        """index usaha (index1) -> nama usaha, langsung dari data assignment di server
        (jawaban berkunci 'nama_usaha#<index1>'). {} bila tidak terbaca."""
        url = ASSIGNMENT_API.format(aid=aid)
        try:
            d = self.page.evaluate("async (u) => (await fetch(u, {credentials: 'same-origin'})).json()", url)
            isi = json.loads(d["data"]["data"]) if isinstance(d["data"].get("data"), str) else d["data"]["data"]
            peta = {}
            for a in isi.get("answers", []):
                m = re.match(r"nama_usaha(_edit)?#(\w+)$", a.get("dataKey", ""))
                if m and isinstance(a.get("answer"), str) and a["answer"].strip():
                    if m.group(1) or m.group(2) not in peta:  # nama_usaha_edit (hasil koreksi) diutamakan
                        peta[m.group(2)] = a["answer"].strip()
            return peta
        except Exception:
            return {}

    def buka_usaha(self, u, aid):
        """Buka kartu usaha di BLOK II untuk usaha sasaran.
        Utama: index usaha dari Excel (kunci) -> nama usaha di data server -> kartu bernama sama persis.
        Cadangan (data server tak terbaca): pencocokan nama Excel yang ketat."""
        peta = self.peta_usaha(aid)
        self.ke_blok(BLOK_USAHA)
        self.wait(lambda: len(self.page.evaluate(JS_CARDS)) > 0, 20, "daftar kartu usaha")
        cards = self.page.evaluate(JS_CARDS)
        if peta:
            idx = u["index1"]
            if idx not in peta:
                raise PerluManual(f"index usaha {idx} ('{u['nama_usaha']}') tidak ada di data assignment "
                                  f"(usaha di FASIH: {', '.join(f'{k}={v}' for k, v in sorted(peta.items()))})")
            # "<...>" di nama usaha dianggap tag HTML oleh halaman, jadi tidak tampil di kartu
            tampil = lambda s: norm(re.sub(r"<[^>]*>", " ", s))
            nama_srv = tampil(peta[idx])
            sama = [c for c in cards if tampil(c["nama"]) == nama_srv]
            if not sama:  # cadangan: kartu yang sangat mirip dengan nama DI SERVER (bukan nama Excel)
                mirip = [(difflib.SequenceMatcher(None, nama_srv, tampil(c["nama"])).ratio(), c) for c in cards]
                mirip = [x for x in sorted(mirip, key=lambda x: -x[0]) if x[0] >= 0.9]
                if len(mirip) == 1:
                    sama = [mirip[0][1]]
            if len(sama) == 1:
                card, cara = sama[0], f"index {idx} = '{peta[idx]}'"
            elif len(sama) > 1:
                # nama kembar: urutan kartu mengikuti urutan index usaha yang bernama sama
                urut = [k for k in sorted(peta, key=lambda k: (len(k), k)) if norm(peta[k]) == nama_srv]
                if len(urut) != len(sama):
                    raise PerluManual(f"nama kembar '{peta[idx]}' tidak bisa dipetakan ke kartu")
                card, cara = sama[urut.index(idx)], f"index {idx}, nama kembar ke-{urut.index(idx) + 1}"
            else:
                raise PerluManual(f"kartu bernama '{peta[idx]}' (index {idx}) tidak ada di BLOK II")
            self.page.locator("[data-nested-view]").nth(card["idx"]).evaluate("e => e.click()")
            self.wait(lambda: any(LABEL_TAHUNAN.search(f["label"]) or LABEL_BULANAN.search(f["label"])
                                  for f in self.page.evaluate(JS_FIELDS)), 25, "rincian 27.a / 31.a")
            jeda("navigasi")
            return card["nama"], cara
        target = norm(u["nama_usaha"])
        skor = [(difflib.SequenceMatcher(None, target, norm(c["nama"])).ratio(), c) for c in cards]
        skor.sort(key=lambda x: -x[0])
        kembar = [(sk, c) for sk, c in skor if skor and sk >= NAMA_MIN_SKOR and skor[0][0] - sk <= 0.05]
        pesaing = [sk for sk, _ in skor[1:] if sk >= 0.6]  # kartu lain yang namanya juga mirip
        if skor and skor[0][0] >= 0.95 and len(kembar) == 1:
            s, card = skor[0]
            cara = f"nama cocok {s:.2f}"
        elif skor and skor[0][0] >= NAMA_MIN_SKOR and len(kembar) == 1 and not pesaing:
            # nama tidak persis sama: hanya diterima bila tidak ada kartu lain yang juga mirip
            # (mis. "KEBUN BURHAN" vs "KEBUN PALA (BURHAN ALIM)" & "KEBUN KELAPA (BURHAN ALIM)" -> manual)
            s, card = skor[0]
            cara = f"nama mirip {s:.2f}, tanpa pesaing"
        elif len(kembar) > 1 and u["index1"].isdigit() and len(u["index1"]) <= 3 and \
                any(c["no"] == u["index1"] for _, c in kembar):
            # nama usaha kembar dalam satu assignment: index1 pendek (1/2/3) = nomor kartu
            s, card = next((sk, c) for sk, c in kembar if c["no"] == u["index1"])
            cara = f"nama kembar, kartu no {card['no']}"
        # catatan: aturan "kartu tunggal" dihapus - usaha bisa ada di bagian lain kuesioner
        # (kasus "Berkebun kelapa": satu-satunya kartu BLOK II adalah usaha lain)
        else:
            daftar = "; ".join(f"{c['no']}. {c['nama']} ({sk:.2f})" for sk, c in skor[:4])
            raise PerluManual(f"kartu usaha tidak bisa dipastikan untuk '{u['nama_usaha']}': {daftar}")
        self.page.locator("[data-nested-view]").nth(card["idx"]).evaluate("e => e.click()")
        self.wait(lambda: any(LABEL_TAHUNAN.search(f["label"]) or LABEL_BULANAN.search(f["label"])
                              for f in self.page.evaluate(JS_FIELDS)), 25, "rincian 27.a / 31.a")
        jeda("navigasi")
        return card["nama"], cara

    def field_pendapatan(self):
        """27.a (tahunan) bila ada; 31.a (bulanan) bila 27.a tidak ada atau kosong sementara 31.a terisi."""
        fields = self.page.evaluate(JS_FIELDS)
        th = [f for f in fields if LABEL_TAHUNAN.search(f["label"])]
        bl = [f for f in fields if LABEL_BULANAN.search(f["label"])]
        if len(th) > 1 or len(bl) > 1:
            raise FormError(f"rincian pendapatan ganda (27.a: {len(th)}, 31.a: {len(bl)})")
        if th and (th[0]["value"].strip() or not (bl and bl[0]["value"].strip())):
            return th[0], "r.27.a"
        if bl:
            return bl[0], "r.31.a"
        if th:
            return th[0], "r.27.a"
        raise FormError("rincian 27.a / 31.a tidak ditemukan")

    def _buka_catatan(self, field_id):
        ok = self.page.evaluate("""(fid) => {
          const inp = document.getElementById(fid); if (!inp) return false;
          inp.scrollIntoView({block: 'center'});
          const y = inp.getBoundingClientRect().top;
          const b = [...document.querySelectorAll('button[title=Catatan], button[aria-label=Catatan]')]
            .filter(b => b.offsetParent).map(b => ({b, dy: Math.abs(b.getBoundingClientRect().top - y)}))
            .sort((a, c) => a.dy - c.dy)[0];
          if (!b || b.dy > 60) return false;
          b.b.click(); return true;
        }""", field_id)
        if not ok:
            raise FormError("tombol Catatan untuk rincian pendapatan tidak ditemukan")
        dlg = self.page.locator("[role=dialog]").filter(has=self.page.locator("textarea[placeholder*='catatan' i]"))
        self.wait(lambda: dlg.count() > 0, 10, "dialog Catatan")
        return dlg

    def _tutup_catatan(self, dlg):
        time.sleep(random.uniform(0.5, 1.2))
        dlg.first.locator("button[aria-label=Dismiss]").first.click()
        self.wait(lambda: dlg.count() == 0, 10, "dialog Catatan tertutup")

    def catatan(self, field_id, tulis):
        """Buka Catatan rincian. Sudah ada '#0' -> 'SUDAH_ADA'. Selain itu: tulis '#0' (bila tulis) -> 'TERTULIS',
        atau 'BELUM_ADA' (dry-run, tidak menulis)."""
        dlg = self._buka_catatan(field_id)
        d = dlg.first
        jeda("baca")
        isi = d.inner_text()
        if ADA_KODE0.search(isi.replace("Catatan", " ")):
            self._tutup_catatan(dlg)
            return "SUDAH_ADA"
        if not tulis:
            self._tutup_catatan(dlg)
            return "BELUM_ADA"
        ta = d.locator("textarea").first
        ta.click()
        ta.fill(TEKS_REMARK)
        time.sleep(random.uniform(*JEDA["tempel"]))
        d.locator("button[title=Simpan]").first.click()
        self.wait(lambda: ADA_KODE0.search(d.inner_text()) is not None, 15, "catatan tersimpan")
        self._tutup_catatan(dlg)
        return "TERTULIS"


# ---------------------------------------------------------------- kirim (sama dengan bot NIK/KK)
def kirim(form, aid, log):
    page = form.page
    for _ in range(5):
        jeda("kirim")
        if page.locator("[role=dialog]").filter(has_text="Belum diisi").count() == 0:
            for _ in range(3):
                sisa = page.locator("[role=dialog], [role=alertdialog]").filter(visible=True)
                if sisa.count() == 0:
                    break
                tutup = sisa.first.locator("button", has_text=re.compile(r"^\s*(Batal|Close)\s*$")).filter(visible=True)
                if tutup.count() == 0:
                    tutup = sisa.first.locator("button[aria-label=Dismiss]")
                tutup.first.click(timeout=5000)
                jeda("baca")
            page.locator("button", has_text=re.compile(r"^\s*Kirim\s*$")).filter(visible=True).first.click()
        form.wait(lambda: page.locator("[role=dialog]").filter(has_text="Belum diisi").count() > 0, 20, "dialog Kirim")
        time.sleep(random.uniform(1.2, 2.6))
        dlg = page.locator("[role=dialog]").filter(has_text="Belum diisi").first
        teks = dlg.inner_text()
        m = re.search(r"GALAT\s*\n\s*(\d+)", teks, re.I)
        if (int(m.group(1)) if m else 0) == 0:
            dlg.locator("button", has_text=re.compile(r"^\s*Kirim\s*$")).first.click()
            konf = page.locator("[role=dialog], [role=alertdialog]").filter(has_text="Konfirmasi Kirim")
            # Bukti sukses: (a) teks "Berhasil Disubmit" (dicek lewat elemen, bukan membaca seluruh halaman
            # besar yang sering timeout), atau (b) respons server 2xx atas permintaan kirim (POST/PUT).
            respons = []
            def catat(r):
                try:
                    if r.request.method in ("POST", "PUT", "PATCH") and "bps.go.id" in r.url:
                        respons.append((r.request.method, r.url.split("?")[0], r.status))
                except Exception:
                    pass
            page.on("response", catat)
            sukses_teks = lambda: page.get_by_text("Berhasil Disubmit").count() > 0
            # Kirim FASIH = 3 langkah: .../s3/edit/presign-url (izin unggah) -> PUT data -> .../s3/edit (simpan).
            # Hanya langkah TERAKHIR yang membuktikan tersimpan; presign 2xx saja BUKAN bukti.
            sukses_server = lambda: any(200 <= s < 300 and re.search(r"/s3/edit/?$", u, re.I) for _, u, s in respons)
            try:
                form.wait(lambda: konf.count() > 0 or sukses_teks(), 20, "dialog Konfirmasi Kirim")
                if konf.count() > 0:
                    time.sleep(random.uniform(0.8, 1.8))
                    konf.first.locator("button", has_text=re.compile(r"^\s*Konfirmasi\s*$")).first.click()
                # server bisa lambat (terutama saat beberapa bot paralel): tunggu lebih lama
                form.wait(lambda: sukses_teks() or sukses_server(), 90, "'Berhasil Disubmit'")
                bukti = "dialog sukses" if sukses_teks() else "respons server 2xx"
            except FormError:
                shot = LOG_DIR / f"kirim_{aid}.png"
                page.screenshot(path=str(shot))
                jejak = "; ".join(f"{m} {u.split('bps.go.id')[-1]} {s}" for m, u, s in respons[-5:]) or "tanpa POST/PUT"
                log_kirim(log, aid, "ERROR", f"tidak muncul 'Berhasil Disubmit'; jaringan: {jejak}; screenshot {shot.name}")
                return "ERROR"
            finally:
                page.remove_listener("response", catat)
            m = re.search(r"PERINGATAN\s*\n\s*(\d+)", teks, re.I)
            log_kirim(log, aid, "TERKIRIM", f"peringatan diabaikan: {m.group(1) if m else 0}; bukti: {bukti}")
            return "TERKIRIM"
        dlg.get_by_text("Perlu diperbaiki").first.click()
        time.sleep(random.uniform(1.2, 2.6))
        daftar = page.locator("[role=dialog]").first.inner_text()
        g = next((g for g in GALAT_DIKENAL if g["teks"] in daftar), None)
        if g is None:
            shot = LOG_DIR / f"galat_{aid}.png"
            page.screenshot(path=str(shot))
            ringkas = re.sub(r"\s+", " ", daftar)[:300]
            log_kirim(log, aid, "GALAT_TIDAK_DIKENAL", f"{ringkas}; screenshot {shot.name}")
            page.locator("[role=dialog] button[aria-label=Dismiss]").first.click()
            return "ERROR"
        item = page.locator("[role=dialog]").first.locator("div", has_text=g["teks"]).filter(
            has=page.locator("button[title=Lihat]")).last
        item.locator("button[title=Lihat]").first.click()
        form.wait(lambda: page.locator("[role=dialog]").count() == 0, 10, "pindah ke pertanyaan galat")
        time.sleep(random.uniform(1.0, 2.0))
        g["perbaiki"](form)
        log_kirim(log, aid, "GALAT_DIPERBAIKI", g["teks"])
    log_kirim(log, aid, "ERROR", "galat masih ada setelah 5x perbaikan")
    return "ERROR"


def _switch_anomali(page):
    for sel in ("#cek_anomali_button input[type=checkbox]", "#cek_anomali_button [role=switch]",
                "[id*='cek_anomali'] input[type=checkbox]", "[id*='cek_anomali'] [role=switch]"):
        loc = page.locator(sel).first
        if loc.count() > 0:
            return loc
    return None


def _switch_aktif(loc):
    try:
        if loc.get_attribute("type") == "checkbox":
            return loc.is_checked()
    except Exception:
        pass
    return (loc.get_attribute("aria-checked") or "").lower() == "true"


def perbaiki_tampilkan_anomali(form):
    form.wait(lambda: _switch_anomali(form.page) is not None, 20, "switch 'Tampilkan Anomali'")
    sw = _switch_anomali(form.page)
    try:  # menggulir hanya kosmetik; jangan gagalkan perbaikan bila halaman sedang berat
        sw.scroll_into_view_if_needed(timeout=8000)
    except Exception:
        pass
    time.sleep(random.uniform(0.4, 1.0))
    if not _switch_aktif(sw):
        sw.evaluate("e => (e.closest('[role=switch]') || e).click()")
        form.wait(lambda: _switch_aktif(_switch_anomali(form.page) or sw), 10, "switch anomali aktif")
    jeda("field", form.cepat)


GALAT_DIKENAL = [{"teks": "Tampilkan Anomali Usaha dan Keluarga", "perbaiki": perbaiki_tampilkan_anomali}]


# ---------------------------------------------------------------- proses
def log_kirim(log, aid, status, ket):
    log({"waktu": datetime.now().strftime("%H:%M:%S"), "assignment_id": aid, "kunci": "", "nama_usaha": "",
         "rincian": "KIRIM", "status": status, "keterangan": ket})


def proses_assignment(form, aid, usaha, live, log):
    """Kembalikan (hasil_assignment, {kunci: status})."""
    direvoke = form.buka(aid, boleh_revoke=live)
    if direvoke:
        log_kirim(log, aid, "REVOKE", "assignment terkunci -> Revoke (Konfirmasi) -> /edit dibuka ulang")
    status = {}
    for u in usaha:
        rec = {"waktu": datetime.now().strftime("%H:%M:%S"), "assignment_id": aid, "kunci": u["kunci"],
               "nama_usaha": u["nama_usaha"], "rincian": "", "status": "", "keterangan": ""}
        try:
            nama_form, cara = form.buka_usaha(u, aid)
            f, rincian = form.field_pendapatan()
            rec["rincian"] = rincian
            rec["keterangan"] = f"kartu '{nama_form}' ({cara})"
            if f["disabled"]:
                raise FormError(f"{rincian} tidak bisa diedit (disabled)")
            hasil = form.catatan(f["id"], tulis=live)
            rec["status"] = {"BELUM_ADA": "DRYRUN_OK"}.get(hasil, hasil)
            jeda("field", form.cepat)
        except PerluManual as e:
            rec["status"], rec["keterangan"] = "PERLU_MANUAL", str(e)[:300]
        except Exception as e:
            if isinstance(e, (SystemExit, Berhenti, SesiHabis)):
                raise
            rec["status"] = "ERROR"
            rec["keterangan"] = (rec["keterangan"] + "; " if rec["keterangan"] else "") + str(e).splitlines()[0][:200]
        log(rec)
        status[u["kunci"]] = rec["status"]
    if not live:
        return ("ERROR" if "ERROR" in status.values() else "OK"), status
    # Kirim bila ada catatan baru, bila tadi di-revoke (status harus dipulihkan), atau bila #0 sudah ada
    # di assignment yang masih bisa diedit (catatan dari run sebelumnya mungkin tersimpan tapi belum terkirim).
    perlu_kirim = "TERTULIS" in status.values() or direvoke or "SUDAH_ADA" in status.values()
    if not perlu_kirim:
        log_kirim(log, aid, "DILEWATI", "tidak ada catatan baru, tidak dikirim")
        return ("ERROR" if "ERROR" in status.values() else "SKIP"), status
    hasil = kirim(form, aid, log)
    if hasil == "TERKIRIM":
        status = {k: ("SELESAI" if v == "TERTULIS" else v) for k, v in status.items()}
    else:  # belum terkirim: jangan dianggap final, diproses lagi pada run berikutnya
        status = {k: ("BELUM_TERKIRIM" if v in ("TERTULIS", "SUDAH_ADA") else v) for k, v in status.items()}
    return hasil, status


KUNCI_DIR = LOG_DIR / "sedang_dikerjakan"
KUNCI_KEDALUWARSA = 30 * 60  # detik; kunci lebih tua dari ini dianggap sisa bot yang mati


def kunci_ambil(aid, port):
    """Tandai assignment sedang dikerjakan bot ini. False bila sedang dikerjakan bot lain."""
    KUNCI_DIR.mkdir(parents=True, exist_ok=True)
    f = KUNCI_DIR / f"{aid}.lock"
    if f.exists() and time.time() - f.stat().st_mtime > KUNCI_KEDALUWARSA:
        f.unlink(missing_ok=True)
    try:
        with open(f, "x") as fh:
            fh.write(str(port))
        return True
    except FileExistsError:
        return False


def kunci_lepas(aid):
    (KUNCI_DIR / f"{aid}.lock").unlink(missing_ok=True)


def _dicoba_dir():
    return LOG_DIR / "sudah_dicoba" / datetime.now().strftime("%Y-%m-%d")


def sudah_dicoba(aid):
    """Sudah dicoba bot mana pun HARI INI (berhasil atau gagal) -> jangan dibuka lagi bot lain hari ini."""
    return (_dicoba_dir() / aid).exists()


def tandai_dicoba(aid):
    _dicoba_dir().mkdir(parents=True, exist_ok=True)
    (_dicoba_dir() / aid).touch()


TUNGGU_LOGIN = 60 * 60  # detik; menunggu login ulang manual paling lama segini, lalu berhenti


def cek_sesi_api(ctx, aid):
    """'AKTIF' / 'BOT' (halaman Bot Detected) / 'HABIS'. Lewat API dengan cookie browser: tidak memindahkan tab
    mana pun, jadi tidak mengganggu jendela tempat Anda sedang login."""
    try:
        r = ctx.request.get(FASIH + ASSIGNMENT_API.format(aid=aid), max_redirects=0, timeout=20000)
        t = r.text()
    except Exception:
        return "HABIS"
    if "Bot Detected" in t[:300]:
        return "BOT"
    try:
        return "AKTIF" if r.ok and json.loads(t).get("success") else "HABIS"
    except ValueError:
        return "HABIS"


def tunggu_login_manual(ctx, aid, port):
    """Bot TIDAK login sendiri: menunggu Anda login ulang di jendela Chrome port ini, lalu menutup tab lama."""
    print(f"!! Sesi login habis di port {port}. Silakan login ulang MANUAL di jendela Chrome port ini "
          f"(Lanjutkan dengan SSO -> Log In). Bot menunggu maks. {TUNGGU_LOGIN // 60} menit lalu melanjutkan sendiri.",
          flush=True)
    batas, lapor_bot = time.time() + TUNGGU_LOGIN, False
    while time.time() < batas:
        time.sleep(random.uniform(20, 40))
        st = cek_sesi_api(ctx, aid)
        if st == "AKTIF":
            break
        if st == "BOT" and not lapor_bot:
            print("!! Server menampilkan halaman 'Bot Detected'. Sebaiknya hentikan bot (Ctrl+C) & lanjutkan nanti.",
                  flush=True)
            lapor_bot = True
    else:
        raise SystemExit(f"!! Sesi tidak aktif lagi dalam {TUNGGU_LOGIN // 60} menit. Bot berhenti; jalankan lagi setelah login.")
    # tab lama (termasuk halaman login) ditutup supaya Chrome tidak berat; satu tab disisakan agar jendela tetap ada
    halaman = [x for x in ctx.pages if not x.is_closed()]
    for pg in halaman[:-1]:
        tutup_tab(pg)
    print(f"== Sesi aktif lagi; {max(len(halaman) - 1, 0)} tab lama ditutup, melanjutkan.", flush=True)
    time.sleep(random.uniform(5, 10))


def tutup_tab(pg):
    try:
        if pg.is_closed():
            return
        if len([x for x in pg.context.pages if not x.is_closed()]) <= 1:
            pg.goto("about:blank")
        else:
            pg.close(run_before_unload=False)
    except Exception:
        pass


def tampilkan_status(groups):
    done = baca_rekap()
    semua = [u["kunci"] for us in groups.values() for u in us]
    sel = sum(1 for k in semua if done.get(k) == "SELESAI")
    ada = sum(1 for k in semua if done.get(k) == "SUDAH_ADA")
    sisa = [k for k in semua if k not in done]
    print(f"Sasaran {len(semua)} usaha | SELESAI (ditulis & terkirim) {sel} | sudah ada #0 {ada} | sisa {len(sisa)}")
    LOG_DIR.mkdir(exist_ok=True)
    out = LOG_DIR / "sisa_belum_selesai.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["kunci", "nama_usaha", "kecamatan", "desa", "link_fasih"])
        for aid, us in groups.items():
            for u in us:
                if u["kunci"] in sisa:
                    w.writerow([u["kunci"], u["nama_usaha"], u["kecamatan"], u["desa"], EDIT_URL.format(sid=SURVEY_ID, aid=aid)])
    print(f"Daftar sisa: {out}")


MAKS_TAB_ERROR = 3


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--live", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--status", action="store_true", help="ringkasan progres dari log, tidak membuka browser")
    ap.add_argument("--file")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--ids")
    ap.add_argument("--port", type=int, default=9340)
    ap.add_argument("--shard", help="bagian K/N dari daftar, mis. 2/3 untuk bot ke-2 dari 3")
    ap.add_argument("--cepat", action="store_true")
    ap.add_argument("--ulang", action="store_true", help="jangan lewati yang sudah SELESAI/SUDAH_ADA di log")
    ap.add_argument("--mundur", action="store_true", help="kerjakan dari akhir daftar (bot kedua tanpa --shard)")
    args = ap.parse_args()
    live = args.live and not args.dry_run

    groups = load_targets(cari_file(args.file))
    if args.status:
        tampilkan_status(groups)
        return
    ids = list(groups)
    if args.ids:
        want = {s.strip() for s in args.ids.split(",")}
        ids = [i for i in ids if i in want]
    if args.shard:
        k, n = (int(x) for x in args.shard.split("/"))
        ids = ids[k - 1::n]
        print(f"Shard {k}/{n}: {len(ids)} assignment")
    if not args.ulang:
        done = baca_rekap()
        lewat = 0
        for aid in list(ids):
            sisa = [u for u in groups[aid] if u["kunci"] not in done]
            lewat += len(groups[aid]) - len(sisa)
            groups[aid] = sisa
        ids = [a for a in ids if groups[a]]
        if lewat:
            print(f"Dilewati {lewat} usaha yang sudah SELESAI / sudah ada #0 di log")
    lewati = baca_lewati()
    if lewati:
        n_lw = sum(1 for a in ids if a in lewati)
        ids = [a for a in ids if a not in lewati]
        print(f"Dilewati {n_lw} assignment dari {LEWATI_FILE.name} (kerjakan manual)")
    if args.mundur:
        ids.reverse()
    if args.limit:
        ids = ids[:args.limit]

    LOG_DIR.mkdir(exist_ok=True)
    mode = "live" if live else "dryrun"
    log_path = LOG_DIR / f"hasil_{mode}_{datetime.now():%Y%m%d_%H%M%S}_p{args.port}.csv"
    rekap = LOG_DIR / f"rekap_p{args.port}.csv"
    cols = ["waktu", "assignment_id", "kunci", "nama_usaha", "rincian", "status", "keterangan"]
    fh = open(log_path, "w", newline="", encoding="utf-8-sig")
    w = csv.DictWriter(fh, fieldnames=cols, restval="")
    w.writeheader()
    fr = wr = None
    if live:
        baru = not rekap.exists()
        fr = open(rekap, "a", newline="", encoding="utf-8-sig")
        wr = csv.DictWriter(fr, fieldnames=["tanggal"] + cols, restval="")
        if baru:
            wr.writeheader()

    def log(rec):
        w.writerow(rec)
        fh.flush()
        print(f"   [{rec['status']}] {rec.get('rincian', '')} {rec.get('nama_usaha', '')[:40]} {rec.get('keterangan', '')}")

    def rekap_tulis(aid, status):
        if not wr:
            return
        for u in groups[aid]:
            st = status.get(u["kunci"], "")
            wr.writerow({"tanggal": datetime.now().strftime("%Y-%m-%d"), "waktu": datetime.now().strftime("%H:%M:%S"),
                         "assignment_id": aid, "kunci": u["kunci"], "nama_usaha": u["nama_usaha"], "status": st})
        fr.flush()

    n_usaha = sum(len(groups[a]) for a in ids)
    print(f"Mode: {mode.upper()} | {len(ids)} assignment, {n_usaha} usaha | log: {log_path}")
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
        except Exception:
            raise SystemExit(f"!! Chrome port {args.port} belum terbuka. Jalankan buka_chrome.bat {args.port} lalu login SSO.")
        ctx = browser.contexts[0]
        form = Form(None, args.cepat)
        tab_lama, tab_error, error_beruntun = None, [], 0
        istirahat_berikut = random.randint(*JEDA["istirahat_tiap"])
        sesi_ulang = 0
        for n, aid in enumerate(ids, 1):
            # bot lain (port lain) bisa sudah/sedang mengerjakan assignment ini
            if not args.ulang:
                selesai_lain = baca_rekap()
                groups[aid] = [u for u in groups[aid] if u["kunci"] not in selesai_lain]
                if not groups[aid]:
                    continue
            if live and not args.ulang and sudah_dicoba(aid):
                print(f"\n[{n}/{len(ids)}] {aid} sudah dicoba hari ini (bot lain/run sebelumnya), dilewati")
                continue
            if live and not kunci_ambil(aid, args.port):
                print(f"\n[{n}/{len(ids)}] {aid} sedang dikerjakan bot lain, dilewati")
                continue
            print(f"\n[{n}/{len(ids)}] {aid} ({len(groups[aid])} usaha)")
            page = ctx.new_page()
            form.page = page
            if tab_lama is not None:
                tutup_tab(tab_lama)
                tab_lama = None
            status = {}
            try:
                hasil, status = proses_assignment(form, aid, groups[aid], live, log)
            except PerluRevoke as e:
                hasil = "REVOKE"
                log_kirim(log, aid, "PERLU_REVOKE", str(e))
            except Berhenti as e:
                log_kirim(log, aid, "BERHENTI", str(e))
                page.screenshot(path=str(LOG_DIR / f"berhenti_{aid}.png"))
                print("!! Run dihentikan:", e)
                break
            except SesiHabis:
                # assignment ini belum dikerjakan tuntas: kunci dilepas, TIDAK ditandai sudah dicoba,
                # lalu diulang setelah Anda login lagi (paling banyak 3x berturut-turut di assignment yang sama)
                kunci_lepas(aid)
                tab_error, tab_lama = [], None
                sesi_ulang += 1
                if sesi_ulang > 3:
                    raise SystemExit(f"!! Sesi putus 3x berturut-turut di {aid}. Bot berhenti; cek manual.")
                tunggu_login_manual(ctx, aid, args.port)
                ids.insert(n, aid)
                continue
            except Exception as e:
                if isinstance(e, SystemExit):
                    raise
                hasil = "ERROR"
                log_kirim(log, aid, "ERROR", str(e).splitlines()[0][:200])
            sesi_ulang = 0
            rekap_tulis(aid, status)
            if live:
                tandai_dicoba(aid)  # berhasil maupun gagal: bot lain tidak membukanya lagi hari ini
            kunci_lepas(aid)
            if hasil == "ERROR":
                try:
                    page.screenshot(path=str(LOG_DIR / f"error_{aid}.png"))
                except Exception:
                    pass
                tab_error.append(page)
                while len(tab_error) > MAKS_TAB_ERROR:
                    tutup_tab(tab_error.pop(0))
                error_beruntun += 1
                if error_beruntun >= 3:
                    print("!! 3 error berturut-turut, bot berhenti. Cek log & screenshot.")
                    break
            else:
                error_beruntun = 0
                tab_lama = page
            if n < len(ids):
                if n >= istirahat_berikut:
                    print("   ... istirahat sejenak")
                    jeda("istirahat", args.cepat)
                    istirahat_berikut = n + random.randint(*JEDA["istirahat_tiap"])
                else:
                    jeda("antar", args.cepat)
        if tab_lama is not None:
            tutup_tab(tab_lama)
        # koneksi CDP tidak ditutup dengan browser.close() agar Chrome tetap terbuka
    fh.close()
    if fr:
        fr.close()
    print(f"\nSelesai. Log: {log_path}" + (f" | Rekap: {rekap}" if live else ""))
    if live:
        tampilkan_status(load_targets(cari_file(args.file)))


if __name__ == "__main__":
    main()
