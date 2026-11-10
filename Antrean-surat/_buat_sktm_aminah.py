import shutil
from docx import Document

BASE = r"C:\Users\Sulthonu Tora Aulia\OneDrive\Dokumen\Project Desa MargaAgung"
TPL = BASE + r"\Template-surat\surat keterangan tidak mampu.docx"
DST = BASE + r"\Antrean-surat\SR-20261110-001 - Surat Keterangan Tidak Mampu - UJI SITI AMINAH.docx"

TANGGAL = "10 November 2026"

nilai = [
    ("nama", "UJI SITI AMINAH"),
    ("nik ", "1801999999990002"),
    ("tpt tgl lahir", "Marga Agung, 02-02-1985 / 41 tahun"),
    ("agama", "Islam"),
    ("setatus perkawinan", "Kawin"),
    ("pekerjaan", "Ibu Rumah Tangga"),
    ("alamat", "RT 02 Dusun I Marga Agung Kec. Jati Agung"),
    ("mendapatkan", "Bantuan Sosial"),
]


def set_text(p, text):
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


shutil.copyfile(TPL, DST)
doc = Document(DST)

for p in doc.paragraphs:
    low = p.text.lower().strip()
    if low.startswith("marga agung"):
        set_text(p, "      Marga Agung, " + TANGGAL)
        print("OK: tanggal ->", TANGGAL)
        continue
    if ":" not in p.text:
        continue
    for key, val in nilai:
        if low.startswith(key):
            set_text(p, p.text.split(":", 1)[0] + ": " + val)
            print("OK:", key, "->", val)
            break

doc.save(DST)

doc = Document(DST)
print("---- HASIL ----")
for p in doc.paragraphs:
    if p.text.strip():
        print(repr(p.text))
print("SIMPAN:", DST)