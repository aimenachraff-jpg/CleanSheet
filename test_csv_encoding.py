import sys, tempfile, unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import clean_sheet as cs

ARABIC_LINE = "الاسم,الكمية,السعر\nأحمد,3,1200 دج\n"


def write_bytes(path, data):
    path.write_bytes(data)


class TestDecodeText(unittest.TestCase):
    def test_utf8_sig_arabic(self):
        self.assertEqual(cs.decode_text(ARABIC_LINE.encode("utf-8-sig")).splitlines()[0], "الاسم,الكمية,السعر")

    def test_plain_utf8_arabic(self):
        self.assertEqual(cs.decode_text(ARABIC_LINE.encode("utf-8")).splitlines()[0], "الاسم,الكمية,السعر")

    def test_windows_1256_arabic(self):
        # Arabic Windows-1256 file must not be misread as cp1252 (mojibake)
        text = cs.decode_text(ARABIC_LINE.encode("windows-1256"))
        self.assertEqual(text.splitlines()[0], "الاسم,الكمية,السعر")
        self.assertEqual(text.splitlines()[1].split(",")[0], "أحمد")

    def test_iso_8859_6_arabic_does_not_mojibake_to_latin(self):
        # iso-8859-6 and windows-1256 are ambiguous single-byte Arabic codecs;
        # windows-1256 is tried first (the common Windows case). Assert an
        # iso-8859-6 Arabic file is not decoded as Latin mojibake.
        text = cs.decode_text(ARABIC_LINE.encode("iso-8859-6"))
        self.assertTrue(cs._looks_arabic(text), "Arabic bytes must not decode to Latin mojibake")
        # first line should contain Arabic-script letters, not cp1252 latin garbage
        self.assertTrue(any(ch in "الص" for ch in text.splitlines()[0]))

    def test_utf16_bom(self):
        text = cs.decode_text(ARABIC_LINE.encode("utf-16"))
        self.assertEqual(text.splitlines()[0], "الاسم,الكمية,السعر")

    def test_cp1252_latin_stays_latin(self):
        latin = "Nom,Quantit,Prix\nCaf,2,1,50\n".encode("cp1252")
        text = cs.decode_text(latin)
        self.assertEqual(text.splitlines()[0], "Nom,Quantit,Prix")

    def test_garbage_bytes_handled_gracefully(self):
        # cp1252 last resort decodes any byte sequence: no crash, returns str
        text = cs.decode_text(b"\xff\xfe\x00\x00\xff\xfe\xff\xff")
        self.assertIsInstance(text, str)

    def test_tiny_ambiguous_csv_does_not_crash(self):
        # Sniffer can raise on tiny/unusual files; must fall back to comma
        p = Path(tempfile.mkdtemp()) / "tiny.csv"
        p.write_text("Name\nAli\nSara\n", encoding="utf-8")
        r = cs.run(p, out_dir=Path(tempfile.mkdtemp()) / "o")
        self.assertEqual(r["rows_out"], 2)


class TestCsvEncodingEndToEnd(unittest.TestCase):
    def test_windows_1256_arabic_csv_cleans(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "ventes.csv"
            src.write_bytes(ARABIC_LINE.encode("windows-1256"))
            r = cs.run(src, out_dir=d / "o", locale="dz")
            clean = list(__import__("openpyxl").load_workbook(r["xlsx"])["Clean"].values)
            self.assertEqual(clean[1][0], "أحمد")
            self.assertEqual(clean[1][1], 3)
            self.assertEqual(clean[1][2], 1200.0)


if __name__ == "__main__": unittest.main(verbosity=1)
