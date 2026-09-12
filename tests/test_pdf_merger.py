import tempfile
import unittest
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pitcsuite.tools.pdf_merger import merge, ordered_pdfs, Cancelled

class MergeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.a = self.root / "a"
        self.b = self.root / "b"
        self.a.mkdir()
        self.b.mkdir()
        self.output = self.root / "out.pdf"
    def pdf(self, folder, name, width, pages=1):
        with PdfWriter() as writer:
            for _ in range(pages):
                writer.add_blank_page(width=width, height=200)
            writer.write(folder / name)
    def test_single_numeric_order_and_repeat(self):
        for name, width in [("10.pdf", 110), ("2 (1).pdf", 121), ("2.pdf", 120), ("1.pdf", 101)]:
            self.pdf(self.a, name, width)
        output = self.a / "Merged_a.pdf"
        for _ in range(2):
            merge([self.a], output)
            self.assertEqual([float(p.mediabox.width) for p in PdfReader(output).pages], [101,120,121,110])
    def test_interleave_duplicates_and_uneven_folders(self):
        for folder, name, width in [(self.a,"1.pdf",101),(self.a,"1 (1).pdf",102),(self.b,"1.pdf",103),(self.b,"2.pdf",104),(self.b,"10.pdf",105)]:
            self.pdf(folder,name,width)
        merge([self.a,self.b], self.output, interleave=True)
        self.assertEqual([float(p.mediabox.width) for p in PdfReader(self.output).pages], [101,103,102,104,105])
    def test_padding_both_modes(self):
        self.pdf(self.a,"1.pdf",100,3)
        self.pdf(self.a,"2.pdf",110,2)
        for interleave in [False,True]:
            for per_side, count in [(0,5),(1,6),(2,8),(4,16)]:
                merge([self.a],self.output,interleave=interleave,pages_per_side=per_side)
                self.assertEqual(len(PdfReader(self.output).pages),count)
    def test_failure_preserves_output(self):
        self.pdf(self.a,"1.pdf",100)
        (self.a/"2.pdf").write_bytes(b"broken")
        self.output.write_bytes(b"previous")
        with self.assertRaises(ValueError):
            merge([self.a],self.output)
        self.assertEqual(self.output.read_bytes(),b"previous")
    def test_cancellation_preserves_output(self):
        self.pdf(self.a,"1.pdf",100,3)
        self.output.write_bytes(b"previous")
        with self.assertRaises(Cancelled):
            merge([self.a],self.output,cancelled=lambda:True)
        self.assertEqual(self.output.read_bytes(),b"previous")
    def test_stamp_uses_filename(self):
        self.pdf(self.a, "12.pdf", 300)
        merge([self.a], self.output, stamp=True)
        self.assertIn("Sr. No. 12", PdfReader(self.output).pages[0].extract_text())

    def test_empty_and_duplicate_folders(self):
        with self.assertRaises(ValueError):
            merge([self.a],self.output)
        with self.assertRaises(ValueError):
            ordered_pdfs([self.a,self.a],self.output)

if __name__ == "__main__":
    unittest.main()
