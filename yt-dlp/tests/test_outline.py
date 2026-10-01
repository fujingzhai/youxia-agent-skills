"""离线回归：标题深度、字体强调、脚注注入及最终文件校验。"""
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts' if (ROOT / 'scripts/render_docx.py').exists() else ROOT))
import render_docx as renderer

class OutlineTests(unittest.TestCase):
    def test_heading_levels_survive_save_and_footnotes(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / 'result.docx'
            md = '\n\n'.join('#' * i + ' 标题' + str(i) for i in range(1, 7))
            md += '\n\n**正文强调句**\n\n正文^[真实脚注]\n\n```\n## 代码不是标题\n```'
            st = renderer.render_markdown(md, out)
            self.assertEqual(st['headings'], 6)
            self.assertEqual(st['footnotes'], 1)
            renderer.verify_outline(out, [('标题' + str(i), i) for i in range(1, 7)])
            with ZipFile(out) as z:
                self.assertIn('真实脚注', z.read('word/footnotes.xml').decode())
            doc = renderer.docx.Document(out)
            p = next(p for p in doc.paragraphs if p.text == '标题6')
            p._p.get_or_add_pPr().remove(p._p.get_or_add_pPr().find(renderer.qn('w:outlineLvl')))
            doc.save(out)
            with self.assertRaises(ValueError):
                renderer.verify_outline(out, [('标题' + str(i), i) for i in range(1, 7)])

if __name__ == '__main__':
    unittest.main()
