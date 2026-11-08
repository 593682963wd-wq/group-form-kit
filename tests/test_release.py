"""作者：Wangyuan；日期：2026-11-08。职责：通过合成输入复验已确认缺陷。"""
import unittest,tempfile,zipfile,re
from pathlib import Path
from openpyxl import Workbook
from collection import read_xlsx
C={k:k for k in ['id','group','origin','destination','weekdays','start','end','end_day_offset']}
class ReleaseTests(unittest.TestCase):
 def test_missing_dimension_allowed(self) -> None:
  with tempfile.TemporaryDirectory() as d:
   p=Path(d)/'source.xlsx';w=Workbook();w.active.append(list(C));w.active.append(['a','g','A','B','1','10:00','11:00',0]);w.save(p);w.close()
   with zipfile.ZipFile(p) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
   parts['xl/worksheets/sheet1.xml']=re.sub(rb'<dimension[^>]*/>',b'',parts['xl/worksheets/sheet1.xml'])
   with zipfile.ZipFile(p,'w') as z:
    for n,v in parts.items():z.writestr(n,v)
   self.assertEqual(read_xlsx(p,'Sheet',C)[0]['id'],'a')
 def test_uncached_formula_row_not_silently_skipped(self) -> None:
  with tempfile.TemporaryDirectory() as d:
   p=Path(d)/'source.xlsx';w=Workbook();w.active.append(list(C));w.active.append(['=1']*8);w.save(p);w.close()
   with self.assertRaisesRegex(ValueError,'无缓存公式'):read_xlsx(p,'Sheet',C)
