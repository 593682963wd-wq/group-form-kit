"""作者：Wangyuan；日期：2026-10-29；职责：仅以从零合成输入验证核心与异常路径。"""

import tempfile, unittest
from pathlib import Path
from collection import *
class CollectionTests(unittest.TestCase):
    def setUp(self):
        self.rows=[{'id':'a','group':'g','origin':'地点甲','destination':'地点乙','weekdays':'135','start':'23:00','end':'01:00','end_day_offset':1}]
        self.contract=package('demo',group_plan(self.rows),['t1','t2'],[{'name':'count','type':'integer','required':True},{'name':'duration','type':'duration','required':False}])
    def test_weekday_and_crossday_validation(self):
        self.assertEqual(weekdays('3,1,3,5'),'135'); self.assertEqual(duration('23:00','01:00',1),120)
        for bad in ['8','0','Mon','']:
            with self.assertRaises(ValueError): weekdays(bad)
        with self.assertRaises(ValueError): duration('23:00','01:00',0)
        with self.assertRaises(ValueError): clock('24:00')
    def test_directional_filter_and_week_split(self):
        self.assertEqual(group_plan(self.rows,origin='地点乙'),[])
        self.assertEqual([x['id'] for x in group_plan(self.rows,split_weekdays=True)],['g:1','g:3','g:5'])
    def test_dynamic_headers_and_duplicate_identity(self):
        keys=['id','group','origin','destination','weekdays','start','end','end_day_offset']; cols={k:k for k in keys}
        table=[['title'],list(reversed(keys)),[self.rows[0][k] for k in reversed(keys)]]
        parsed=parse_table(table,cols); self.assertEqual(parsed[0]['duration_minutes'],120)
        with self.assertRaises(ValueError): parse_table(table+[table[-1]],cols)
    def test_package_hash_and_contract_tampering(self):
        validate_package(self.contract)
        with self.assertRaises(ValueError): validate_package({**self.contract,'plan':'changed'})
    def test_submission_completion_and_revision(self):
        with tempfile.TemporaryDirectory() as tmp:
            store=Collection(Path(tmp)/'c.sqlite',self.contract); args=dict(team='t1',group_id='g',values={'count':2,'duration':'02:30'},station='地点甲',expected_version=0,fingerprint=self.contract['fingerprint'])
            self.assertEqual(store.submit(**args)['version'],1); self.assertTrue(store.submit(**args)['reused'])
            with self.assertRaises(ValueError): store.submit(**{**args,'values':{'count':3}})
            result=store.summary(); self.assertEqual(result['completion'][0]['done'],1); self.assertEqual(result['completion'][1]['missing'],['g']); self.assertEqual(result['rows'][0]['values']['duration'],150)
    def test_invalid_counts_station_and_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            store=Collection(Path(tmp)/'c.sqlite',self.contract); args=dict(team='t1',group_id='g',values={'count':True},station='地点甲',expected_version=0,fingerprint=self.contract['fingerprint'])
            with self.assertRaises(ValueError): store.submit(**args)
            with self.assertRaises(ValueError): store.submit(**{**args,'values':{'count':2},'station':'未知'})
            with self.assertRaises(ValueError): store.submit(**{**args,'values':{'count':2},'fingerprint':'changed'})
    def test_xlsx_template_roundtrip(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'form.xlsx'; write_template(self.contract,path); book=load_workbook(path)
            self.assertEqual(book['收集表'].max_row,3); self.assertEqual(book['合同']['B1'].value,self.contract['fingerprint']); self.assertEqual(len(book['收集表'].data_validations.dataValidation),1); book.close()

if __name__ == "__main__":
    unittest.main()
