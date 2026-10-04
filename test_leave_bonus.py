import importlib.util
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('tracker', Path(__file__).with_name('LeaveBonusTracker_V0_8.py'))
import sys
m = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = m
spec.loader.exec_module(m)

class TrackerTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / 'test.sqlite3'
        self.db = m.LeaveBonusDB(self.path)
    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()
    def test_conflict_preserves_both(self):
        a = self.db.save_leave(None, '2099-02-01', 8, 'A')
        b = self.db.save_leave(None, '2099-02-02', 4, 'B')
        with self.assertRaises(ValueError): self.db.save_leave(a, '2099-02-02', 2, 'changed')
        self.assertEqual(self.db.get_leave_by_date('2099-02-01')['id'], a)
        self.assertEqual(self.db.get_leave_by_date('2099-02-02')['id'], b)
        self.assertEqual(self.db.get_leave_by_date('2099-02-01')['hours'], 8)
    def test_write_failure_rolls_back(self):
        count = self.db.conn.execute('SELECT COUNT(*) FROM app_entities').fetchone()[0]
        with patch.object(self.db, '_save_values', side_effect=RuntimeError('injected')):
            with self.assertRaises(RuntimeError): self.db.save_leave(None, '2099-01-01', 8, '')
        self.assertEqual(count, self.db.conn.execute('SELECT COUNT(*) FROM app_entities').fetchone()[0])
    def test_validation(self):
        for hours in [-1, 9, float('nan'), float('inf')]:
            with self.assertRaises(ValueError): self.db.save_leave(None, '2099-01-01', hours, '')
        for year in [0, 10000]:
            with self.assertRaises(ValueError): self.db.ensure_default_year(year)
        with self.assertRaises(ValueError): m.safe_int('2026.5')
        with self.assertRaises(ValueError): self.db.save_salary(m.SalarySetting(None,2099,'2098-01-01',0,1,0,0,0,0))
        with self.assertRaises(ValueError): self.db.save_annual_setting(m.AnnualSetting(2099, total_leave_days=-1))
    def test_bonus_boundary(self):
        self.db.save_salary(m.SalarySetting(None,2099,'2099-01-01',330,28440,17130,4870,1800,2000))
        for i in range(1,11): self.db.save_leave(None, f'2099-01-{i:02}',8,'')
        result = self.db.calc_year(2099)
        self.assertEqual(result['no_leave_base_bonus'],16000)
        self.assertAlmostEqual(result['remaining_overtime'],160 * 50440 / 240)
        self.db.save_leave(None,'2099-01-11',4,'')
        self.assertEqual(self.db.calc_year(2099)['used_leave_subsidy'],300)
    def test_legacy_with_generic_and_restart(self):
        self.db.conn.execute('CREATE TABLE leave_records (id INTEGER PRIMARY KEY, leave_date TEXT, hours REAL, note TEXT)')
        self.db.conn.execute("INSERT INTO leave_records VALUES (1,'2098-01-01',4,'legacy')")
        self.db.conn.commit()
        self.db._migrate_old_fixed_tables_if_needed()
        self.assertEqual(self.db.get_leave_by_date('2098-01-01')['hours'],4)
        self.db._migrate_old_fixed_tables_if_needed()
        self.assertEqual(len(self.db.list_leave_records(2098)),1)
    def test_legacy_conflict_rolls_back(self):
        self.db.save_leave(None,'2098-01-02',8,'existing')
        self.db.conn.execute('CREATE TABLE leave_records (id INTEGER PRIMARY KEY, leave_date TEXT, hours REAL, note TEXT)')
        self.db.conn.executemany('INSERT INTO leave_records VALUES (?,?,?,?)',[(1,'2098-01-01',4,'new'),(2,'2098-01-02',4,'conflict')])
        self.db.conn.commit()
        with self.assertRaises(ValueError): self.db._migrate_old_fixed_tables_if_needed()
        self.assertIsNone(self.db.get_leave_by_date('2098-01-01'))
        self.assertEqual(self.db.get_leave_by_date('2098-01-02')['hours'],8)
    def test_upgrade_backup(self):
        self.db.conn.execute('PRAGMA user_version=0')
        self.db.conn.commit()
        self.db.close()
        self.db = m.LeaveBonusDB(self.path)
        backups = list((self.path.parent/'backups').glob('*.sqlite3'))
        self.assertEqual(len(backups),1)
        with sqlite3.connect(backups[0]) as conn:
            self.assertEqual(conn.execute('PRAGMA integrity_check').fetchone()[0],'ok')
        self.db.close()
        self.db = m.LeaveBonusDB(self.path)
        self.assertEqual(len(list((self.path.parent/'backups').glob('*.sqlite3'))),1)
    def test_csv(self):
        target = Path(self.tmp.name)/'summary.csv'
        self.db.export_summary_csv(target)
        self.assertTrue(target.read_bytes().startswith(b'\xef\xbb\xbf'))
        self.assertIn('獎金總計',target.read_text(encoding='utf-8-sig'))


    def claim(self, **changes):
        args=dict(record_id=None,year_month="2099-10",travel_amount=2500,overtime_amount=4800,submit_date="2099-10-05")
        args.update(changes)
        return self.db.save_expense(**args)
    def test_status_and_totals(self):
        eid=self.claim()
        self.assertEqual(self.db.list_expenses(2099)[0]["status"],"待入帳")
        self.claim(record_id=eid,travel_deposit_date="2099-10-10")
        result=self.db.calc_year(2099)
        self.assertEqual(result["expense_paid"],2500)
        self.assertEqual(result["expense_pending"],4800)
        self.assertEqual(self.db.list_expenses(2099)[0]["status"],"部分入帳")
        self.claim(record_id=eid,travel_deposit_date="2099-10-10",overtime_deposit_date="2099-10-12")
        self.assertEqual(self.db.list_expenses(2099)[0]["status"],"已入帳")
        self.claim(record_id=eid)
        self.assertEqual(self.db.calc_year(2099)["expense_paid"],0)
    def test_single_item_paid(self):
        self.claim(overtime_amount=0,travel_deposit_date="2099-10-10")
        self.assertEqual(self.db.list_expenses(2099)[0]["status"],"已入帳")
    def test_duplicate_and_same_month(self):
        self.claim()
        with self.assertRaises(ValueError): self.claim()
        self.claim(submit_date="2099-10-06")
        self.assertEqual(len(self.db.list_expenses(2099)),2)
    def test_invalid_expenses(self):
        for change in [dict(year_month="2099-13"),dict(travel_amount=-1),dict(travel_amount=float("nan")),
                       dict(travel_amount=0,overtime_amount=0),dict(travel_deposit_date="2099-10-01"),
                       dict(overtime_amount=0,overtime_deposit_date="2099-10-10")]:
            with self.assertRaises(ValueError): self.claim(**change)
        self.assertEqual(len(self.db.list_expenses(2099)),0)
    def test_cross_year_and_delete(self):
        eid=self.claim()
        self.claim(year_month="2098-12",submit_date="2099-01-03")
        self.assertEqual(self.db.calc_year(2098)["expense_claimed"],7300)
        self.assertEqual(self.db.calc_year(2099)["expense_claimed"],7300)
        self.db.delete_expense(eid)
        self.assertEqual(self.db.calc_year(2099)["expense_claimed"],0)
    def test_v071_upgrade(self):
        self.db.save_leave(None,"2099-03-01",4,"keep")
        self.db.conn.execute("PRAGMA user_version=701")
        self.db.conn.commit()
        self.db.close()
        self.db=m.LeaveBonusDB(self.path)
        self.assertEqual(self.db.get_leave_by_date("2099-03-01")["hours"],4)
        self.assertEqual(self.db.conn.execute("PRAGMA user_version").fetchone()[0],800)

if __name__ == '__main__': unittest.main(verbosity=2)
