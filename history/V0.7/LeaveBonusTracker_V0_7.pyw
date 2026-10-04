# -*- coding: utf-8 -*-
"""
休假獎金統計小本 V0.7
功能：
- 以行事曆點選單一日期，或鍵入日期，登記該日休假時數。
- 統計已休假日數、未休假/剩餘日數。
- 試算不休假獎金、年終獎金、考績獎金。
- 每年的薪水、薪點與加給可新增、更正。
- 資料庫改為通用型資料表：app_entities + app_values，後續新增欄位或資料類型較容易擴充。
- 休假登記存檔後不再跳出確認視窗，直接更新清單與統計。
- 行事曆已登記休假的日期以紅色底標示。
- 鍵入日期存檔前會先檢查是否已有同日紀錄，避免重複登記。
- 第二分頁行事曆一次顯示兩個月，方便跨月登記。
- 第三分頁薪水加給改為跨年度總表，年度參數移入同頁第二行，並可引入上一年度資料。

資料：預設使用同資料夾 leave_bonus_data.sqlite3 儲存。

參考計算邏輯：
- 1 日 = 8 小時。
- 參考表的「小數制」顯示：4 小時 = 0.4、8 小時 = 0.8；實際統計以小時計算更穩定。
- 不休假獎金 = 剩餘休假加班費 + 不休假補助 + 已休假補助。
- 剩餘休假加班費 = 剩餘休假時數 × (月支薪津+專業加給+專業加給二) / 30 / 8。
- 不休假補助 = 已休假時數滿 10 日時 16,000 元；未滿 10 日為 0。
- 已休假補助 = max(已休假日數 - 10, 0) × 600 元。
- 年終獎金 = (月支薪津+專業加給+專業加給二) × 1.5。
- 考績獎金 = (月支薪津+專業加給+專業加給二) × 1.0。
"""

from __future__ import annotations

import calendar
import csv
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "休假獎金統計小本"
APP_VERSION = "V0.7"
DB_FILE = "leave_bonus_data.sqlite3"
WORK_HOURS_PER_DAY = 8.0

ENTITY_ANNUAL = "annual_setting"
ENTITY_SALARY = "salary_setting"
ENTITY_LEAVE = "leave_record"


def app_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def today_iso() -> str:
    return date.today().isoformat()


def now_text() -> str:
    return datetime.now().isoformat(timespec="seconds")


def parse_iso_date(text: str) -> date:
    return datetime.strptime(str(text).strip(), "%Y-%m-%d").date()


def money(value: float | int | None) -> str:
    if value is None:
        value = 0
    return f"{round(float(value)):,}"


def number(value: float | int | None, digits: int = 2) -> str:
    if value is None:
        value = 0
    value = float(value)
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.{digits}f}".rstrip("0").rstrip(".")


def hours_to_text(hours: float | int | None) -> str:
    if hours is None:
        hours = 0
    hours = float(hours)
    sign = "-" if hours < 0 else ""
    hours = abs(hours)
    days = int(hours // WORK_HOURS_PER_DAY)
    remain = round(hours - days * WORK_HOURS_PER_DAY, 2)
    return f"{sign}{days} 日 {number(remain, 2)} 小時"


def leave_decimal(hours: float | int | None) -> float:
    """參考表小數制：1 小時 = 0.1，4 小時特例顯示 0.4。"""
    if hours is None:
        return 0.0
    h = float(hours)
    if abs(h / WORK_HOURS_PER_DAY - 0.5) < 1e-9:
        return 0.4
    return h * 0.1


def safe_float(text: str, default: float = 0.0) -> float:
    text = str(text).strip().replace(",", "")
    if text == "":
        return default
    return float(text)


def safe_int(text: str, default: int = 0) -> int:
    text = str(text).strip().replace(",", "")
    if text == "":
        return default
    return int(float(text))


def daterange(start: date, end: date) -> list[date]:
    if end < start:
        start, end = end, start
    days = (end - start).days
    return [start + timedelta(days=i) for i in range(days + 1)]


@dataclass
class AnnualSetting:
    year: int
    total_leave_days: float = 30.0
    no_leave_base_bonus: float = 16000.0
    leave_subsidy_per_day: float = 600.0
    year_end_months: float = 1.5
    performance_months: float = 1.0


@dataclass
class SalarySetting:
    id: int | None
    year: int
    effective_date: str
    pay_point: float
    salary: float
    professional_1: float
    professional_2: float
    hazard_allowance: float
    retention_allowance: float
    note: str = ""

    @property
    def base_monthly(self) -> float:
        return self.salary + self.professional_1 + self.professional_2

    @property
    def total_monthly(self) -> float:
        return self.base_monthly + self.hazard_allowance + self.retention_allowance


class LeaveBonusDB:
    """通用型資料庫存取層。

    app_entities：每一筆資料的主體，例如一筆休假、一筆薪水設定、一筆年度參數。
    app_values：每一筆資料的欄位值，例如 hours、salary、professional_1。
    """

    def __init__(self, path: Path):
        self.path = path
        self.conn = sqlite3.connect(str(path))
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.init_db()

    def init_db(self) -> None:
        cur = self.conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS app_entities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entity_type TEXT NOT NULL,
                record_key TEXT NOT NULL DEFAULT '',
                year INTEGER,
                record_date TEXT,
                sort_order INTEGER NOT NULL DEFAULT 0,
                note TEXT NOT NULL DEFAULT '',
                is_deleted INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS app_values (
                entity_id INTEGER NOT NULL,
                field_key TEXT NOT NULL,
                field_value TEXT NOT NULL DEFAULT '',
                value_type TEXT NOT NULL DEFAULT 'text',
                updated_at TEXT NOT NULL,
                PRIMARY KEY (entity_id, field_key),
                FOREIGN KEY (entity_id) REFERENCES app_entities(id) ON DELETE CASCADE
            )
            """
        )
        cur.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_app_entities_record_key
            ON app_entities(entity_type, record_key)
            WHERE record_key <> ''
            """
        )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_app_entities_type_year ON app_entities(entity_type, year, is_deleted)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_app_entities_date ON app_entities(record_date, is_deleted)")
        self.conn.commit()
        self._migrate_old_fixed_tables_if_needed()
        self.ensure_default_year(date.today().year)

    def _table_exists(self, name: str) -> bool:
        row = self.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (name,),
        ).fetchone()
        return row is not None

    def _generic_has_data(self) -> bool:
        row = self.conn.execute(
            "SELECT COUNT(*) AS c FROM app_entities WHERE entity_type IN (?, ?, ?)",
            (ENTITY_ANNUAL, ENTITY_SALARY, ENTITY_LEAVE),
        ).fetchone()
        return int(row["c"]) > 0

    def _migrate_old_fixed_tables_if_needed(self) -> None:
        """把 V0.1 固定資料表搬到 V0.2 通用資料表。

        只在通用表尚無主要資料時執行，避免重複匯入。
        """
        if self._generic_has_data():
            return
        migrated = False
        if self._table_exists("annual_settings"):
            rows = self.conn.execute("SELECT * FROM annual_settings").fetchall()
            for r in rows:
                self.save_annual_setting(
                    AnnualSetting(
                        year=int(r["year"]),
                        total_leave_days=float(r["total_leave_days"]),
                        no_leave_base_bonus=float(r["no_leave_base_bonus"]),
                        leave_subsidy_per_day=float(r["leave_subsidy_per_day"]),
                        year_end_months=float(r["year_end_months"]),
                        performance_months=float(r["performance_months"]),
                    )
                )
                migrated = True
        if self._table_exists("salary_settings"):
            rows = self.conn.execute("SELECT * FROM salary_settings ORDER BY id").fetchall()
            for r in rows:
                self.save_salary(
                    SalarySetting(
                        id=None,
                        year=int(r["year"]),
                        effective_date=r["effective_date"],
                        pay_point=float(r["pay_point"]),
                        salary=float(r["salary"]),
                        professional_1=float(r["professional_1"]),
                        professional_2=float(r["professional_2"]),
                        hazard_allowance=float(r["hazard_allowance"]),
                        retention_allowance=float(r["retention_allowance"]),
                        note=r["note"] or "",
                    )
                )
                migrated = True
        if self._table_exists("leave_records"):
            rows = self.conn.execute("SELECT * FROM leave_records ORDER BY id").fetchall()
            for r in rows:
                self.save_leave(None, r["leave_date"], float(r["hours"]), r["note"] or "")
                migrated = True
        if migrated:
            self.conn.commit()

    def _value_type(self, value: object) -> str:
        if isinstance(value, bool):
            return "integer"
        if isinstance(value, int):
            return "integer"
        if isinstance(value, float):
            return "real"
        return "text"

    def _save_values(self, entity_id: int, values: dict[str, object]) -> None:
        now = now_text()
        for key, value in values.items():
            self.conn.execute(
                """
                INSERT INTO app_values (entity_id, field_key, field_value, value_type, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(entity_id, field_key) DO UPDATE SET
                    field_value=excluded.field_value,
                    value_type=excluded.value_type,
                    updated_at=excluded.updated_at
                """,
                (entity_id, key, str(value), self._value_type(value), now),
            )

    def _values_for(self, entity_id: int) -> dict[str, str]:
        rows = self.conn.execute(
            "SELECT field_key, field_value FROM app_values WHERE entity_id=?",
            (entity_id,),
        ).fetchall()
        return {r["field_key"]: r["field_value"] for r in rows}

    def _upsert_entity(
        self,
        *,
        entity_type: str,
        record_key: str,
        year: int | None,
        record_date: str | None,
        note: str = "",
        entity_id: int | None = None,
    ) -> int:
        now = now_text()
        if entity_id is not None:
            if record_key:
                conflict = self.conn.execute(
                    """
                    SELECT id FROM app_entities
                    WHERE entity_type=? AND record_key=? AND id<>?
                    """,
                    (entity_type, record_key, entity_id),
                ).fetchone()
                if conflict is not None:
                    self.conn.execute("DELETE FROM app_entities WHERE id=?", (int(conflict["id"]),))
            self.conn.execute(
                """
                UPDATE app_entities SET
                    entity_type=?, record_key=?, year=?, record_date=?, note=?, is_deleted=0, updated_at=?
                WHERE id=?
                """,
                (entity_type, record_key, year, record_date, note, now, entity_id),
            )
            return int(entity_id)

        existing = None
        if record_key:
            existing = self.conn.execute(
                """
                SELECT id FROM app_entities
                WHERE entity_type=? AND record_key=?
                """,
                (entity_type, record_key),
            ).fetchone()
        if existing is not None:
            eid = int(existing["id"])
            self.conn.execute(
                """
                UPDATE app_entities SET
                    year=?, record_date=?, note=?, is_deleted=0, updated_at=?
                WHERE id=?
                """,
                (year, record_date, note, now, eid),
            )
            return eid

        cur = self.conn.execute(
            """
            INSERT INTO app_entities
            (entity_type, record_key, year, record_date, note, is_deleted, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 0, ?, ?)
            """,
            (entity_type, record_key, year, record_date, note, now, now),
        )
        return int(cur.lastrowid)

    def ensure_default_year(self, year: int) -> None:
        if year <= 0:
            year = date.today().year
        if self._annual_entity(year) is None:
            self.save_annual_setting(AnnualSetting(year=year))
        salary_rows = self.conn.execute(
            """
            SELECT id FROM app_entities
            WHERE entity_type=? AND year=? AND is_deleted=0
            LIMIT 1
            """,
            (ENTITY_SALARY, year),
        ).fetchall()
        if not salary_rows:
            self.save_salary(
                SalarySetting(
                    id=None,
                    year=year,
                    effective_date=f"{year}-01-01",
                    pay_point=0,
                    salary=0,
                    professional_1=0,
                    professional_2=0,
                    hazard_allowance=0,
                    retention_allowance=0,
                    note="請填入當年薪水與加給",
                )
            )
        self.conn.commit()

    def years(self) -> list[int]:
        rows = self.conn.execute(
            """
            SELECT DISTINCT year FROM app_entities
            WHERE is_deleted=0 AND year IS NOT NULL
            ORDER BY year DESC
            """
        ).fetchall()
        years = [int(r["year"]) for r in rows if r["year"] is not None]
        if date.today().year not in years:
            years.insert(0, date.today().year)
        return sorted(set(years), reverse=True)

    def _annual_entity(self, year: int) -> sqlite3.Row | None:
        return self.conn.execute(
            """
            SELECT * FROM app_entities
            WHERE entity_type=? AND record_key=? AND is_deleted=0
            LIMIT 1
            """,
            (ENTITY_ANNUAL, f"annual:{year}"),
        ).fetchone()

    def get_annual_setting(self, year: int) -> AnnualSetting:
        row = self._annual_entity(year)
        if row is None:
            self.save_annual_setting(AnnualSetting(year=year))
            row = self._annual_entity(year)
        values = self._values_for(int(row["id"]))
        return AnnualSetting(
            year=int(row["year"] or year),
            total_leave_days=safe_float(values.get("total_leave_days", "30"), 30),
            no_leave_base_bonus=safe_float(values.get("no_leave_base_bonus", "16000"), 16000),
            leave_subsidy_per_day=safe_float(values.get("leave_subsidy_per_day", "600"), 600),
            year_end_months=safe_float(values.get("year_end_months", "1.5"), 1.5),
            performance_months=safe_float(values.get("performance_months", "1.0"), 1.0),
        )

    def save_annual_setting(self, s: AnnualSetting) -> None:
        eid = self._upsert_entity(
            entity_type=ENTITY_ANNUAL,
            record_key=f"annual:{s.year}",
            year=s.year,
            record_date=f"{s.year}-01-01",
            note="年度休假與獎金參數",
        )
        self._save_values(
            eid,
            {
                "total_leave_days": s.total_leave_days,
                "no_leave_base_bonus": s.no_leave_base_bonus,
                "leave_subsidy_per_day": s.leave_subsidy_per_day,
                "year_end_months": s.year_end_months,
                "performance_months": s.performance_months,
            },
        )
        self.conn.commit()

    def _salary_from_entity(self, row: sqlite3.Row) -> SalarySetting:
        values = self._values_for(int(row["id"]))
        return SalarySetting(
            id=int(row["id"]),
            year=int(row["year"] or safe_int(values.get("year", str(date.today().year)))),
            effective_date=str(row["record_date"] or values.get("effective_date", f"{date.today().year}-01-01")),
            pay_point=safe_float(values.get("pay_point", "0")),
            salary=safe_float(values.get("salary", "0")),
            professional_1=safe_float(values.get("professional_1", "0")),
            professional_2=safe_float(values.get("professional_2", "0")),
            hazard_allowance=safe_float(values.get("hazard_allowance", "0")),
            retention_allowance=safe_float(values.get("retention_allowance", "0")),
            note=row["note"] or values.get("note", ""),
        )

    def list_salaries(self, year: int | None = None) -> list[SalarySetting]:
        if year is None:
            rows = self.conn.execute(
                """
                SELECT * FROM app_entities
                WHERE entity_type=? AND is_deleted=0
                ORDER BY year DESC, record_date DESC, id DESC
                """,
                (ENTITY_SALARY,),
            ).fetchall()
        else:
            rows = self.conn.execute(
                """
                SELECT * FROM app_entities
                WHERE entity_type=? AND year=? AND is_deleted=0
                ORDER BY record_date DESC, id DESC
                """,
                (ENTITY_SALARY, year),
            ).fetchall()
        return [self._salary_from_entity(r) for r in rows]

    def latest_salary(self, year: int) -> SalarySetting:
        self.ensure_default_year(year)
        rows = self.list_salaries(year)
        if not rows:
            return SalarySetting(None, year, f"{year}-01-01", 0, 0, 0, 0, 0, 0, "")
        return rows[0]

    def save_salary(self, s: SalarySetting) -> int:
        effective = parse_iso_date(s.effective_date).isoformat()
        record_key = ""
        if s.id is None:
            record_key = f"salary:{s.year}:{effective}:{uuid.uuid4().hex[:10]}"
        else:
            row = self.conn.execute("SELECT record_key FROM app_entities WHERE id=?", (s.id,)).fetchone()
            record_key = row["record_key"] if row is not None else f"salary:{s.year}:{effective}:{uuid.uuid4().hex[:10]}"
        eid = self._upsert_entity(
            entity_type=ENTITY_SALARY,
            record_key=record_key,
            year=s.year,
            record_date=effective,
            note=s.note,
            entity_id=s.id,
        )
        self._save_values(
            eid,
            {
                "year": s.year,
                "effective_date": effective,
                "pay_point": s.pay_point,
                "salary": s.salary,
                "professional_1": s.professional_1,
                "professional_2": s.professional_2,
                "hazard_allowance": s.hazard_allowance,
                "retention_allowance": s.retention_allowance,
                "note": s.note,
            },
        )
        self.conn.commit()
        return eid

    def delete_salary(self, salary_id: int) -> None:
        self.conn.execute("UPDATE app_entities SET is_deleted=1, updated_at=? WHERE id=?", (now_text(), salary_id))
        self.conn.commit()

    def list_leave_records(self, year: int) -> list[dict[str, object]]:
        rows = self.conn.execute(
            """
            SELECT * FROM app_entities
            WHERE entity_type=? AND is_deleted=0
              AND record_date >= ? AND record_date <= ?
            ORDER BY record_date, id
            """,
            (ENTITY_LEAVE, f"{year}-01-01", f"{year}-12-31"),
        ).fetchall()
        records: list[dict[str, object]] = []
        for r in rows:
            values = self._values_for(int(r["id"]))
            records.append(
                {
                    "id": int(r["id"]),
                    "leave_date": r["record_date"],
                    "hours": safe_float(values.get("hours", "0")),
                    "leave_type": values.get("leave_type", "休假"),
                    "note": r["note"] or values.get("note", ""),
                }
            )
        return records

    def get_leave_by_date(self, leave_date: str) -> dict[str, object] | None:
        d = parse_iso_date(leave_date)
        row = self.conn.execute(
            """
            SELECT * FROM app_entities
            WHERE entity_type=? AND record_key=? AND is_deleted=0
            LIMIT 1
            """,
            (ENTITY_LEAVE, f"leave:{d.isoformat()}"),
        ).fetchone()
        if row is None:
            return None
        values = self._values_for(int(row["id"]))
        return {
            "id": int(row["id"]),
            "leave_date": row["record_date"],
            "hours": safe_float(values.get("hours", "0")),
            "leave_type": values.get("leave_type", "休假"),
            "note": row["note"] or values.get("note", ""),
        }

    def leave_dates_in_month(self, year: int, month: int) -> set[date]:
        first = date(year, month, 1)
        last_day = calendar.monthrange(year, month)[1]
        last = date(year, month, last_day)
        rows = self.conn.execute(
            """
            SELECT record_date FROM app_entities
            WHERE entity_type=? AND is_deleted=0
              AND record_date >= ? AND record_date <= ?
            """,
            (ENTITY_LEAVE, first.isoformat(), last.isoformat()),
        ).fetchall()
        result: set[date] = set()
        for r in rows:
            try:
                result.add(parse_iso_date(str(r["record_date"])))
            except Exception:
                pass
        return result

    def save_leave(self, record_id: int | None, leave_date: str, hours: float, note: str, leave_type: str = "休假") -> int:
        d = parse_iso_date(leave_date)
        record_key = f"leave:{d.isoformat()}"
        eid = self._upsert_entity(
            entity_type=ENTITY_LEAVE,
            record_key=record_key,
            year=d.year,
            record_date=d.isoformat(),
            note=note,
            entity_id=record_id,
        )
        self._save_values(eid, {"hours": hours, "leave_type": leave_type, "note": note})
        self.conn.commit()
        self.ensure_default_year(d.year)
        return eid

    def save_leave_range(self, dates: list[date], hours: float, note: str, leave_type: str = "休假") -> int:
        count = 0
        for d in dates:
            self.save_leave(None, d.isoformat(), hours, note, leave_type)
            count += 1
        return count

    def delete_leave(self, record_id: int) -> None:
        self.conn.execute("UPDATE app_entities SET is_deleted=1, updated_at=? WHERE id=?", (now_text(), record_id))
        self.conn.commit()

    def calc_year(self, year: int) -> dict[str, float | int | str]:
        annual = self.get_annual_setting(year)
        salary = self.latest_salary(year)
        records = self.list_leave_records(year)
        used_hours = sum(float(r["hours"]) for r in records)
        total_leave_hours = annual.total_leave_days * WORK_HOURS_PER_DAY
        remaining_hours = max(total_leave_hours - used_hours, 0)
        used_days = used_hours / WORK_HOURS_PER_DAY
        remaining_days = remaining_hours / WORK_HOURS_PER_DAY
        base_monthly = salary.base_monthly
        hourly_rate = base_monthly / 30.0 / WORK_HOURS_PER_DAY if base_monthly else 0.0
        remaining_overtime = remaining_hours * hourly_rate
        no_leave_base_bonus = annual.no_leave_base_bonus if used_hours >= 10 * WORK_HOURS_PER_DAY else 0.0
        used_leave_subsidy_days = max(used_days - 10.0, 0.0)
        used_leave_subsidy = used_leave_subsidy_days * annual.leave_subsidy_per_day
        no_leave_bonus = remaining_overtime + no_leave_base_bonus + used_leave_subsidy
        year_end_bonus = base_monthly * annual.year_end_months
        performance_bonus = base_monthly * annual.performance_months
        total_bonus = no_leave_bonus + year_end_bonus + performance_bonus
        return {
            "year": year,
            "record_count": len(records),
            "total_leave_days": annual.total_leave_days,
            "total_leave_hours": total_leave_hours,
            "used_hours": used_hours,
            "used_days": used_days,
            "remaining_hours": remaining_hours,
            "remaining_days": remaining_days,
            "no_leave_award_days": min(used_days, 10.0),
            "used_leave_subsidy_days": used_leave_subsidy_days,
            "remaining_overtime_days": remaining_days,
            "base_monthly": base_monthly,
            "total_monthly": salary.total_monthly,
            "hourly_rate": hourly_rate,
            "remaining_overtime": remaining_overtime,
            "no_leave_base_bonus": no_leave_base_bonus,
            "used_leave_subsidy": used_leave_subsidy,
            "no_leave_bonus": no_leave_bonus,
            "year_end_bonus": year_end_bonus,
            "performance_bonus": performance_bonus,
            "total_bonus": total_bonus,
            "salary_note": salary.note,
        }

    def export_summary_csv(self, path: Path) -> None:
        years = sorted(self.years())
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "年度",
                    "總休假日",
                    "已休假",
                    "剩餘休假",
                    "月支薪津+專業加給",
                    "剩餘休假加班費",
                    "不休假補助",
                    "已休假補助",
                    "不休假獎金合計",
                    "年終獎金",
                    "考績獎金",
                    "獎金總計",
                ]
            )
            for y in years:
                s = self.calc_year(y)
                writer.writerow(
                    [
                        y,
                        s["total_leave_days"],
                        hours_to_text(float(s["used_hours"])),
                        hours_to_text(float(s["remaining_hours"])),
                        round(float(s["base_monthly"])),
                        round(float(s["remaining_overtime"])),
                        round(float(s["no_leave_base_bonus"])),
                        round(float(s["used_leave_subsidy"])),
                        round(float(s["no_leave_bonus"])),
                        round(float(s["year_end_bonus"])),
                        round(float(s["performance_bonus"])),
                        round(float(s["total_bonus"])),
                    ]
                )


class LeaveBonusApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1240x780")
        self.minsize(1020, 660)
        self.db = LeaveBonusDB(app_dir() / DB_FILE)
        self.selected_leave_id: int | None = None
        self.selected_salary_id: int | None = None
        self.year_var = tk.StringVar(value=str(date.today().year))
        self.calendar_year = date.today().year
        self.calendar_month = date.today().month
        self.cal_start_date: date | None = None
        self.cal_end_date: date | None = None
        self.calendar_buttons: dict[date, tk.Button] = {}
        self.calendar_title_vars: list[tk.StringVar] = []
        self.calendar_grids: list[ttk.Frame] = []
        self._build_style()
        self._build_ui()
        self.refresh_all()

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#F5F5F7")
        style.configure("TLabel", background="#F5F5F7", font=("Microsoft JhengHei UI", 10))
        style.configure("Title.TLabel", font=("Microsoft JhengHei UI", 16, "bold"), foreground="#1D1D1F")
        style.configure("Hint.TLabel", font=("Microsoft JhengHei UI", 9), foreground="#6E6E73")
        style.configure("Card.TFrame", background="#FFFFFF", relief="solid", borderwidth=1)
        style.configure("CardTitle.TLabel", background="#FFFFFF", font=("Microsoft JhengHei UI", 10, "bold"), foreground="#6E6E73")
        style.configure("CardValue.TLabel", background="#FFFFFF", font=("Microsoft JhengHei UI", 18, "bold"), foreground="#1D1D1F")
        style.configure("TButton", font=("Microsoft JhengHei UI", 10), padding=(10, 5))
        style.configure("Treeview", font=("Microsoft JhengHei UI", 10), rowheight=27)
        style.configure("Treeview.Heading", font=("Microsoft JhengHei UI", 10, "bold"))

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=(16, 12, 16, 8))
        top.pack(fill="x")
        ttk.Label(top, text=f"{APP_TITLE} {APP_VERSION}", style="Title.TLabel").pack(side="left")
        ttk.Label(top, text="年度：").pack(side="left", padx=(24, 4))
        self.year_combo = ttk.Combobox(top, textvariable=self.year_var, width=8, state="normal")
        self.year_combo.pack(side="left")
        self.year_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_year_changed())
        self.year_combo.bind("<Return>", lambda _e: self.on_year_changed())
        ttk.Button(top, text="重新整理", command=self.refresh_all).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="匯出年度統計 CSV", command=self.export_csv).pack(side="right")

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.tab_dashboard = ttk.Frame(self.nb, padding=12)
        self.tab_leave = ttk.Frame(self.nb, padding=12)
        self.tab_salary = ttk.Frame(self.nb, padding=12)
        # V0.7：年度參數移入第三分頁「薪水加給」，不再另開年度參數分頁。
        self.tab_db = ttk.Frame(self.nb, padding=12)
        self.nb.add(self.tab_dashboard, text="年度總覽")
        self.nb.add(self.tab_leave, text="休假登記")
        self.nb.add(self.tab_salary, text="薪水加給")
        self.nb.add(self.tab_db, text="資料庫說明")

        self._build_dashboard_tab()
        self._build_leave_tab()
        self._build_salary_tab()
        self._build_db_tab()

    def current_year(self) -> int:
        y = safe_int(self.year_var.get(), date.today().year)
        self.db.ensure_default_year(y)
        return y

    def on_year_changed(self) -> None:
        y = self.current_year()
        if self.calendar_year != y:
            self.calendar_year = y
            self.calendar_month = date.today().month if y == date.today().year else 1
            self.clear_calendar_selection(refresh=False)
        self.refresh_all()

    def _card(self, parent: tk.Widget, title: str) -> tuple[ttk.Frame, ttk.Label]:
        frame = ttk.Frame(parent, style="Card.TFrame", padding=12)
        ttk.Label(frame, text=title, style="CardTitle.TLabel").pack(anchor="w")
        label = ttk.Label(frame, text="-", style="CardValue.TLabel")
        label.pack(anchor="w", pady=(6, 0))
        return frame, label

    def _build_dashboard_tab(self) -> None:
        cards = ttk.Frame(self.tab_dashboard)
        cards.pack(fill="x")
        self.card_values: dict[str, ttk.Label] = {}
        card_defs = [
            ("used", "已休假日數"),
            ("remaining", "剩餘休假日數"),
            ("no_leave", "不休假獎金"),
            ("year_end", "年終獎金"),
            ("performance", "考績獎金"),
            ("total", "獎金總計"),
        ]
        for i, (key, title) in enumerate(card_defs):
            frame, label = self._card(cards, title)
            frame.grid(row=0, column=i, sticky="nsew", padx=5, pady=5)
            cards.columnconfigure(i, weight=1)
            self.card_values[key] = label

        detail = ttk.LabelFrame(self.tab_dashboard, text="年度計算明細", padding=8)
        detail.pack(fill="x", pady=(12, 8))
        self.detail_text = tk.Text(detail, height=9, wrap="word", font=("Microsoft JhengHei UI", 10), bg="#FFFFFF")
        self.detail_text.pack(fill="x")
        self.detail_text.configure(state="disabled")

        summary_frame = ttk.LabelFrame(self.tab_dashboard, text="跨年度統計", padding=8)
        summary_frame.pack(fill="both", expand=True, pady=(8, 0))
        cols = ("year", "total_leave", "used", "remaining", "base_monthly", "no_leave", "year_end", "performance", "total")
        self.summary_tree = ttk.Treeview(summary_frame, columns=cols, show="headings")
        headings = {
            "year": "年度",
            "total_leave": "總休假日",
            "used": "已休假",
            "remaining": "剩餘",
            "base_monthly": "月支本薪+專加",
            "no_leave": "不休假獎金",
            "year_end": "年終",
            "performance": "考績",
            "total": "合計",
        }
        widths = {
            "year": 70,
            "total_leave": 90,
            "used": 110,
            "remaining": 110,
            "base_monthly": 120,
            "no_leave": 120,
            "year_end": 100,
            "performance": 100,
            "total": 120,
        }
        for c in cols:
            self.summary_tree.heading(c, text=headings[c])
            self.summary_tree.column(c, width=widths[c], anchor="e" if c != "year" else "center")
        self.summary_tree.pack(fill="both", expand=True, side="left")
        ybar = ttk.Scrollbar(summary_frame, orient="vertical", command=self.summary_tree.yview)
        ybar.pack(side="right", fill="y")
        self.summary_tree.configure(yscrollcommand=ybar.set)

    def _build_leave_tab(self) -> None:
        upper = ttk.Frame(self.tab_leave)
        upper.pack(fill="x")

        calendar_box = ttk.LabelFrame(upper, text="行事曆單日登記", padding=10)
        calendar_box.pack(side="left", fill="both", expand=False, padx=(0, 12))

        cal_top = ttk.Frame(calendar_box)
        cal_top.pack(fill="x", pady=(0, 8))
        ttk.Button(cal_top, text="‹", width=3, command=lambda: self.change_calendar_month(-1)).pack(side="left")
        self.calendar_title_var = tk.StringVar()
        ttk.Label(cal_top, textvariable=self.calendar_title_var, font=("Microsoft JhengHei UI", 12, "bold")).pack(side="left", expand=True)
        ttk.Button(cal_top, text="›", width=3, command=lambda: self.change_calendar_month(1)).pack(side="right")

        self.calendar_title_vars.clear()
        self.calendar_grids.clear()
        self.calendar_wrap = ttk.Frame(calendar_box)
        self.calendar_wrap.pack(fill="both")
        weekdays = ["一", "二", "三", "四", "五", "六", "日"]
        for idx in range(2):
            month_box = ttk.Frame(self.calendar_wrap)
            month_box.grid(row=0, column=idx, sticky="nsew", padx=(0, 8) if idx == 0 else (8, 0))
            self.calendar_wrap.columnconfigure(idx, weight=1)
            title_var = tk.StringVar()
            ttk.Label(month_box, textvariable=title_var, anchor="center", font=("Microsoft JhengHei UI", 10, "bold")).pack(fill="x", pady=(0, 4))
            grid = ttk.Frame(month_box)
            grid.pack(fill="both")
            for i, name in enumerate(weekdays):
                lbl = ttk.Label(grid, text=name, anchor="center", font=("Microsoft JhengHei UI", 10, "bold"))
                lbl.grid(row=0, column=i, sticky="nsew", padx=1, pady=(0, 3))
                grid.columnconfigure(i, weight=1, minsize=36)
            for row in range(1, 7):
                grid.rowconfigure(row, weight=1, minsize=34)
            self.calendar_title_vars.append(title_var)
            self.calendar_grids.append(grid)

        control_box = ttk.LabelFrame(upper, text="登記設定", padding=10)
        control_box.pack(side="left", fill="both", expand=True)
        self.leave_hours_var = tk.StringVar(value="8")
        self.leave_note_var = tk.StringVar()
        self.leave_date_var = tk.StringVar()
        self.selected_range_var = tk.StringVar(value="尚未選取日期")
        ttk.Label(control_box, text="操作方式：點選行事曆日期，或鍵入 YYYY-MM-DD；已登記休假日會以紅底顯示。", style="Hint.TLabel").grid(
            row=0, column=0, columnspan=5, sticky="w", pady=(0, 6)
        )
        ttk.Label(control_box, text="選取日期").grid(row=1, column=0, sticky="w")
        ttk.Label(control_box, textvariable=self.selected_range_var, font=("Microsoft JhengHei UI", 10, "bold")).grid(
            row=2, column=0, columnspan=5, sticky="w", pady=(0, 8)
        )
        ttk.Label(control_box, text="鍵入日期").grid(row=3, column=0, sticky="w")
        ttk.Entry(control_box, textvariable=self.leave_date_var, width=14).grid(row=4, column=0, sticky="w", padx=(0, 12))
        ttk.Button(control_box, text="套用日期", command=self.apply_keyed_leave_date).grid(row=4, column=1, sticky="w", padx=(0, 12))
        ttk.Label(control_box, text="休假時數").grid(row=5, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            control_box,
            textvariable=self.leave_hours_var,
            width=10,
            values=["0", "1", "2", "3", "4", "5", "6", "7", "8"],
        ).grid(row=6, column=0, sticky="w", padx=(0, 12))
        ttk.Label(control_box, text="備註").grid(row=5, column=1, sticky="w", pady=(8, 0))
        ttk.Entry(control_box, textvariable=self.leave_note_var, width=28).grid(row=6, column=1, columnspan=2, sticky="w", padx=(0, 12))
        ttk.Button(control_box, text="登記當日", command=self.save_calendar_leave).grid(row=7, column=0, sticky="w", pady=(8, 0), padx=(0, 6))
        ttk.Button(control_box, text="清除選取", command=self.clear_calendar_selection).grid(row=7, column=1, sticky="w", pady=(8, 0), padx=(0, 6))
        ttk.Button(control_box, text="新增空白", command=self.clear_leave_form).grid(row=7, column=2, sticky="w", pady=(8, 0), padx=(0, 6))
        ttk.Button(control_box, text="刪除明細選取", command=self.delete_leave).grid(row=7, column=3, sticky="w", pady=(8, 0))
        self.leave_hint_var = tk.StringVar(value="已登記日期為紅底；鍵入日期若已有紀錄，會先提示並帶入該筆資料，不會直接覆蓋。")
        ttk.Label(control_box, textvariable=self.leave_hint_var, style="Hint.TLabel").grid(row=8, column=0, columnspan=5, sticky="w", pady=(8, 0))
        control_box.columnconfigure(3, weight=1)

        table_frame = ttk.LabelFrame(self.tab_leave, text="休假明細", padding=8)
        table_frame.pack(fill="both", expand=True, pady=(12, 0))
        cols = ("id", "date", "hours", "decimal", "cum_hours", "used", "remaining", "note")
        self.leave_tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        config = [
            ("id", "ID", 0, "center"),
            ("date", "日期", 105, "center"),
            ("hours", "休假時數", 90, "e"),
            ("decimal", "小數制", 80, "e"),
            ("cum_hours", "總休假時數", 100, "e"),
            ("used", "總休假日數", 130, "center"),
            ("remaining", "剩餘休假日數", 130, "center"),
            ("note", "備註", 300, "w"),
        ]
        for c, h, w, a in config:
            self.leave_tree.heading(c, text=h)
            self.leave_tree.column(c, width=w, anchor=a, stretch=(w > 0))
        self.leave_tree.column("id", width=0, minwidth=0, stretch=False)
        self.leave_tree.pack(fill="both", expand=True, side="left")
        ybar = ttk.Scrollbar(table_frame, orient="vertical", command=self.leave_tree.yview)
        ybar.pack(side="right", fill="y")
        self.leave_tree.configure(yscrollcommand=ybar.set)
        self.leave_tree.bind("<<TreeviewSelect>>", self.on_leave_selected)

    def _build_salary_tab(self) -> None:
        form = ttk.LabelFrame(self.tab_salary, text="新增 / 更正薪水與加給", padding=10)
        form.pack(fill="x")
        self.sal_year_var = tk.StringVar(value=str(date.today().year))
        self.sal_effective_var = tk.StringVar(value=f"{date.today().year}-01-01")
        self.sal_point_var = tk.StringVar(value="0")
        self.sal_salary_var = tk.StringVar(value="0")
        self.sal_pro1_var = tk.StringVar(value="0")
        self.sal_pro2_var = tk.StringVar(value="0")
        self.sal_hazard_var = tk.StringVar(value="0")
        self.sal_retention_var = tk.StringVar(value="0")
        self.sal_note_var = tk.StringVar()
        fields = [
            ("年度", self.sal_year_var, 8),
            ("生效日", self.sal_effective_var, 12),
            ("薪點", self.sal_point_var, 8),
            ("薪津", self.sal_salary_var, 10),
            ("專業加給", self.sal_pro1_var, 10),
            ("專業加給二", self.sal_pro2_var, 10),
            ("危險津貼", self.sal_hazard_var, 10),
            ("留才津貼", self.sal_retention_var, 10),
        ]
        # V0.7：欄位全部由左到右、由上到下依序排列，避免欄位擠到右側。
        for i, (label, var, width) in enumerate(fields):
            row = i // 4
            col = (i % 4) * 2
            ttk.Label(form, text=label).grid(row=row, column=col, sticky="w", padx=(0, 4), pady=3)
            ttk.Entry(form, textvariable=var, width=width).grid(row=row, column=col + 1, sticky="w", padx=(0, 14), pady=3)
        ttk.Label(form, text="備註").grid(row=2, column=0, sticky="w", padx=(0, 4), pady=(8, 0))
        ttk.Entry(form, textvariable=self.sal_note_var, width=18).grid(row=2, column=1, sticky="w", padx=(0, 14), pady=(8, 0))
        ttk.Button(form, text="引入上一年度資料", command=self.import_previous_year_data).grid(row=2, column=2, columnspan=2, sticky="w", padx=(0, 8), pady=(8, 0))
        ttk.Button(form, text="存檔", command=self.save_salary).grid(row=2, column=4, sticky="w", padx=(0, 8), pady=(8, 0))
        ttk.Button(form, text="新增空白", command=self.clear_salary_form).grid(row=2, column=5, sticky="w", padx=(0, 8), pady=(8, 0))
        ttk.Button(form, text="刪除選取", command=self.delete_salary).grid(row=2, column=6, sticky="w", padx=(0, 8), pady=(8, 0))
        form.columnconfigure(7, weight=1)

        # V0.7：年度參數移到薪水加給分頁第二行區域。
        annual_form = ttk.LabelFrame(self.tab_salary, text="年度參數", padding=10)
        annual_form.pack(fill="x", pady=(8, 0))
        self.ann_year_var = tk.StringVar(value=str(date.today().year))
        self.ann_total_days_var = tk.StringVar(value="30")
        self.ann_base_bonus_var = tk.StringVar(value="16000")
        self.ann_subsidy_var = tk.StringVar(value="600")
        self.ann_year_end_var = tk.StringVar(value="1.5")
        self.ann_performance_var = tk.StringVar(value="1.0")
        annual_fields = [
            ("年度", self.ann_year_var, 8),
            ("總休假日", self.ann_total_days_var, 10),
            ("不休假補助", self.ann_base_bonus_var, 12),
            ("已休假補助/日", self.ann_subsidy_var, 12),
            ("年終月數", self.ann_year_end_var, 10),
            ("考績月數", self.ann_performance_var, 10),
        ]
        for i, (label, var, width) in enumerate(annual_fields):
            ttk.Label(annual_form, text=label).grid(row=0, column=i, sticky="w", padx=(0, 10))
            ttk.Entry(annual_form, textvariable=var, width=width).grid(row=1, column=i, sticky="w", padx=(0, 10))
        ttk.Button(annual_form, text="載入年度", command=self.load_annual_form).grid(row=1, column=len(annual_fields), sticky="w", padx=(0, 8))
        ttk.Button(annual_form, text="存檔年度參數", command=self.save_annual).grid(row=1, column=len(annual_fields) + 1, sticky="w", padx=(0, 8))
        annual_form.columnconfigure(len(annual_fields) + 2, weight=1)

        table_frame = ttk.LabelFrame(self.tab_salary, text="薪水與加給每年度資料（可跨年度顯示；計算採該年度最新生效日）", padding=8)
        table_frame.pack(fill="both", expand=True, pady=(12, 0))
        cols = ("id", "year", "effective", "point", "salary", "pro1", "pro2", "hazard", "retention", "base", "total", "note")
        self.salary_tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        headers = [
            ("id", "ID", 0, "center"),
            ("year", "年度", 70, "center"),
            ("effective", "生效日", 100, "center"),
            ("point", "薪點", 70, "e"),
            ("salary", "薪津", 90, "e"),
            ("pro1", "專業加給", 95, "e"),
            ("pro2", "專業加給二", 105, "e"),
            ("hazard", "危險津貼", 95, "e"),
            ("retention", "留才津貼", 95, "e"),
            ("base", "本薪+專加", 110, "e"),
            ("total", "月合計", 100, "e"),
            ("note", "備註", 80, "w"),
        ]
        for c, h, w, a in headers:
            self.salary_tree.heading(c, text=h)
            self.salary_tree.column(c, width=w, anchor=a, stretch=(w > 0))
        self.salary_tree.column("id", width=0, minwidth=0, stretch=False)
        self.salary_tree.pack(fill="both", expand=True, side="left")
        ybar = ttk.Scrollbar(table_frame, orient="vertical", command=self.salary_tree.yview)
        ybar.pack(side="right", fill="y")
        self.salary_tree.configure(yscrollcommand=ybar.set)
        self.salary_tree.bind("<<TreeviewSelect>>", self.on_salary_selected)

    def _build_annual_tab(self) -> None:
        """V0.7：年度參數已移入薪水加給分頁，保留空方法以維持舊版呼叫相容。"""
        return

    def _build_db_tab(self) -> None:
        frame = ttk.LabelFrame(self.tab_db, text="V0.7 通用型資料庫表格", padding=12)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word", font=("Consolas", 10), bg="#FFFFFF")
        text.pack(fill="both", expand=True)
        text.insert(
            "1.0",
            "資料庫檔案：leave_bonus_data.sqlite3\n\n"
            "主要資料表：\n"
            "1) app_entities\n"
            "   - id：資料 ID\n"
            "   - entity_type：資料類型，例如 annual_setting、salary_setting、leave_record\n"
            "   - record_key：通用唯一鍵，例如 annual:2026、leave:2026-07-02\n"
            "   - year：年度\n"
            "   - record_date：日期或生效日\n"
            "   - note：備註\n"
            "   - is_deleted：軟刪除旗標\n"
            "   - created_at / updated_at：建立與更新時間\n\n"
            "2) app_values\n"
            "   - entity_id：對應 app_entities.id\n"
            "   - field_key：欄位名稱，例如 hours、salary、professional_1\n"
            "   - field_value：欄位內容，統一以文字儲存，程式讀取時轉換型別\n"
            "   - value_type：integer、real、text\n\n"
            "這種結構的好處：\n"
            "- 之後增加假別、加給、補助、獎金欄位，不必一直 CREATE 新表或 ALTER 固定欄位。\n"
            "- 不同類型資料都用同一組主表/欄位值表保存。\n"
            "- V0.1 的 annual_settings、salary_settings、leave_records 若存在，首次啟動會自動搬到通用表。\n",
        )
        text.configure(state="disabled")

    # -------- calendar --------
    def change_calendar_month(self, delta: int) -> None:
        m = self.calendar_month + delta
        y = self.calendar_year
        while m < 1:
            m += 12
            y -= 1
        while m > 12:
            m -= 12
            y += 1
        self.calendar_year = y
        self.calendar_month = m
        self.year_var.set(str(y))
        self.refresh_all()

    def calendar_month_by_offset(self, offset: int) -> tuple[int, int]:
        m = self.calendar_month + offset
        y = self.calendar_year
        while m < 1:
            m += 12
            y -= 1
        while m > 12:
            m -= 12
            y += 1
        return y, m

    def refresh_calendar(self) -> None:
        for grid in self.calendar_grids:
            for child in list(grid.grid_slaves()):
                info = child.grid_info()
                if int(info.get("row", 0)) >= 1:
                    child.destroy()
        self.calendar_buttons.clear()
        first_y, first_m = self.calendar_month_by_offset(0)
        second_y, second_m = self.calendar_month_by_offset(1)
        self.calendar_title_var.set(f"顯示：{first_y} 年 {first_m:02d} 月 ～ {second_y} 年 {second_m:02d} 月")
        cal = calendar.Calendar(firstweekday=0)
        today = date.today()

        for idx, grid in enumerate(self.calendar_grids):
            y, m = self.calendar_month_by_offset(idx)
            if idx < len(self.calendar_title_vars):
                self.calendar_title_vars[idx].set(f"{y} 年 {m:02d} 月")
            weeks = cal.monthdatescalendar(y, m)
            while len(weeks) < 6:
                last = weeks[-1][-1]
                weeks.append([last + timedelta(days=i) for i in range(1, 8)])
            for r, week in enumerate(weeks[:6], start=1):
                for c, d in enumerate(week):
                    is_current_month = d.month == m
                    btn = tk.Button(
                        grid,
                        text=str(d.day) if is_current_month else "",
                        width=3,
                        height=1,
                        relief="ridge",
                        bd=1,
                        font=("Microsoft JhengHei UI", 10),
                        command=(lambda day=d: self.on_calendar_day_click(day)) if is_current_month else None,
                    )
                    btn.grid(row=r, column=c, sticky="nsew", padx=1, pady=1)
                    if not is_current_month:
                        btn.configure(state="disabled", bg="#F5F5F7", relief="flat")
                    else:
                        self.calendar_buttons[d] = btn
                        if d == today:
                            btn.configure(bg="#FFF4CE", activebackground="#FFF4CE")
                        else:
                            btn.configure(bg="#FFFFFF", activebackground="#E5F1FF")
        self.update_calendar_highlight()
        self.update_selected_range_label()

    def on_calendar_day_click(self, d: date) -> None:
        """V0.6：點選日期只代表該單日；若該日已有紀錄，直接帶入表單供更正。"""
        self.set_leave_form_date(d, load_existing=True)

    def set_leave_form_date(self, d: date, load_existing: bool = True) -> None:
        self.cal_start_date = d
        self.cal_end_date = None
        self.leave_date_var.set(d.isoformat())
        existing = self.db.get_leave_by_date(d.isoformat()) if load_existing else None
        if existing is not None:
            self.selected_leave_id = int(existing["id"])
            self.leave_hours_var.set(number(float(existing["hours"])))
            self.leave_note_var.set(str(existing["note"] or ""))
            self.select_leave_tree_id(self.selected_leave_id)
        else:
            self.selected_leave_id = None
            self.leave_note_var.set("")
            if hasattr(self, "leave_tree"):
                self.leave_tree.selection_remove(self.leave_tree.selection())
        self.update_calendar_highlight()
        self.update_selected_range_label()

    def apply_keyed_leave_date(self) -> None:
        try:
            d = parse_iso_date(self.leave_date_var.get())
        except Exception:
            messagebox.showerror("日期格式錯誤", "請輸入 YYYY-MM-DD 格式，例如 2026-07-07。")
            return
        self.calendar_year = d.year
        self.calendar_month = d.month
        self.year_var.set(str(d.year))
        self.refresh_calendar()
        self.set_leave_form_date(d, load_existing=True)

    def select_leave_tree_id(self, record_id: int) -> None:
        if not hasattr(self, "leave_tree"):
            return
        for item in self.leave_tree.get_children():
            values = self.leave_tree.item(item, "values")
            if values and int(values[0]) == int(record_id):
                self.leave_tree.selection_set(item)
                self.leave_tree.focus(item)
                self.leave_tree.see(item)
                return

    def selected_dates(self) -> list[date]:
        if self.cal_start_date is None:
            return []
        return [self.cal_start_date]

    def update_calendar_highlight(self) -> None:
        selected = set(self.selected_dates())
        registered: set[date] = set()
        for y, m in sorted({(d.year, d.month) for d in self.calendar_buttons}):
            registered.update(self.db.leave_dates_in_month(y, m))
        today = date.today()
        for d, btn in self.calendar_buttons.items():
            if d in registered:
                btn.configure(
                    bg="#FF3B30",
                    fg="#FFFFFF",
                    activebackground="#D70015",
                    activeforeground="#FFFFFF",
                    relief="sunken" if d in selected else "ridge",
                )
            elif d in selected:
                btn.configure(bg="#007AFF", fg="#FFFFFF", activebackground="#0067D8", activeforeground="#FFFFFF", relief="sunken")
            elif d == today:
                btn.configure(bg="#FFF4CE", fg="#1D1D1F", activebackground="#FFF4CE", activeforeground="#1D1D1F", relief="ridge")
            else:
                btn.configure(bg="#FFFFFF", fg="#1D1D1F", activebackground="#E5F1FF", activeforeground="#1D1D1F", relief="ridge")

    def update_selected_range_label(self) -> None:
        dates = self.selected_dates()
        if not dates:
            self.selected_range_var.set("尚未選取日期")
            return
        if len(dates) == 1:
            self.selected_range_var.set(f"{dates[0].isoformat()}（1 天）")
        else:
            self.selected_range_var.set(f"{dates[0].isoformat()} ～ {dates[-1].isoformat()}（{len(dates)} 天）")

    def clear_calendar_selection(self, refresh: bool = True) -> None:
        self.cal_start_date = None
        self.cal_end_date = None
        self.selected_leave_id = None
        self.leave_note_var.set("")
        self.leave_hours_var.set("8")
        if hasattr(self, "leave_date_var"):
            self.leave_date_var.set("")
        if hasattr(self, "leave_tree"):
            self.leave_tree.selection_remove(self.leave_tree.selection())
        if refresh:
            self.update_calendar_highlight()
            self.update_selected_range_label()

    # -------- refresh --------
    def refresh_all(self) -> None:
        y = self.current_year()
        self.year_combo["values"] = [str(v) for v in self.db.years()]
        self.sal_year_var.set(str(y)) if hasattr(self, "sal_year_var") and not self.selected_salary_id else None
        self.ann_year_var.set(str(y)) if hasattr(self, "ann_year_var") else None
        if self.calendar_year != y:
            self.calendar_year = y
        self.refresh_calendar()
        self.refresh_dashboard()
        self.refresh_leave_tree()
        self.refresh_salary_tree()
        self.load_annual_form_silent()

    def refresh_dashboard(self) -> None:
        y = self.current_year()
        s = self.db.calc_year(y)
        self.card_values["used"].configure(text=hours_to_text(float(s["used_hours"])))
        self.card_values["remaining"].configure(text=hours_to_text(float(s["remaining_hours"])))
        self.card_values["no_leave"].configure(text=money(float(s["no_leave_bonus"])))
        self.card_values["year_end"].configure(text=money(float(s["year_end_bonus"])))
        self.card_values["performance"].configure(text=money(float(s["performance_bonus"])))
        self.card_values["total"].configure(text=money(float(s["total_bonus"])))
        detail = (
            f"年度：{y}\n"
            f"總休假：{number(s['total_leave_days'])} 日；已休：{hours_to_text(float(s['used_hours']))}，剩餘：{hours_to_text(float(s['remaining_hours']))}\n"
            f"不休假獎金日數：{number(s['no_leave_award_days'])} 日；休假補助日數：{number(s['used_leave_subsidy_days'])} 日；未休假加班費日數：{number(s['remaining_overtime_days'])} 日\n"
            f"月支本薪+專業加給：{money(float(s['base_monthly']))}；月支合計含危險/留才：{money(float(s['total_monthly']))}\n"
            f"每小時加班費基準：{number(s['hourly_rate'])}\n"
            f"剩餘休假加班費：{money(float(s['remaining_overtime']))}；不休假補助：{money(float(s['no_leave_base_bonus']))}；已休假補助：{money(float(s['used_leave_subsidy']))}\n"
            f"不休假獎金合計：{money(float(s['no_leave_bonus']))}；年終：{money(float(s['year_end_bonus']))}；考績：{money(float(s['performance_bonus']))}\n"
        )
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", detail)
        self.detail_text.configure(state="disabled")

        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        for yy in sorted(self.db.years(), reverse=True):
            sy = self.db.calc_year(yy)
            self.summary_tree.insert(
                "",
                "end",
                values=(
                    yy,
                    number(sy["total_leave_days"]),
                    hours_to_text(float(sy["used_hours"])),
                    hours_to_text(float(sy["remaining_hours"])),
                    money(float(sy["base_monthly"])),
                    money(float(sy["no_leave_bonus"])),
                    money(float(sy["year_end_bonus"])),
                    money(float(sy["performance_bonus"])),
                    money(float(sy["total_bonus"])),
                ),
            )

    def refresh_leave_tree(self) -> None:
        for item in self.leave_tree.get_children():
            self.leave_tree.delete(item)
        y = self.current_year()
        annual = self.db.get_annual_setting(y)
        total_hours = annual.total_leave_days * WORK_HOURS_PER_DAY
        cum_hours = 0.0
        cum_decimal = 0.0
        for r in self.db.list_leave_records(y):
            h = float(r["hours"])
            cum_hours += h
            cum_decimal += leave_decimal(h)
            remain = total_hours - cum_hours
            self.leave_tree.insert(
                "",
                "end",
                values=(
                    int(r["id"]),
                    r["leave_date"],
                    number(h),
                    number(cum_decimal),
                    number(cum_hours),
                    hours_to_text(cum_hours),
                    hours_to_text(remain),
                    r["note"] or "",
                ),
            )

    def refresh_salary_tree(self) -> None:
        for item in self.salary_tree.get_children():
            self.salary_tree.delete(item)
        for s in self.db.list_salaries(None):
            self.salary_tree.insert(
                "",
                "end",
                values=(
                    s.id,
                    s.year,
                    s.effective_date,
                    number(s.pay_point),
                    money(s.salary),
                    money(s.professional_1),
                    money(s.professional_2),
                    money(s.hazard_allowance),
                    money(s.retention_allowance),
                    money(s.base_monthly),
                    money(s.total_monthly),
                    s.note,
                ),
            )

    # -------- leave actions --------
    def save_calendar_leave(self) -> None:
        try:
            typed_date = self.leave_date_var.get().strip() if hasattr(self, "leave_date_var") else ""
            if typed_date:
                target_date = parse_iso_date(typed_date)
            else:
                dates = self.selected_dates()
                if not dates:
                    messagebox.showinfo("請先選取日期", "請在行事曆上點選要登記的休假日期，或鍵入 YYYY-MM-DD。")
                    return
                target_date = dates[0]
            h = safe_float(self.leave_hours_var.get())
            if h < 0 or h > 8:
                raise ValueError("休假時數請輸入 0 到 8。")
            note = self.leave_note_var.get().strip()

            existing = self.db.get_leave_by_date(target_date.isoformat())
            if existing is not None and int(existing["id"]) != int(self.selected_leave_id or 0):
                self.year_var.set(str(target_date.year))
                self.calendar_year = target_date.year
                self.calendar_month = target_date.month
                self.refresh_all()
                self.set_leave_form_date(target_date, load_existing=True)
                messagebox.showwarning(
                    "日期重複",
                    f"{target_date.isoformat()} 已有休假紀錄，已自動帶入該筆資料。\n請直接更正後再按登記當日，不會建立重複日期。",
                )
                return

            if self.selected_leave_id is not None:
                self.db.save_leave(self.selected_leave_id, target_date.isoformat(), h, note)
            else:
                self.db.save_leave(None, target_date.isoformat(), h, note)
            self.year_var.set(str(target_date.year))
            self.calendar_year = target_date.year
            self.calendar_month = target_date.month
            saved_id = self.db.get_leave_by_date(target_date.isoformat())
            self.selected_leave_id = int(saved_id["id"]) if saved_id else None
            self.refresh_all()
            if self.selected_leave_id is not None:
                self.select_leave_tree_id(self.selected_leave_id)
            # V0.6：休假登記存檔後不跳出成功確認視窗，重複日期才提示。
        except Exception as exc:
            messagebox.showerror("無法存檔", str(exc))

    def clear_leave_form(self) -> None:
        self.clear_calendar_selection()

    def delete_leave(self) -> None:
        if self.selected_leave_id is None:
            messagebox.showinfo("請先選取", "請先選取要刪除的休假紀錄。")
            return
        if messagebox.askyesno("確認刪除", "確定刪除選取的休假紀錄？"):
            self.db.delete_leave(self.selected_leave_id)
            self.clear_leave_form()
            self.refresh_all()

    def on_leave_selected(self, _event=None) -> None:
        sel = self.leave_tree.selection()
        if not sel:
            return
        values = self.leave_tree.item(sel[0], "values")
        self.selected_leave_id = int(values[0])
        d = parse_iso_date(str(values[1]))
        self.calendar_year = d.year
        self.calendar_month = d.month
        self.year_var.set(str(d.year))
        self.cal_start_date = d
        self.cal_end_date = None
        self.leave_date_var.set(d.isoformat())
        self.leave_hours_var.set(str(values[2]).replace(",", ""))
        self.leave_note_var.set(str(values[7]))
        self.refresh_calendar()

    # -------- salary actions --------
    def import_previous_year_data(self) -> None:
        """V0.7：把上一年度的薪水加給與年度參數帶入目前年度表單，供新年度快速建立。"""
        try:
            target_year = safe_int(self.sal_year_var.get() or self.year_var.get(), self.current_year())
            if target_year <= 1:
                target_year = self.current_year()
            previous_year = target_year - 1

            previous_salaries = self.db.list_salaries(previous_year)
            if not previous_salaries:
                messagebox.showinfo("沒有上一年度資料", f"找不到 {previous_year} 年的薪水加給資料。")
                return
            prev_salary = previous_salaries[0]
            self.selected_salary_id = None
            self.sal_year_var.set(str(target_year))
            self.sal_effective_var.set(f"{target_year}-01-01")
            self.sal_point_var.set(number(prev_salary.pay_point))
            self.sal_salary_var.set(number(prev_salary.salary))
            self.sal_pro1_var.set(number(prev_salary.professional_1))
            self.sal_pro2_var.set(number(prev_salary.professional_2))
            self.sal_hazard_var.set(number(prev_salary.hazard_allowance))
            self.sal_retention_var.set(number(prev_salary.retention_allowance))
            self.sal_note_var.set(prev_salary.note if prev_salary.note != "請填入當年薪水與加給" else "")

            prev_annual = self.db.get_annual_setting(previous_year)
            self.ann_year_var.set(str(target_year))
            self.ann_total_days_var.set(number(prev_annual.total_leave_days))
            self.ann_base_bonus_var.set(number(prev_annual.no_leave_base_bonus))
            self.ann_subsidy_var.set(number(prev_annual.leave_subsidy_per_day))
            self.ann_year_end_var.set(number(prev_annual.year_end_months))
            self.ann_performance_var.set(number(prev_annual.performance_months))

            if hasattr(self, "salary_tree"):
                self.salary_tree.selection_remove(self.salary_tree.selection())
        except Exception as exc:
            messagebox.showerror("無法引入", str(exc))

    def save_salary(self) -> None:
        try:
            y = safe_int(self.sal_year_var.get())
            effective = parse_iso_date(self.sal_effective_var.get()).isoformat()
            s = SalarySetting(
                id=self.selected_salary_id,
                year=y,
                effective_date=effective,
                pay_point=safe_float(self.sal_point_var.get()),
                salary=safe_float(self.sal_salary_var.get()),
                professional_1=safe_float(self.sal_pro1_var.get()),
                professional_2=safe_float(self.sal_pro2_var.get()),
                hazard_allowance=safe_float(self.sal_hazard_var.get()),
                retention_allowance=safe_float(self.sal_retention_var.get()),
                note=self.sal_note_var.get().strip(),
            )
            self.db.save_salary(s)
            self.year_var.set(str(y))
            self.selected_salary_id = None
            self.clear_salary_form(keep_year=True)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("無法存檔", str(exc))

    def clear_salary_form(self, keep_year: bool = False) -> None:
        self.selected_salary_id = None
        y = self.current_year() if keep_year else date.today().year
        self.sal_year_var.set(str(y))
        self.sal_effective_var.set(f"{y}-01-01")
        self.sal_point_var.set("0")
        self.sal_salary_var.set("0")
        self.sal_pro1_var.set("0")
        self.sal_pro2_var.set("0")
        self.sal_hazard_var.set("0")
        self.sal_retention_var.set("0")
        self.sal_note_var.set("")
        self.salary_tree.selection_remove(self.salary_tree.selection())

    def delete_salary(self) -> None:
        if self.selected_salary_id is None:
            messagebox.showinfo("請先選取", "請先選取要刪除的薪水加給紀錄。")
            return
        if messagebox.askyesno("確認刪除", "確定刪除選取的薪水加給紀錄？"):
            self.db.delete_salary(self.selected_salary_id)
            self.clear_salary_form(keep_year=True)
            self.refresh_all()

    def on_salary_selected(self, _event=None) -> None:
        sel = self.salary_tree.selection()
        if not sel:
            return
        values = self.salary_tree.item(sel[0], "values")
        self.selected_salary_id = int(values[0])
        self.sal_year_var.set(str(values[1]))
        self.sal_effective_var.set(str(values[2]))
        self.sal_point_var.set(str(values[3]).replace(",", ""))
        self.sal_salary_var.set(str(values[4]).replace(",", ""))
        self.sal_pro1_var.set(str(values[5]).replace(",", ""))
        self.sal_pro2_var.set(str(values[6]).replace(",", ""))
        self.sal_hazard_var.set(str(values[7]).replace(",", ""))
        self.sal_retention_var.set(str(values[8]).replace(",", ""))
        self.sal_note_var.set(str(values[11]))

    # -------- annual actions --------
    def load_annual_form_silent(self) -> None:
        try:
            y = safe_int(self.ann_year_var.get() or self.year_var.get(), self.current_year())
            s = self.db.get_annual_setting(y)
            self.ann_year_var.set(str(s.year))
            self.ann_total_days_var.set(number(s.total_leave_days))
            self.ann_base_bonus_var.set(number(s.no_leave_base_bonus))
            self.ann_subsidy_var.set(number(s.leave_subsidy_per_day))
            self.ann_year_end_var.set(number(s.year_end_months))
            self.ann_performance_var.set(number(s.performance_months))
        except Exception:
            pass

    def load_annual_form(self) -> None:
        try:
            y = safe_int(self.ann_year_var.get() or self.year_var.get(), self.current_year())
            s = self.db.get_annual_setting(y)
            self.ann_year_var.set(str(s.year))
            self.ann_total_days_var.set(number(s.total_leave_days))
            self.ann_base_bonus_var.set(number(s.no_leave_base_bonus))
            self.ann_subsidy_var.set(number(s.leave_subsidy_per_day))
            self.ann_year_end_var.set(number(s.year_end_months))
            self.ann_performance_var.set(number(s.performance_months))
        except Exception as exc:
            messagebox.showerror("無法載入", str(exc))

    def save_annual(self) -> None:
        try:
            y = safe_int(self.ann_year_var.get())
            s = AnnualSetting(
                year=y,
                total_leave_days=safe_float(self.ann_total_days_var.get(), 30),
                no_leave_base_bonus=safe_float(self.ann_base_bonus_var.get(), 16000),
                leave_subsidy_per_day=safe_float(self.ann_subsidy_var.get(), 600),
                year_end_months=safe_float(self.ann_year_end_var.get(), 1.5),
                performance_months=safe_float(self.ann_performance_var.get(), 1.0),
            )
            self.db.save_annual_setting(s)
            self.year_var.set(str(y))
            self.calendar_year = y
            self.refresh_all()
            messagebox.showinfo("已存檔", "年度參數已更新。")
        except Exception as exc:
            messagebox.showerror("無法存檔", str(exc))

    def export_csv(self) -> None:
        default = f"休假獎金年度統計_{datetime.now().strftime('%Y%m%d%H%M%S')}.csv"
        path = filedialog.asksaveasfilename(
            title="匯出年度統計 CSV",
            defaultextension=".csv",
            initialfile=default,
            filetypes=[("CSV 檔案", "*.csv"), ("所有檔案", "*.*")],
        )
        if not path:
            return
        try:
            self.db.export_summary_csv(Path(path))
            messagebox.showinfo("匯出完成", f"已匯出：\n{path}")
        except Exception as exc:
            messagebox.showerror("匯出失敗", str(exc))


if __name__ == "__main__":
    app = LeaveBonusApp()
    app.mainloop()
