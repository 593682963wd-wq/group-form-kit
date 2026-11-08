"""作者：Wangyuan；日期：2026-10-29；职责：计划分组、可配置收集表与离线回传。"""
from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from contextlib import closing
from datetime import datetime
from pathlib import Path
from typing import Any


def canonical(value: Any) -> str:
    """生成稳定数据指纹。"""
    return json.dumps(value, ensure_ascii=False, sort_keys=True, allow_nan=False, separators=(",", ":"))


def weekdays(value: str) -> str:
    """规范化1至7的执行星期；未知字符不静默吞掉。"""
    clean = re.sub(r"[\s,，/;；、]", "", str(value))
    if not clean or any(x not in "1234567" for x in clean): raise ValueError("星期须为1至7")
    return "".join(x for x in "1234567" if x in clean)


def clock(value: str) -> int:
    """解析明确HH:MM时刻。"""
    if not re.fullmatch(r"\d{1,2}:\d{2}", str(value)): raise ValueError("时刻须为HH:MM")
    hour, minute = map(int, str(value).split(":"))
    if not 0 <= hour < 24 or not 0 <= minute < 60: raise ValueError("时刻越界")
    return hour * 60 + minute


def duration(start: str, end: str, end_day_offset: int) -> int:
    """跨日次数显式指定，不根据负差擅自猜第二天。"""
    if type(end_day_offset) is not int or end_day_offset < 0: raise ValueError("跨日天数无效")
    result = clock(end) + end_day_offset * 1440 - clock(start)
    if result <= 0: raise ValueError("时长须为正")
    return result


def parse_table(rows: list[list[Any]], columns: dict[str, str], header_search: int = 10) -> list[dict]:
    """在前若干行动态匹配调用方提供的表头合同，不猜具体季节格式。"""
    required = {"id", "group", "origin", "destination", "weekdays", "start", "end", "end_day_offset"}
    if set(columns) != required or len(set(columns.values())) != len(columns): raise ValueError("列合同须完整且不重复")
    header = None
    for index, row in enumerate(rows[:header_search]):
        texts = [str(x).strip() if x is not None else "" for x in row]
        if set(columns.values()) <= set(texts):
            if any(texts.count(name) != 1 for name in columns.values()): raise ValueError("表头重复")
            header = (index, {key: texts.index(name) for key, name in columns.items()}); break
    if header is None: raise ValueError("未找到完整表头")
    result = []; seen = set()
    for number, row in enumerate(rows[header[0] + 1:], header[0] + 2):
        if not any(x not in (None, "") for x in row): continue
        if len(row) <= max(header[1].values()): raise ValueError(f"第{number}行缺列")
        item = {key: str(row[index]).strip() if row[index] is not None else "" for key, index in header[1].items()}
        if not all(item[k] for k in required) or item["id"] in seen: raise ValueError(f"第{number}行缺值或身份重复")
        seen.add(item["id"]); item["weekdays"] = weekdays(item["weekdays"])
        offset = item["end_day_offset"]
        if not re.fullmatch(r"\d+", offset): raise ValueError("跨日偏移须为整数")
        item["end_day_offset"] = int(offset)
        item["duration_minutes"] = duration(item["start"], item["end"], item["end_day_offset"])
        item["source_row"] = number; result.append(item)
    return result


def read_xlsx(path: str | Path, sheet: str, columns: dict[str, str]) -> list[dict]:
    """读取显式工作表，不把公式无缓存值当作已知值。"""
    from openpyxl import load_workbook
    workbook = load_workbook(path, read_only=True, data_only=True)
    formulas = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet not in workbook.sheetnames: raise ValueError("工作表不存在")
        worksheet = workbook[sheet]; source = formulas[sheet]
        if (worksheet.max_row or 0) > 100_000 or (worksheet.max_column or 0) > 100: raise ValueError("工作表超过预算")
        rows = []
        from itertools import zip_longest
        for number, (values, cells) in enumerate(zip_longest(worksheet.iter_rows(values_only=True), source.iter_rows(), fillvalue=()),1):
            if number>100_000 or max(len(values),len(cells))>100:raise ValueError("工作表超过预算")
            for index,cell in enumerate(cells):
                if cell.data_type=="f" and (index>=len(values) or values[index] is None):raise ValueError(f"第{number}行含无缓存公式，须先重新计算工作簿")
            rows.append(list(values))
        return parse_table(rows, columns)
    finally:
        workbook.close(); formulas.close()


def group_plan(rows: list[dict], *, origin: str | None = None, destination: str | None = None,
               split_weekdays: bool = False) -> list[dict]:
    """保持方向筛选，按计划组和可选星期拆组，驻点候选来自显式地点。"""
    groups = {}; seen = set()
    for row in rows:
        if row["id"] in seen: raise ValueError("计划身份重复")
        seen.add(row["id"])
        if origin is not None and row["origin"] != origin or destination is not None and row["destination"] != destination: continue
        days = weekdays(row["weekdays"])
        for part in days if split_weekdays else [days]:
            key = f"{row['group']}:{part}" if split_weekdays else row["group"]
            group = groups.setdefault(key, {"id": key, "weekdays": "", "items": [], "stations": []})
            group["weekdays"] = weekdays(group["weekdays"] + part)
            group["items"].append(row)
            group["stations"] = sorted(set(group["stations"]) | {row["origin"], row["destination"]})
    return sorted(groups.values(), key=lambda x: x["id"])


def package(plan_id: str, groups: list[dict], teams: list[str], fields: list[dict]) -> dict:
    """生成自描述离线填写包，不携带地址、凭据或真实人员。"""
    if not plan_id or len(set(teams)) != len(teams) or not teams or len({x["id"] for x in groups}) != len(groups): raise ValueError("计划/部门/分组身份无效")
    if len({x["name"] for x in fields}) != len(fields) or any(x["type"] not in {"text", "integer", "duration"} for x in fields): raise ValueError("收集字段合同无效")
    body = {"schema": 1, "plan": plan_id, "groups": groups, "teams": teams, "fields": fields}
    return {**body, "fingerprint": hashlib.sha256(canonical(body).encode()).hexdigest()}


def validate_package(value: dict) -> None:
    """核对离线包未被无意改写；指纹不代替签名或身份认证。"""
    body = {k: v for k, v in value.items() if k != "fingerprint"}
    if value.get("schema") != 1 or hashlib.sha256(canonical(body).encode()).hexdigest() != value.get("fingerprint"): raise ValueError("离线包版本或指纹不符")
    rebuilt = package(value["plan"], value["groups"], value["teams"], value["fields"])
    if rebuilt != value: raise ValueError("离线包含未知字段")


class Collection:
    """每个组/团队独立版本，离线回传幂等并明确完整率。"""

    def __init__(self, path: str | Path, contract: dict) -> None:
        """以显式填写包初始化，不连接网络。"""
        validate_package(contract); self.contract = contract; self.path = str(path)
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute("CREATE TABLE IF NOT EXISTS entries(plan TEXT,team TEXT,group_id TEXT,body TEXT,version INTEGER,PRIMARY KEY(plan,team,group_id))")

    def submit(self, team: str, group_id: str, values: dict, station: str,
               *, expected_version: int, fingerprint: str) -> dict:
        """回传必须对应冻结合同；数值无效时整条拒绝，不截断取整。"""
        if fingerprint != self.contract["fingerprint"] or team not in self.contract["teams"]: raise ValueError("回传版本或团队不符")
        groups = {x["id"]: x for x in self.contract["groups"]}
        if group_id not in groups or station not in groups[group_id]["stations"]: raise ValueError("组或驻点不符")
        fields = {x["name"]: x for x in self.contract["fields"]}
        if set(values) - set(fields): raise ValueError("未知字段")
        cleaned = {}
        for key, field in fields.items():
            value = values.get(key)
            if value is None or value == "":
                if field.get("required"): raise ValueError("必填项缺失")
                cleaned[key] = None; continue
            if field["type"] == "integer":
                if type(value) is not int or value < 0: raise ValueError("人数须为非负整数")
            elif field["type"] == "duration":
                if not isinstance(value, str) or not re.fullmatch(r"\d+:\d{2}", value) or int(value.split(":")[1]) >= 60: raise ValueError("时长须为HH:MM")
                hours, minutes = map(int, value.split(":")); value = hours * 60 + minutes
            elif not isinstance(value, str): raise ValueError("文本字段类型无效")
            cleaned[key] = value
        body = canonical({"station": station, "values": cleaned, "fingerprint": fingerprint})
        identity = (self.contract["plan"], team, group_id)
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT body,version FROM entries WHERE plan=? AND team=? AND group_id=?", identity).fetchone()
            if row and row[0] == body: return {"version": row[1], "reused": True}
            version = row[1] if row else 0
            if expected_version != version: raise ValueError("回传版本冲突")
            db.execute("INSERT OR REPLACE INTO entries VALUES(?,?,?,?,?)", (*identity, body, version + 1))
            return {"version": version + 1, "reused": False}

    def summary(self) -> dict:
        """统计当前合同有效回传，旧合同材料不算完成。"""
        with closing(sqlite3.connect(self.path)) as db, db:
            rows = [{"team": t, "group": g, "version": v, **json.loads(b)} for t, g, b, v in db.execute("SELECT team,group_id,body,version FROM entries WHERE plan=?", (self.contract["plan"],))]
        rows = [x for x in rows if x["fingerprint"] == self.contract["fingerprint"]]
        done = {(x["team"], x["group"]) for x in rows}
        return {"rows": rows, "completion": [{"team": team, "done": sum((team, g["id"]) in done for g in self.contract["groups"]), "total": len(self.contract["groups"]),
                  "missing": [g["id"] for g in self.contract["groups"] if (team, g["id"]) not in done]} for team in self.contract["teams"]]}


def write_template(contract: dict, path: str | Path) -> None:
    """从零生成中性Excel收集表，不读取原模板。"""
    from openpyxl import Workbook
    from openpyxl.worksheet.datavalidation import DataValidation
    validate_package(contract); workbook = Workbook(); worksheet = workbook.active; worksheet.title = "收集表"
    worksheet.append(["plan", "team", "group", "station", *[x["name"] for x in contract["fields"]]])
    for team in contract["teams"]:
        for group in contract["groups"]:
            worksheet.append([contract["plan"], team, group["id"], None] + [None] * len(contract["fields"]))
    worksheet.freeze_panes = "E2"; worksheet.auto_filter.ref = worksheet.dimensions
    metadata = workbook.create_sheet("合同"); metadata.append(["fingerprint", contract["fingerprint"]]); metadata.append(["schema", 1])
    for index, field in enumerate(contract["fields"], 5):
        if field["type"] == "integer" and worksheet.max_row >= 2:
            check = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=not field.get("required", False)); check.error = "请输入非负整数"; check.showErrorMessage = True; worksheet.add_data_validation(check)
            check.add(f"{worksheet.cell(1,index).column_letter}2:{worksheet.cell(1,index).column_letter}{worksheet.max_row}")
    # 合同标签都是文本；禁止字符串被电子表格解释为公式。
    for sheet in workbook:
        for row in sheet:
            for cell in row:
                if isinstance(cell.value, str): cell.data_type = "s"
    workbook.save(path); workbook.close()
