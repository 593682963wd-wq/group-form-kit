"""作者：Wangyuan；日期：2026-10-29；职责：独立离线JSON命令行入口。"""
from __future__ import annotations
import argparse
import json
import sys
import tempfile
from pathlib import Path
from typing import Any
from collection import parse_table, read_xlsx, group_plan, package, Collection, write_template
def run(payload: dict, output: Path) -> Any:
    """读取显式计划、生成表或回传，状态保存在--output。"""
    action = payload["action"]; args = payload.get("args", {})
    actions = {"parse": parse_table, "read_xlsx": read_xlsx, "group": group_plan, "package": package}
    if action in actions: return actions[action](**args)
    contract = payload["contract"]
    if action == "template":
        write_template(contract, output / "collection.xlsx"); return {"output": "collection.xlsx"}
    store = Collection(output / "collection.sqlite", contract)
    if action == "submit": return store.submit(**args)
    if action == "summary": return store.summary()
    raise ValueError("未知动作")

def demo(output: Path) -> dict:
    """生成两个合成团队的填写包、Excel与一次回传。"""
    rows = [{"id": "item-1", "group": "group-1", "origin": "地点甲", "destination": "地点乙", "weekdays": "135", "start": "08:00", "end": "09:00", "end_day_offset": 0}]
    contract = package("plan-demo", group_plan(rows), ["team-A", "team-B"], [{"name": "count", "type": "integer", "required": True}])
    write_template(contract, output / "collection.xlsx")
    (output / "package.json").write_text(__import__("json").dumps(contract, ensure_ascii=False, indent=2), encoding="utf-8")
    store = Collection(output / "collection.sqlite", contract)
    store.submit("team-A", "group-1", {"count": 2}, "地点甲", expected_version=0, fingerprint=contract["fingerprint"])
    return {"package": contract, "summary": store.summary(), "template": "collection.xlsx"}

def main() -> int:
    """运行合成演示或显式输入；错误返回非零，不访问外部服务。"""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=["demo", "run"])
    parser.add_argument("input", nargs="?")
    parser.add_argument("--output", help="显式输出目录；默认临时目录")
    args = parser.parse_args()
    try:
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(args.output or temporary)
            output.mkdir(parents=True, exist_ok=True)
            if args.command == "demo":
                result = demo(output)
            else:
                if not args.input:
                    raise ValueError("run必须提供JSON文件")
                source = Path(args.input)
                if source.stat().st_size > 10_000_000:
                    raise ValueError("JSON文件超出预算")
                result = run(json.loads(source.read_text(encoding="utf-8")), output)
            print(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False))
        return 0
    except (ValueError, TypeError, KeyError, OSError) as exc:
        print(json.dumps({"error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2
if __name__ == "__main__":
    raise SystemExit(main())
