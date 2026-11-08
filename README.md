# 可配置分组收集表与离线回传

作者：Wangyuan · 2026-10-29

本项目源码采用 MIT 许可证，见 LICENSE；第三方依赖仍适用各自上游许可。

## 可用能力

- 按显式表头映射读取表格/XLSX，星期规范化、HH:MM与显式跨日分钟校验、稳定身份去重。
- 方向筛选、分组与按星期拆组、由输入地点导出驻点候选。
- 显式团队/字段合同、整数/时长/文本校验，带指纹离线JSON填写包、从零生成Excel收集表与数据验证。
- SQLite每团队/分组乐观版本、同内容重试、合同版本检查、有效完成率和缺组清单，JSON汇总导出。

## 运行

Python 3.11+。依赖仅为标准库，或 requirements.txt 明确列出的开源解析库。程序默认不联网；安装依赖时由包管理器下载。

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 -B cli.py demo
python3 -B cli.py demo --output demo-output
python3 -B cli.py run example.json --output demo-output
python3 -B -m unittest discover -s tests -v
```

演示数据从零生成；默认临时目录会在命令结束时释放。需要保留文件或数据库时必须指定 --output，输入与输出目录应由使用者自行控制。JSON输出到标准输出，参数错误返回非零状态。

## 输入与输出

动作：`parse`、`read_xlsx`、`group`、`package`、`template`、`submit`、`summary`。每个动作参数名称对应核心文件中的同名函数/方法签名，公共接口有类型标注及中文说明。通用调用格式是 action 与 args；目录、收集模块另接受 tools/events/contract 等顶层显式合同。

```json
{
  "action": "group",
  "args": {
    "rows": [
      {
        "id": "item-1",
        "group": "group-1",
        "origin": "地点甲",
        "destination": "地点乙",
        "weekdays": "135",
        "start": "08:00",
        "end": "09:00",
        "end_day_offset": 0
      }
    ],
    "split_weekdays": true
  }
}
```

XLSX解析使用显式sheet及columns，不猜某年度格式。填写包指纹防止无意改写，不是身份认证或数字签名。Excel为可编辑空表；回传走JSON submit（不自动把任意Excel当成可信回传）。相同计划ID更换合同后旧回传不计完成，建议新计划使用新ID。

## 保留边界

实际三部门/班组/人员、航空类别/时长法规、原值勤模板、云隧道、公司提交接口和原Web填写页未包含。

只保留通用流程与算法；这是经过提取/重构的独立工具或组件集合，并不声称完整迁移原应用、原业务质量或生产验收。源码来源、哈希与原入口映射保存在本地私有审计台账，不随此目录分发。

## 验证

tests/仅使用合成数据与临时文件，覆盖主流程、空/缺失、身份冲突和关键异常；cli.py demo可独立复现完整示例。未运行真实业务数据、联网/正式发送、原系统回归或生产部署。本项目为CLI/库，不包含Web UI。

## 2026-11-08 候选修订

- 未知工作表维度可读取，逐行实施100000×100预算。
- 同时读公式视图，遇无缓存公式明确拒绝，不静默跳过整行。

使用包内合成样例复跑；发布集合根目录的 `verify_release.py --project workflow-collection-forms` 会依次执行测试和首次使用命令。原有版本的测试计数属于当时记录，本轮汇总以集合验证报告为准。
