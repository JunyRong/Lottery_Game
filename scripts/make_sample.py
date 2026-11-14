"""生成示例文件：sample/participants.xlsx 与 sample/prizes.xlsx。"""

from pathlib import Path

from openpyxl import Workbook

NAMES = [
    "张伟", "王芳", "李娜", "刘洋", "陈静", "杨帆", "赵磊", "黄敏", "周杰", "吴倩",
    "徐强", "孙丽", "马超", "朱婷", "胡军", "郭静", "何勇", "高敏", "林峰", "罗娜",
    "郑浩", "梁静", "谢军", "唐敏", "许峰", "邓丽", "冯强", "曾芳", "彭勇", "肖婷",
    "田亮", "董洁", "袁晨", "潘婷", "蒋雯", "蔡明", "余波", "杜鹃", "叶枫", "程琳",
]
DEPARTMENTS = ["研发部", "市场部", "人事部", "财务部", "运营部"]
TIERS = [
    ("一等奖", "笔记本电脑", 1),
    ("二等奖", "蓝牙耳机", 2),
    ("三等奖", "保温杯", 3),
]


def make_participants(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "抽奖名单"
    ws.append(["序号", "姓名", "部门"])
    for i, name in enumerate(NAMES, start=1):
        ws.append([i, name, DEPARTMENTS[(i - 1) % len(DEPARTMENTS)]])
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 14
    wb.save(path)


def make_prizes(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "奖项设置"
    ws.append(["奖项名称", "奖品", "获奖人数"])
    for row in TIERS:
        ws.append(list(row))
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 12
    wb.save(path)


def main() -> None:
    sample_dir = Path(__file__).resolve().parent.parent / "sample"
    sample_dir.mkdir(exist_ok=True)
    participants = sample_dir / "participants.xlsx"
    prizes = sample_dir / "prizes.xlsx"
    make_participants(participants)
    make_prizes(prizes)
    print(f"已生成 {participants}（{len(NAMES)} 人）")
    print(f"已生成 {prizes}（{len(TIERS)} 个奖项）")


if __name__ == "__main__":
    main()
