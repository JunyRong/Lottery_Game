"""无窗口 UI 冒烟测试：构建界面、抽奖一次，并验证手动名单编辑。"""

import sys
from pathlib import Path
import tkinter as tk

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from lottery.excel_io import read_names  # noqa: E402
from lottery.models import Entrant, Winner  # noqa: E402
from lottery.ui import LotteryApp, ManualInputDialog, TierDialog  # noqa: E402


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    app = LotteryApp(root)

    sample = BASE_DIR / "sample" / "participants.xlsx"
    app.names = read_names(sample, "抽奖名单", 1)
    app._rebuild_engine()
    app._refresh_all()
    assert app.selected_tier_id == app.tiers[-1].id, "默认应从最低奖项开始"

    # 奖项顺序调整：上移/下移保持选中项，添加时可选择插入位置
    if len(app.tiers) >= 2:
        original_ids = [t.id for t in app.tiers]
        last_id = original_ids[-1]
        app._move_tier(-1)
        assert app.tiers[-2].id == last_id, "上移未改变奖项顺序"
        assert app.selected_tier_id == last_id, "上移后选中项发生变化"
        app._move_tier(1)
        assert [t.id for t in app.tiers] == original_ids, "下移未恢复原顺序"

    dialog = TierDialog(root, tiers=app.tiers)
    assert dialog.position_combo.current() == len(app.tiers), "新增奖项默认应追加到末尾"
    dialog.position_combo.current(0)
    dialog.name_var.set("特等奖")
    dialog._on_ok()
    assert dialog.position == 0 and dialog.result is not None, "未记录插入位置"

    lowest = app.tiers[-1]
    # 优先用人数 >= 2 的奖项，确保覆盖批量结果的网格排布
    batch_tier = lowest
    for tier in reversed(app.tiers):
        if tier.count >= 2:
            batch_tier = tier
            break
    batch_index = app.tiers.index(batch_tier)
    expected = min(batch_tier.count, len(app.names))
    app.selected_tier_id = batch_tier.id
    app._refresh_tiers()
    app._finish_draw(batch_tier, draw_all=True)
    root.update_idletasks()

    assert len(app.winners_tree.get_children()) == expected, "全部抽取未刷新名单"
    assert app.engine is not None and len(app.engine.winners) == expected
    assert len({w.entrant.id for w in app.engine.winners}) == expected
    row = app.winners_tree.item(app.winners_tree.get_children()[0], "values")
    assert row[0] in {str(w.number) for w in app.engine.winners}, "中奖名单未显示编号"
    stage_text = app.stage_name.get()
    assert "号 ·" in stage_text, "批量抽奖未按单人格式显示获奖者"
    for winner in app.engine.winners:
        assert f"{winner.number} 号 · {winner.name}" in stage_text, "舞台未显示获奖者姓名"
    assert "等 " not in stage_text, "舞台不应再把中奖者折叠为“等 N 人”"
    if expected > 1:
        assert app.stage_grid.place_info(), "批量结果未使用网格展示"
        assert len(app._stage_winner_labels) == expected, "网格未展示全部中奖者"
    if expected == batch_tier.count and batch_index > 0:
        assert (
            app.selected_tier_id == app.tiers[batch_index - 1].id
        ), "抽满后未自动切换奖项"

    # 小窗口布局：抽奖按钮、中奖名单和舞台结果都不能被裁掉
    root.deiconify()
    root.geometry("960x640")
    root.update()
    draw_bottom = (
        app.draw_button.winfo_rooty() - root.winfo_rooty() + app.draw_button.winfo_height()
    )
    winners_bottom = (
        app.winners_tree.winfo_rooty()
        - root.winfo_rooty()
        + app.winners_tree.winfo_height()
    )
    assert draw_bottom <= root.winfo_height(), "小窗口下抽奖按钮被裁掉"
    assert winners_bottom <= root.winfo_height(), "小窗口下中奖名单被裁掉"
    if expected > 1:
        body_top = app.stage_body.winfo_rooty()
        body_bottom = body_top + app.stage_body.winfo_height()
        grid_top = app.stage_grid.winfo_rooty()
        grid_bottom = grid_top + app.stage_grid.winfo_height()
        assert body_top - 2 <= grid_top and grid_bottom <= body_bottom + 2, (
            "小窗口下中奖者网格被裁切"
        )
    else:
        assert (
            app.stage_label.winfo_height() >= app.stage_label.winfo_reqheight()
        ), "小窗口下中奖姓名被裁切"
    assert (
        app.stage_sub_label.winfo_height() >= app.stage_sub_label.winfo_reqheight()
    ), "小窗口下抽奖结果被裁切"

    # 中奖者较多时自动分列，全部显示且不出现滚动条
    long_winners = [
        Winner("test", "测试奖项", "测试奖品", Entrant(i, i + 1, f"测试{i}"))
        for i in range(30)
    ]
    app._set_stage_winners(long_winners)
    root.update()
    assert len(app._stage_winner_labels) == len(long_winners), "长名单未全部展示"
    assert app.stage_grid.place_info(), "长名单未使用网格展示"
    columns = {int(label.grid_info()["column"]) for label in app._stage_winner_labels}
    assert len(columns) > 1, "长名单未自动分列"
    body_top = app.stage_body.winfo_rooty()
    body_bottom = body_top + app.stage_body.winfo_height()
    grid_top = app.stage_grid.winfo_rooty()
    grid_bottom = grid_top + app.stage_grid.winfo_height()
    assert body_top - 2 <= grid_top and grid_bottom <= body_bottom + 2, (
        "长名单网格超出舞台"
    )
    assert not hasattr(app, "stage_scroll"), "舞台不应再使用滚动条"
    for winner in long_winners:
        assert f"{winner.number} 号 · {winner.name}" in app.stage_name.get()
    root.withdraw()

    dialog = ManualInputDialog(root, initial_names=["张三", "李四"])
    dialog.name_entry.insert(0, "张三")
    dialog._add_single()
    dialog.bulk_text.insert("1.0", "王五,赵六")
    dialog._add_bulk()
    assert dialog.names == ["张三", "李四", "张三", "王五", "赵六"], dialog.names
    dialog.tree.selection_set("0")
    dialog._remove_selected()
    dialog._on_ok()
    assert dialog.result == ["李四", "张三", "王五", "赵六"], dialog.result

    root.destroy()
    print("UI smoke test OK")


if __name__ == "__main__":
    main()
