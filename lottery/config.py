"""奖项配置持久化（config/tiers.json）。"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from .models import PrizeTier, new_tier_id


def _base_dir() -> Path:
    """打包成 exe 后，配置写在 exe 同目录，避免落在临时解包目录里。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE_DIR = _base_dir()
CONFIG_DIR = BASE_DIR / "config"
TIERS_FILE = CONFIG_DIR / "tiers.json"


def default_tiers() -> list[PrizeTier]:
    return [
        PrizeTier(name="一等奖", prize="一等奖奖品", count=1),
        PrizeTier(name="二等奖", prize="二等奖奖品", count=2),
        PrizeTier(name="三等奖", prize="三等奖奖品", count=3),
    ]


def load_tiers(path: Path = TIERS_FILE) -> list[PrizeTier]:
    """读取奖项配置；文件缺失或损坏时回退到默认奖项。"""
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default_tiers()
    if not isinstance(raw, list):
        return default_tiers()

    tiers: list[PrizeTier] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        tier_id = item.get("id")
        try:
            tiers.append(
                PrizeTier(
                    name=str(item.get("name", "")),
                    prize=str(item.get("prize", "")),
                    count=int(item.get("count", 0)),
                    id=tier_id if isinstance(tier_id, str) and tier_id else new_tier_id(),
                )
            )
        except (TypeError, ValueError):
            continue
    return tiers or default_tiers()


def save_tiers(tiers: list[PrizeTier], path: Path = TIERS_FILE) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = [
        {"id": t.id, "name": t.name, "prize": t.prize, "count": t.count} for t in tiers
    ]
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
