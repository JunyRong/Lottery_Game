"""数据模型：奖项、参与者、中奖记录。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from uuid import uuid4


def new_tier_id() -> str:
    return uuid4().hex[:12]


@dataclass
class PrizeTier:
    """一个奖项等级。"""

    name: str
    prize: str
    count: int
    id: str = field(default_factory=new_tier_id)

    def __post_init__(self) -> None:
        self.name = self.name.strip()
        self.prize = self.prize.strip()
        if not self.name:
            raise ValueError("奖项名称不能为空")
        if not self.prize:
            self.prize = self.name
        if not isinstance(self.count, int) or self.count < 1:
            raise ValueError("获奖人数必须是大于 0 的整数")


@dataclass(frozen=True)
class Entrant:
    """抽奖名单中的一个人。

    id 唯一（内部使用）；number 是展示编号；姓名允许重复。
    """

    id: int
    number: int
    name: str


@dataclass(frozen=True)
class Winner:
    tier_id: str
    tier_name: str
    prize: str
    entrant: Entrant
    drawn_at: datetime = field(default_factory=datetime.now)

    @property
    def name(self) -> str:
        return self.entrant.name

    @property
    def number(self) -> int:
        return self.entrant.number
