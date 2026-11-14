"""纯逻辑抽奖引擎：不依赖界面，可单元测试。"""

from __future__ import annotations

import random
from collections.abc import Iterable, Sequence

from .models import Entrant, PrizeTier, Winner


class LotteryError(Exception):
    """抽奖无法进行时抛出，消息可直接展示给用户。"""


class LotteryEngine:
    """管理名单、奖项与中奖记录，保证一个人最多中奖一次。"""

    def __init__(
        self,
        entrants: Iterable[Entrant],
        tiers: Sequence[PrizeTier],
        rng: random.Random | None = None,
    ) -> None:
        self._entrants = list(entrants)
        self._tiers = list(tiers)
        self._rng = rng or random.Random()
        self._winners: list[Winner] = []
        # 名单内部按下标去重，避免同一 Entrant 对象重复出现时被抽中两次
        if len({e.id for e in self._entrants}) != len(self._entrants):
            raise ValueError("名单中存在重复的参与者 id")

    @property
    def entrants(self) -> list[Entrant]:
        return list(self._entrants)

    @property
    def tiers(self) -> list[PrizeTier]:
        return list(self._tiers)

    @property
    def winners(self) -> list[Winner]:
        return list(self._winners)

    def tier(self, tier_id: str) -> PrizeTier:
        for tier in self._tiers:
            if tier.id == tier_id:
                return tier
        raise LotteryError("奖项不存在，可能已被删除")

    def winners_for(self, tier_id: str) -> list[Winner]:
        return [w for w in self._winners if w.tier_id == tier_id]

    def is_full(self, tier_id: str) -> bool:
        tier = self.tier(tier_id)
        return len(self.winners_for(tier_id)) >= tier.count

    @property
    def remaining_entrants(self) -> list[Entrant]:
        won = {w.entrant.id for w in self._winners}
        return [e for e in self._entrants if e.id not in won]

    @property
    def remaining_count(self) -> int:
        return len(self._entrants) - len(self._winners)

    def draw(self, tier_id: str) -> Winner:
        tier = self.tier(tier_id)
        if len(self.winners_for(tier_id)) >= tier.count:
            raise LotteryError(f"「{tier.name}」已抽满 {tier.count} 人")
        pool = self.remaining_entrants
        if not pool:
            raise LotteryError("所有人都已中奖，没有可抽取的名单了")
        entrant = self._rng.choice(pool)
        winner = Winner(
            tier_id=tier.id,
            tier_name=tier.name,
            prize=tier.prize,
            entrant=entrant,
        )
        self._winners.append(winner)
        return winner

    def draw_all(self, tier_id: str) -> list[Winner]:
        """抽出该奖项剩余的全部名额；名单不足时抽出所有剩余人员。"""
        tier = self.tier(tier_id)
        slots = tier.count - len(self.winners_for(tier_id))
        if slots <= 0:
            raise LotteryError(f"「{tier.name}」已抽满 {tier.count} 人")
        if not self.remaining_entrants:
            raise LotteryError("所有人都已中奖，没有可抽取的名单了")
        winners: list[Winner] = []
        for _ in range(min(slots, self.remaining_count)):
            winners.append(self.draw(tier_id))
        return winners

    def reset(self) -> None:
        """清空中奖记录，名单和奖项保持不变。"""
        self._winners.clear()
