"""抽奖引擎单元测试（标准库 unittest，无需额外依赖）。"""

import random
import unittest

from lottery.engine import LotteryEngine, LotteryError
from lottery.models import Entrant, PrizeTier


def make_engine(count: int = 5, seed: int = 1) -> LotteryEngine:
    entrants = [Entrant(id=i, number=i + 1, name=f"选手{i}") for i in range(count)]
    tiers = [
        PrizeTier(id="first", name="一等奖", prize="奖品A", count=2),
        PrizeTier(id="second", name="二等奖", prize="奖品B", count=1),
    ]
    return LotteryEngine(entrants, tiers, rng=random.Random(seed))


class TestLotteryEngine(unittest.TestCase):
    def test_draw_records_winner(self):
        engine = make_engine()
        winner = engine.draw("first")
        self.assertEqual(winner.tier_id, "first")
        self.assertEqual(len(engine.winners), 1)
        self.assertIn(winner.entrant, engine.entrants)

    def test_no_repeat_winners(self):
        engine = make_engine(count=10)
        engine.draw("first")
        engine.draw("second")
        engine.draw("first")
        names = [w.entrant.id for w in engine.winners]
        self.assertEqual(len(names), len(set(names)))

    def test_tier_capacity_enforced(self):
        engine = make_engine()
        engine.draw("second")
        with self.assertRaises(LotteryError):
            engine.draw("second")

    def test_draw_all_fills_tier(self):
        engine = make_engine(count=10)
        winners = engine.draw_all("first")
        self.assertEqual(len(winners), 2)
        self.assertTrue(engine.is_full("first"))
        self.assertEqual(len({w.entrant.id for w in winners}), 2)

    def test_draw_all_partial_when_pool_exhausted(self):
        entrants = [Entrant(0, 1, "甲"), Entrant(1, 2, "乙")]
        tiers = [PrizeTier(id="t", name="奖", prize="奖", count=5)]
        engine = LotteryEngine(entrants, tiers, rng=random.Random(0))
        winners = engine.draw_all("t")
        self.assertEqual(len(winners), 2)
        self.assertEqual(engine.remaining_count, 0)
        self.assertFalse(engine.is_full("t"))

    def test_draw_all_full_tier_raises(self):
        engine = make_engine()
        engine.draw_all("second")
        with self.assertRaises(LotteryError):
            engine.draw_all("second")

    def test_draw_from_unknown_tier_raises(self):
        with self.assertRaises(LotteryError):
            make_engine().draw("missing")

    def test_exhausted_pool_raises(self):
        engine = make_engine(count=2)
        engine.draw("first")
        engine.draw("second")
        self.assertEqual(engine.remaining_count, 0)
        with self.assertRaises(LotteryError):
            engine.draw("first")

    def test_reset_keeps_entrants_and_tiers(self):
        engine = make_engine()
        engine.draw("first")
        engine.reset()
        self.assertEqual(engine.winners, [])
        self.assertEqual(engine.remaining_count, 5)
        self.assertEqual(len(engine.tiers), 2)

    def test_same_name_entrants_are_distinct(self):
        entrants = [Entrant(1, 1, "张三"), Entrant(2, 2, "张三")]
        tiers = [PrizeTier(id="t", name="奖", prize="奖", count=2)]
        engine = LotteryEngine(entrants, tiers, rng=random.Random(0))
        engine.draw("t")
        engine.draw("t")
        self.assertEqual(len({w.entrant.id for w in engine.winners}), 2)
        self.assertEqual({w.number for w in engine.winners}, {1, 2})

    def test_winner_exposes_number(self):
        engine = make_engine()
        winner = engine.draw("first")
        self.assertEqual(winner.number, winner.entrant.number)

    def test_duplicate_entrant_ids_rejected(self):
        entrants = [Entrant(1, 1, "A"), Entrant(1, 2, "B")]
        tiers = [PrizeTier(id="t", name="奖", prize="奖", count=1)]
        with self.assertRaises(ValueError):
            LotteryEngine(entrants, tiers)

    def test_seeded_draws_are_deterministic(self):
        def sequence(seed: int) -> list[int]:
            engine = make_engine(seed=seed)
            return [engine.draw("first").entrant.id for _ in range(2)]

        self.assertEqual(sequence(42), sequence(42))


if __name__ == "__main__":
    unittest.main()
