#!/usr/bin/env python3
"""バリデーションスコアリングのスモークテスト"""

import sys
import tempfile
from pathlib import Path

# scripts/ ディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent))

import pytest


class TestScoring:
    """calculate_score のスモークテスト"""

    def test_empty_dir_low_score(self):
        """空ディレクトリでスコアが低い（ファイルなし）"""
        with tempfile.TemporaryDirectory() as tmpdir:
            from validate import calculate_score
            result = calculate_score(Path(tmpdir), skip_diagrams=True)
            # ファイルが無いのでfiles=0, diagrams=0等だが一部デフォルト点が入る場合あり
            assert result["score"] < 50

    def test_scoring_breakdown_adds_up(self):
        """スコアの内訳合計がtotalと一致"""
        with tempfile.TemporaryDirectory() as tmpdir:
            from validate import calculate_score
            result = calculate_score(Path(tmpdir), skip_diagrams=True)
            total = sum(info["score"] for info in result["breakdown"].values())
            assert total == result["score"]

    def test_max_score_not_exceeded(self):
        """各カテゴリのスコアが上限を超えない"""
        with tempfile.TemporaryDirectory() as tmpdir:
            from validate import calculate_score
            result = calculate_score(Path(tmpdir), skip_diagrams=True)
            for cat, info in result["breakdown"].items():
                assert info["score"] <= info["max"], f"{cat}: {info['score']} > {info['max']}"
            assert result["score"] <= 100
