"""CSV読み込み・グルーピングユーティリティ"""

import pandas as pd
from pathlib import Path


def load_and_group(csv_path: Path, group_key: str) -> dict:
    """CSVを読み込み、group_keyでグルーピングしてdict返却"""
    df = pd.read_csv(csv_path)
    groups = {}
    for key, group_df in df.groupby(group_key):
        groups[key] = group_df.reset_index(drop=True)
    return groups


def detect_csv_files(input_dir: Path, patterns: dict) -> dict:
    """input/ディレクトリからCSVをパターンで自動検出

    Args:
        input_dir: 検索対象ディレクトリ
        patterns: {キー名: [マッチするキーワードリスト]} の辞書

    Returns:
        {キー名: Path} の辞書（見つからなければNone）
    """
    files = {k: None for k in patterns}

    for f in input_dir.glob("*.csv"):
        name_lower = f.name.lower()
        for key, keywords in patterns.items():
            if files[key] is not None:
                continue
            if any(kw in f.name or kw in name_lower for kw in keywords):
                files[key] = f

    return files
