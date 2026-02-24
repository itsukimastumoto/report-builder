"""matplotlib グラフ画像生成モジュール"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from pathlib import Path
from datetime import datetime

# 日本語フォント設定（macOS）
plt.rcParams["font.family"] = "Hiragino Sans"
plt.rcParams["axes.unicode_minus"] = False

# カラーパレット（テンプレートに合わせた配色）
COLORS = [
    "#4285F4",  # 青（メイン）
    "#EA4335",  # 赤
    "#34A853",  # 緑
    "#5F6368",  # グレー
    "#00BCD4",  # シアン
    "#FF9800",  # オレンジ
    "#9C27B0",  # 紫
    "#795548",  # 茶
]


def generate_bar_chart(
    dates: list,
    values: list,
    title: str,
    ylabel: str,
    output_path: Path,
    figsize: tuple = (5.77, 2.4),
    color: str = COLORS[0],
    bar_width: float = 0.8,
):
    """日別推移の棒グラフを生成

    Args:
        dates: 日付リスト（文字列 or datetime）
        values: 数値リスト
        title: グラフタイトル
        ylabel: Y軸ラベル
        output_path: 出力画像パス
        figsize: 図のサイズ（インチ）
        color: 棒の色
    """
    fig, ax = plt.subplots(figsize=figsize)
    fig.subplots_adjust(left=0.08, right=0.97, top=0.88, bottom=0.18)

    if dates and isinstance(dates[0], str):
        dates = [datetime.strptime(d, "%Y-%m-%d") for d in dates]

    bars = ax.bar(dates, values, color=color, width=bar_width)
    ax.set_title(title, fontsize=9, pad=4, fontweight="bold")
    ax.set_ylabel(ylabel, fontsize=7, labelpad=2)

    # 棒グラフの上にデータラベルを表示
    for bar, val in zip(bars, values):
        if val > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height(),
                f"{int(val):,}",
                ha="center", va="bottom",
                fontsize=5, color="#333333",
            )

    if bar_width > 1:
        # 週次データ: "M/D週" 表示、各データポイントにティック
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m/%-d週"))
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
    else:
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m/%-d"))
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
    ax.tick_params(axis="both", labelsize=6, pad=2)
    plt.xticks(rotation=0)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", alpha=0.3)

    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def generate_donut_chart(
    values: list,
    labels: list,
    title: str,
    output_path: Path,
    figsize: tuple = (3.0, 2.0),
    colors: list = None,
):
    """ドーナツグラフを1つ生成

    Args:
        values: 数値リスト
        labels: ラベルリスト
        title: グラフタイトル
        output_path: 出力画像パス
        figsize: 図のサイズ
        colors: 色リスト（Noneならデフォルト）
    """
    if colors is None:
        colors = COLORS[: len(values)]

    fig, ax = plt.subplots(figsize=figsize)

    total = sum(values)
    percentages = [(v / total * 100) if total > 0 else 0 for v in values]

    wedges, texts, autotexts = ax.pie(
        values,
        labels=None,
        autopct=lambda p: f"{p:.1f}%" if p >= 4 else "",
        colors=colors,
        startangle=90,
        pctdistance=0.75,
        wedgeprops=dict(width=0.4),
    )

    for t in autotexts:
        t.set_fontsize(7)

    ax.set_title(title, fontsize=9, pad=5)

    ax.legend(
        labels,
        loc="center left",
        bbox_to_anchor=(1, 0.5),
        fontsize=6,
        frameon=False,
    )

    plt.tight_layout()
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def generate_double_donut_chart(
    values1: list,
    labels1: list,
    title1: str,
    values2: list,
    labels2: list,
    title2: str,
    output_path: Path,
    figsize: tuple = (6.2, 2.0),
    colors: list = None,
):
    """横並びのダブルドーナツグラフを生成（テンプレートの「発行総数」「販売総額」に対応）

    Args:
        values1/labels1/title1: 左のドーナツ
        values2/labels2/title2: 右のドーナツ
        output_path: 出力画像パス
        figsize: 図のサイズ
        colors: 色リスト
    """
    if colors is None:
        colors = COLORS

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=figsize)
    fig.subplots_adjust(left=0.02, right=0.82, top=0.88, bottom=0.02, wspace=0.15)

    for ax, values, labels, title in [
        (ax1, values1, labels1, title1),
        (ax2, values2, labels2, title2),
    ]:
        c = colors[: len(values)]
        wedges, texts, autotexts = ax.pie(
            values,
            labels=None,
            autopct=lambda p: f"{p:.1f}%" if p >= 4 else "",
            colors=c,
            startangle=90,
            pctdistance=0.75,
            wedgeprops=dict(width=0.4),
        )
        for t in autotexts:
            t.set_fontsize(6)
        ax.set_title(title, fontsize=8, pad=2, fontweight="bold")

    # 凡例は右端に1つだけ
    all_labels = labels1 if len(labels1) >= len(labels2) else labels2
    all_colors = colors[: len(all_labels)]
    legend_patches = [
        plt.matplotlib.patches.Patch(color=c, label=l)
        for c, l in zip(all_colors, all_labels)
    ]
    fig.legend(
        handles=legend_patches,
        loc="center right",
        bbox_to_anchor=(0.99, 0.5),
        fontsize=6,
        frameon=False,
    )

    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
