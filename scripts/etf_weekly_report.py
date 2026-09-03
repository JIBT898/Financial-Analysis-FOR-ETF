#!/usr/bin/env python3
"""生成沪深 ETF 周度份额变化榜。

口径：周末份额 - 上周末份额；年内变化为周末份额 - 上年最后交易日份额。
数据源：上海证券交易所、深圳证券交易所公开基金份额数据。
"""

from __future__ import annotations

import argparse
import concurrent.futures
import html
import io
import json
import random
import re
import sys
import urllib.parse
import urllib.request
import warnings
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
)


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36"
)


@dataclass(frozen=True)
class Snapshot:
    trade_date: date
    rows: pd.DataFrame


def http_get(url: str, params: dict[str, str], referer: str) -> bytes:
    query = urllib.parse.urlencode(params)
    request = urllib.request.Request(
        f"{url}?{query}",
        headers={"User-Agent": USER_AGENT, "Referer": referer},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read()


def fetch_fund_full_names() -> dict[str, str]:
    """读取基金全称，并将法定长后缀压缩为用户熟悉的 ETF。"""
    request = urllib.request.Request(
        "https://fund.eastmoney.com/js/fundcode_search.js",
        headers={"User-Agent": USER_AGENT, "Referer": "https://fund.eastmoney.com/"},
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        text = response.read().decode("utf-8-sig", errors="replace")
    start = text.find("[")
    end = text.rfind("]")
    if start < 0 or end <= start:
        raise ValueError("基金名称列表格式异常")
    records = json.loads(text[start : end + 1])
    suffixes = (
        "交易型开放式指数证券投资基金",
        "交易型开放式证券投资基金",
        "交易型开放式指数基金",
        "交易型开放式基金",
    )
    result: dict[str, str] = {}
    for item in records:
        if len(item) < 3:
            continue
        code = str(item[0]).zfill(6)
        full_name = str(item[2]).strip()
        for suffix in suffixes:
            full_name = full_name.replace(suffix, "ETF")
        result[code] = full_name
    return result


def compact_legal_name(full_name: str) -> str:
    suffixes = (
        "交易型开放式指数证券投资基金",
        "交易型开放式证券投资基金",
        "交易型开放式指数基金",
        "交易型开放式基金",
    )
    for suffix in suffixes:
        full_name = full_name.replace(suffix, "ETF")
    return full_name.strip()


def fetch_fund_legal_name(code: str) -> tuple[str, str]:
    request = urllib.request.Request(
        f"https://fundf10.eastmoney.com/jbgk_{code}.html",
        headers={"User-Agent": USER_AGENT, "Referer": "https://fund.eastmoney.com/"},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        page = response.read().decode("utf-8-sig", errors="replace")
    match = re.search(r"基金全称</th><td[^>]*>(.*?)</td>", page, flags=re.S)
    if not match:
        raise ValueError(f"{code}未找到基金全称")
    full_name = html.unescape(re.sub(r"<[^>]+>", "", match.group(1))).strip()
    compact_name = compact_legal_name(full_name)
    manager_match = re.search(
        r"基金管理人</th><td><a[^>]*>(.*?)</a>", page, flags=re.S
    )
    if manager_match:
        manager = html.unescape(re.sub(r"<[^>]+>", "", manager_match.group(1))).strip()
        for suffix in ("基金管理股份有限公司", "基金管理有限公司", "基金股份有限公司", "基金"):
            if manager.endswith(suffix):
                manager = manager[: -len(suffix)]
                break
        if manager and manager not in compact_name:
            compact_name = manager + compact_name
    return code, compact_name


def enrich_legal_names(frame: pd.DataFrame, indices: list[int]) -> pd.DataFrame:
    cache_path = Path(__file__).with_name("etf_full_names_cache.json")
    try:
        cache_payload = json.loads(cache_path.read_text(encoding="utf-8"))
        cache = cache_payload.get("names", {}) if cache_payload.get("version") == 2 else {}
    except (FileNotFoundError, json.JSONDecodeError):
        cache = {}

    codes = sorted(set(frame.loc[indices, "code"].astype(str)))
    missing = [code for code in codes if not cache.get(code)]
    if missing:
        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
            futures = {executor.submit(fetch_fund_legal_name, code): code for code in missing}
            for future in concurrent.futures.as_completed(futures):
                code = futures[future]
                try:
                    _, name = future.result()
                    cache[code] = name
                except Exception as exc:  # noqa: BLE001 - 单只失败时回退为简称
                    print(f"提示：{code}基金全称获取失败，保留简称：{exc}", file=sys.stderr)
        cache_path.write_text(
            json.dumps({"version": 2, "names": cache}, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    result = frame.copy()
    mapped = result.loc[indices, "code"].map(cache)
    result.loc[indices, "name"] = mapped.where(
        mapped.notna() & mapped.ne(""), result.loc[indices, "name"]
    )
    return result


def fetch_split_history(code: str) -> tuple[str, list[tuple[date, float, str]]]:
    request = urllib.request.Request(
        f"https://fundf10.eastmoney.com/fhsp_{code}.html",
        headers={"User-Agent": USER_AGENT, "Referer": "https://fund.eastmoney.com/"},
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        page = response.read().decode("utf-8-sig", errors="replace")
    events: list[tuple[date, float, str]] = []
    row_pattern = re.compile(
        r"<tr><td>\d{4}年</td><td>(\d{4}-\d{2}-\d{2})</td>"
        r"<td>(.*?)</td><td>([\d.]+):([\d.]+)</td>",
        flags=re.S,
    )
    for event_day, event_type, left, right in row_pattern.findall(page):
        left_value = float(left)
        right_value = float(right)
        if left_value > 0 and right_value > 0:
            events.append(
                (
                    datetime.strptime(event_day, "%Y-%m-%d").date(),
                    right_value / left_value,
                    html.unescape(re.sub(r"<[^>]+>", "", event_type)).strip(),
                )
            )
    return code, events


def apply_split_rules(
    frame: pd.DataFrame,
    candidate_indices: list[int],
    prior_day: date,
    latest_day: date,
) -> tuple[pd.DataFrame, list[str]]:
    """剔除当周发生拆分的ETF，避免机械份额变化被当作申赎。"""
    result = frame.copy()
    codes = sorted(set(result.loc[candidate_indices, "code"].astype(str)))
    histories: dict[str, list[tuple[date, float, str]]] = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(fetch_split_history, code): code for code in codes}
        for future in concurrent.futures.as_completed(futures):
            code = futures[future]
            try:
                _, histories[code] = future.result()
            except Exception as exc:  # noqa: BLE001 - 获取失败不擅自认定存在拆分
                print(f"提示：{code}拆分记录获取失败：{exc}", file=sys.stderr)
                histories[code] = []

    excluded_indices: list[int] = []
    excluded_notes: list[str] = []
    for index in candidate_indices:
        code = str(result.at[index, "code"])
        events = histories.get(code, [])
        weekly_events = [event for event in events if prior_day < event[0] <= latest_day]
        if weekly_events:
            excluded_indices.append(index)
            details = "、".join(
                f"{event_day:%Y-%m-%d} {event_type}1:{factor:g}"
                for event_day, factor, event_type in weekly_events
            )
            excluded_notes.append(f"{result.at[index, 'key']} {result.at[index, 'name']}（{details}）")
            continue

    if excluded_indices:
        result = result.drop(index=excluded_indices)
    return result, excluded_notes


def fetch_sse(day: date) -> pd.DataFrame:
    raw = http_get(
        "https://query.sse.com.cn/commonQuery.do",
        {
            "isPagination": "true",
            "pageHelp.pageSize": "10000",
            "pageHelp.pageNo": "1",
            "pageHelp.beginPage": "1",
            "pageHelp.cacheSize": "1",
            "pageHelp.endPage": "1",
            "sqlId": "COMMON_SSE_ZQPZ_ETFZL_XXPL_ETFGM_SEARCH_L",
            "STAT_DATE": day.isoformat(),
        },
        "https://www.sse.com.cn/assortment/fund/etf/list/scale/",
    )
    payload = json.loads(raw.decode("utf-8"))
    rows = payload.get("result") or []
    if not rows:
        return pd.DataFrame(columns=["code", "name", "shares", "exchange", "type"])
    frame = pd.DataFrame(rows)
    result = pd.DataFrame(
        {
            "code": frame["SEC_CODE"].astype(str).str.zfill(6),
            "name": frame["SEC_NAME"].astype(str),
            # 上交所接口单位为万份。
            "shares": pd.to_numeric(frame["TOT_VOL"], errors="coerce") * 10_000,
            "exchange": "SH",
            "type": frame.get("ETF_TYPE", ""),
        }
    )
    return result.dropna(subset=["shares"])


def fetch_szse(day: date) -> pd.DataFrame:
    raw = http_get(
        "https://www.szse.cn/api/report/ShowReport",
        {
            "SHOWTYPE": "xlsx",
            "CATALOGID": "scsj_fund_jjgm",
            "TABKEY": "tab1",
            "txtStart": day.isoformat(),
            "txtEnd": day.isoformat(),
            "jjlb": "ETF",
            "random": str(random.random()),
        },
        "https://www.szse.cn/market/fund/volume/etf/index.html",
    )
    frame = pd.read_excel(io.BytesIO(raw), engine="openpyxl").dropna(how="all")
    if frame.empty or "基金代码" not in frame.columns:
        return pd.DataFrame(columns=["code", "name", "shares", "exchange", "type"])
    codes = pd.to_numeric(frame["基金代码"], errors="coerce")
    frame = frame[codes.notna()].copy()
    frame["code"] = codes[codes.notna()].astype(int).astype(str).str.zfill(6)
    shares_col = "基金规模(份)" if "基金规模(份)" in frame.columns else "基金份额"
    result = pd.DataFrame(
        {
            "code": frame["code"],
            "name": frame["基金简称"].astype(str),
            "shares": pd.to_numeric(
                frame[shares_col].astype(str).str.replace(",", "", regex=False),
                errors="coerce",
            ),
            "exchange": "SZ",
            "type": "",
        }
    )
    return result.dropna(subset=["shares"])


def fetch_snapshot(target: date, max_lookback: int = 10) -> Snapshot:
    errors: list[str] = []
    for offset in range(max_lookback + 1):
        day = target - timedelta(days=offset)
        if day.weekday() >= 5:
            continue
        try:
            sse = fetch_sse(day)
            szse = fetch_szse(day)
            if not sse.empty and not szse.empty:
                rows = pd.concat([sse, szse], ignore_index=True)
                rows["key"] = rows["code"] + "." + rows["exchange"]
                return Snapshot(day, rows)
            errors.append(f"{day}: SSE={len(sse)}, SZSE={len(szse)}")
        except Exception as exc:  # noqa: BLE001 - 保留逐日回退能力
            errors.append(f"{day}: {type(exc).__name__}: {exc}")
    raise RuntimeError("未取得可用的沪深两市快照；" + " | ".join(errors))


def previous_friday(day: date) -> date:
    return day - timedelta(days=(day.weekday() - 4) % 7)


def previous_year_end(day: date) -> date:
    return date(day.year - 1, 12, 31)


def combine(latest: Snapshot, prior: Snapshot, year_start: Snapshot) -> pd.DataFrame:
    current = latest.rows[["key", "code", "name", "exchange", "type", "shares"]].copy()
    current = current.rename(columns={"shares": "latest_shares"})
    before = prior.rows[["key", "shares"]].rename(columns={"shares": "prior_shares"})
    beginning = year_start.rows[["key", "shares"]].rename(columns={"shares": "start_shares"})
    merged = current.merge(before, on="key", how="inner").merge(beginning, on="key", how="left")
    merged["weekly_change"] = merged["latest_shares"] - merged["prior_shares"]
    merged["ytd_change"] = merged["latest_shares"] - merged["start_shares"]
    for column in ["latest_shares", "weekly_change", "ytd_change"]:
        merged[column + "_yi"] = merged[column] / 100_000_000
    try:
        full_names = fetch_fund_full_names()
        mapped = merged["code"].map(full_names)
        merged["name"] = mapped.where(mapped.notna() & mapped.ne(""), merged["name"])
    except Exception as exc:  # noqa: BLE001 - 名称增强失败时仍可输出官方简称
        print(f"提示：基金全称映射失败，回退为交易所简称：{exc}", file=sys.stderr)
    return merged


def likely_equity_etf(frame: pd.DataFrame) -> pd.DataFrame:
    """排除货币、债券、商品和REIT类，保留境内外权益ETF。"""
    excluded_words = (
        "货币",
        "现金",
        "添利",
        "保证金",
        "短融",
        "债",
        "同业存单",
        "黄金ETF",
        "黄金9999",
        "豆粕",
        "有色期货",
        "能源化工",
        "REIT",
    )
    names = frame["name"].fillna("")
    mask = ~names.str.contains("|".join(excluded_words), case=False, regex=True)
    # 上交所类型字段可直接识别债券、货币和商品类。
    types = frame["type"].fillna("")
    mask &= ~types.str.contains("债券|货币|商品", regex=True)
    return frame[mask].copy()


def format_table(rows: pd.DataFrame, positive: bool) -> str:
    ordered = rows.sort_values("weekly_change", ascending=not positive).head(10)
    lines = [
        "| 排名 | 代码 | ETF | 本周净申赎（亿份） | 年内份额变化*（亿份） | 最新总份额（亿份） |",
        "|---:|---|---|---:|---:|---:|",
    ]
    for rank, (_, item) in enumerate(ordered.iterrows(), 1):
        lines.append(
            f"| {rank} | {item['key']} | {item['name']} | "
            f"{item['weekly_change_yi']:.2f} | {item['ytd_change_yi']:.2f} | "
            f"{item['latest_shares_yi']:.2f} |"
        )
    return "\n".join(lines)


def prepare_rankings(
    frame: pd.DataFrame, latest: Snapshot, prior: Snapshot
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    equity = likely_equity_etf(frame)
    initial_inflow = equity[equity["weekly_change"] > 0].sort_values(
        "weekly_change", ascending=False
    ).head(15)
    initial_outflow = equity[equity["weekly_change"] < 0].sort_values(
        "weekly_change", ascending=True
    ).head(15)
    raw_candidate_indices = list(initial_inflow.index) + list(initial_outflow.index)
    relative_change = (
        equity.loc[raw_candidate_indices, "weekly_change"].abs()
        / equity.loc[raw_candidate_indices, "prior_shares"].replace(0, pd.NA)
    )
    # 份额拆分通常产生显著的机械变化，只对异常变动项查询拆分记录，减少请求并避免限频。
    candidate_indices = list(relative_change[relative_change >= 0.15].index)
    equity, split_notes = apply_split_rules(
        equity,
        candidate_indices,
        prior.trade_date,
        latest.trade_date,
    )
    inflow = equity[equity["weekly_change"] > 0]
    outflow = equity[equity["weekly_change"] < 0]
    inflow_top = inflow.sort_values("weekly_change", ascending=False).head(10)
    outflow_top = outflow.sort_values("weekly_change", ascending=True).head(10)
    top_indices = list(inflow_top.index) + list(outflow_top.index)
    equity = enrich_legal_names(equity, top_indices)
    inflow = equity[equity["weekly_change"] > 0]
    outflow = equity[equity["weekly_change"] < 0]
    inflow_top = inflow.sort_values("weekly_change", ascending=False).head(10)
    outflow_top = outflow.sort_values("weekly_change", ascending=True).head(10)
    return inflow_top, outflow_top, split_notes


def build_report(
    inflow_top: pd.DataFrame,
    outflow_top: pd.DataFrame,
    split_notes: list[str],
    latest: Snapshot,
    prior: Snapshot,
    year_start: Snapshot,
) -> str:
    inflow_names = "、".join(inflow_top.head(3)["name"].tolist())
    outflow_names = "、".join(outflow_top.head(3)["name"].tolist())
    return "\n".join(
        [
            f"# ETF周度申赎榜（截至{latest.trade_date:%Y-%m-%d}）",
            "",
            f"> 口径：ETF总份额变化；本周为{prior.trade_date:%Y-%m-%d}至"
            f"{latest.trade_date:%Y-%m-%d}，年内基准为{year_start.trade_date:%Y-%m-%d}。"
            "单位：亿份。范围：沪深两市权益ETF，排除货币、债券、商品和REIT。",
            "",
            "## 净申购前10",
            "",
            format_table(inflow_top, positive=True),
            "",
            "## 净赎回前10",
            "",
            format_table(outflow_top, positive=False),
            "",
            "## 一句话观察",
            "",
            f"净申购主要集中在{inflow_names}；净赎回主要集中在{outflow_names}。",
            "",
            *(
                ["拆分校正：已剔除" + "；".join(split_notes) + "。", ""]
                if split_notes
                else []
            ),
            "说明：这里的“申赎”指ETF总份额增减，不是二级市场成交额或按价格折算的资金净流入。"
            "当周发生份额拆分或折算的ETF会被剔除。",
            "*年内份额变化沿用参考图片口径，为当前份额减上年最后交易日份额，可能包含年内拆分影响。",
            "数据源：[上海证券交易所](https://www.sse.com.cn/assortment/fund/etf/list/scale/)；"
            "[深圳证券交易所](https://www.szse.cn/market/fund/volume/etf/index.html)。",
        ]
    )


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        Path("C:/Windows/Fonts/msyhbd.ttc" if bold else "C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def fit_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.FreeTypeFont,
    max_width: int,
    max_lines: int = 2,
) -> list[str]:
    """按像素宽度切分中文名称，最多两行，必要时用省略号。"""
    lines: list[str] = []
    remaining = text
    while remaining and len(lines) < max_lines:
        line = ""
        for char in remaining:
            candidate = line + char
            if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
                line = candidate
            else:
                break
        if not line:
            line = remaining[0]
        lines.append(line)
        remaining = remaining[len(line) :]
    if remaining:
        last = lines[-1]
        while last and draw.textbbox((0, 0), last + "…", font=font)[2] > max_width:
            last = last[:-1]
        lines[-1] = last + "…"
    return lines


def render_mobile_image(
    inflow_top: pd.DataFrame,
    outflow_top: pd.DataFrame,
    split_notes: list[str],
    latest: Snapshot,
    prior: Snapshot,
    year_start: Snapshot,
    output_path: Path,
) -> None:
    """生成适合手机查看的A股风格竖版长图。"""
    width = 1080
    margin = 42
    section_header_h = 74
    column_header_h = 58
    row_h = 106
    section_gap = 28
    top_h = 252
    footer_h = 230 + (48 if split_notes else 0)
    section_h = section_header_h + column_header_h + row_h * 10
    height = top_h + section_h * 2 + section_gap + footer_h + 40

    colors = {
        "bg": "#F4F5F7",
        "card": "#FFFFFF",
        "ink": "#171A21",
        "muted": "#747B87",
        "line": "#E8E9ED",
        "red": "#D8292F",
        "red_dark": "#AA151B",
        "green": "#16845B",
        "gold": "#C69A43",
        "soft_red": "#FFF1F1",
        "soft_green": "#EDF8F3",
    }
    image = Image.new("RGB", (width, height), colors["bg"])
    draw = ImageDraw.Draw(image)
    title_font = load_font(52, bold=True)
    badge_font = load_font(23, bold=True)
    sub_font = load_font(25)
    section_font = load_font(32, bold=True)
    column_font = load_font(22, bold=True)
    rank_font = load_font(22, bold=True)
    name_font = load_font(25, bold=True)
    code_font = load_font(20)
    number_font = load_font(26, bold=True)
    small_number_font = load_font(23)
    footer_font = load_font(21)

    # 顶部信息区
    draw.rounded_rectangle((margin, 36, width - margin, 210), radius=26, fill=colors["card"])
    draw.rectangle((margin, 36, margin + 12, 210), fill=colors["red"])
    draw.text((margin + 38, 58), "ETF周度份额变化榜", font=title_font, fill=colors["ink"])
    date_label = f"截至 {latest.trade_date:%Y.%m.%d}"
    date_w = draw.textbbox((0, 0), date_label, font=badge_font)[2]
    badge_box = (width - margin - date_w - 34, 60, width - margin - 12, 106)
    draw.rounded_rectangle(badge_box, radius=20, fill=colors["red"])
    draw.text((badge_box[0] + 17, badge_box[1] + 8), date_label, font=badge_font, fill="white")
    draw.text(
        (margin + 38, 145),
        f"沪深权益ETF｜{prior.trade_date:%m.%d}—{latest.trade_date:%m.%d}｜单位：亿份",
        font=sub_font,
        fill=colors["muted"],
    )

    def draw_section(y: int, title: str, rows: pd.DataFrame, positive: bool) -> int:
        accent = colors["red"] if positive else colors["green"]
        soft = colors["soft_red"] if positive else colors["soft_green"]
        box = (margin, y, width - margin, y + section_h)
        draw.rounded_rectangle(box, radius=24, fill=colors["card"])
        draw.rounded_rectangle(
            (margin, y, width - margin, y + section_header_h),
            radius=24,
            fill=accent,
        )
        draw.rectangle(
            (margin, y + section_header_h - 24, width - margin, y + section_header_h),
            fill=accent,
        )
        draw.text((margin + 28, y + 17), title, font=section_font, fill="white")
        draw.text((margin + 730, y + 23), "按份额变化排名", font=column_font, fill="#FFE9E9" if positive else "#D9F1E8")

        header_y = y + section_header_h
        draw.rectangle((margin, header_y, width - margin, header_y + column_header_h), fill=soft)
        draw.text((margin + 30, header_y + 15), "ETF", font=column_font, fill=colors["muted"])
        centers = [700, 838, 970]
        for center, label in zip(centers, ["本周", "年内*", "总份额"]):
            label_w = draw.textbbox((0, 0), label, font=column_font)[2]
            draw.text((center - label_w / 2, header_y + 15), label, font=column_font, fill=colors["muted"])

        ordered = rows.sort_values("weekly_change", ascending=not positive).head(10)
        for rank, (_, item) in enumerate(ordered.iterrows(), 1):
            row_y = header_y + column_header_h + (rank - 1) * row_h
            if rank > 1:
                draw.line((margin + 22, row_y, width - margin - 22, row_y), fill=colors["line"], width=1)
            circle = (margin + 22, row_y + 34, margin + 62, row_y + 74)
            draw.ellipse(circle, fill=accent if rank <= 3 else "#ECEEF2")
            rank_text = str(rank)
            rank_box = draw.textbbox((0, 0), rank_text, font=rank_font)
            draw.text(
                ((circle[0] + circle[2] - rank_box[2]) / 2, row_y + 40),
                rank_text,
                font=rank_font,
                fill="white" if rank <= 3 else colors["muted"],
            )
            name_x = margin + 82
            name_lines = fit_text(draw, str(item["name"]), name_font, 515, max_lines=2)
            name_y = row_y + (17 if len(name_lines) == 1 else 8)
            for line_no, line in enumerate(name_lines):
                draw.text((name_x, name_y + line_no * 30), line, font=name_font, fill=colors["ink"])
            draw.text((name_x, row_y + 76), str(item["key"]), font=code_font, fill=colors["muted"])

            values = [
                (float(item["weekly_change_yi"]), number_font, accent),
                (float(item["ytd_change_yi"]), small_number_font, colors["ink"]),
                (float(item["latest_shares_yi"]), small_number_font, colors["ink"]),
            ]
            for center, (value, font, fill) in zip(centers, values):
                label = f"{value:.2f}"
                label_w = draw.textbbox((0, 0), label, font=font)[2]
                draw.text((center - label_w / 2, row_y + 38), label, font=font, fill=fill)
        return y + section_h

    y = top_h
    y = draw_section(y, "净申购 TOP 10", inflow_top, positive=True)
    y += section_gap
    y = draw_section(y, "净赎回 TOP 10", outflow_top, positive=False)

    footer_y = y + 24
    draw.text((margin + 8, footer_y), "数据口径", font=column_font, fill=colors["ink"])
    footer_y += 38
    notes = [
        "本周＝最新交易日总份额－上周同期总份额；不是二级市场成交额或资金净流入。",
        f"年内*＝最新总份额－{year_start.trade_date:%Y.%m.%d}总份额，可能包含年内拆分影响。",
        "当周发生份额拆分或折算的ETF已从排名中剔除。",
    ]
    if split_notes:
        notes.append("本期剔除：" + "；".join(split_notes))
    notes.append("数据源：上海证券交易所、深圳证券交易所公开ETF份额数据")
    for note in notes:
        lines = fit_text(draw, "• " + note, footer_font, width - margin * 2 - 20, max_lines=2)
        for line in lines:
            draw.text((margin + 8, footer_y), line, font=footer_font, fill=colors["muted"])
            footer_y += 29
        footer_y += 6

    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path, format="PNG", optimize=True)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("--as-of", help="目标周五，格式 YYYY-MM-DD；默认取当前日期之前最近周五")
    parser.add_argument("--output", help="可选：将Markdown报告写入指定路径")
    parser.add_argument("--image", help="可选：将手机长图写入指定路径；默认写入reports目录")
    parser.add_argument("--json", action="store_true", help="同时输出用于复核的JSON明细")
    args = parser.parse_args()

    requested = (
        datetime.strptime(args.as_of, "%Y-%m-%d").date()
        if args.as_of
        else previous_friday(date.today())
    )
    latest = fetch_snapshot(requested)
    prior = fetch_snapshot(latest.trade_date - timedelta(days=7))
    year_start = fetch_snapshot(previous_year_end(latest.trade_date))
    combined = combine(latest, prior, year_start)
    inflow_top, outflow_top, split_notes = prepare_rankings(combined, latest, prior)
    report = build_report(inflow_top, outflow_top, split_notes, latest, prior, year_start)
    image_path = (
        Path(args.image)
        if args.image
        else Path.cwd()
        / "reports"
        / f"ETF周度份额变化榜_{latest.trade_date:%Y-%m-%d}.png"
    )
    render_mobile_image(
        inflow_top,
        outflow_top,
        split_notes,
        latest,
        prior,
        year_start,
        image_path,
    )
    print(report)
    print(f"\n图片已生成：{image_path.resolve()}")

    if args.json:
        records = combined[
            ["key", "name", "type", "weekly_change_yi", "ytd_change_yi", "latest_shares_yi"]
        ].to_dict(orient="records")
        print("\n<!-- JSON_REVIEW_DATA\n" + json.dumps(records, ensure_ascii=False) + "\n-->")

    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(report + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
