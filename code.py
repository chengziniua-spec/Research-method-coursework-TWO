from __future__ import annotations

import json
import math
import re
import textwrap
import zipfile
import importlib
import sys
from pathlib import Path
from typing import Callable
from xml.etree import ElementTree as ET

# Some teaching lab Python installs include matplotlib but omit its packaging
# dependency. Pip vendors a compatible copy, so reuse it when necessary.
try:
    import packaging.version  # noqa: F401
except ModuleNotFoundError:
    sys.modules["packaging"] = importlib.import_module("pip._vendor.packaging")
    sys.modules["packaging.version"] = importlib.import_module("pip._vendor.packaging.version")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import patches
from matplotlib.colors import LinearSegmentedColormap, Normalize
from matplotlib.ticker import FuncFormatter


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "search method file"
OUTPUT_DIR = ROOT / "visual_outputs"

XLS_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

METRIC_FIELDS = [
    "fce",
    "admissions",
    "male",
    "female",
    "gender_unknown",
    "emergency",
    "waiting_list",
    "planned",
    "other_admission",
    "age_0_14",
    "age_15_59",
    "age_60_74",
    "age_75_plus",
    "day_case",
    "bed_days",
]

CHAPTER_ORDER = [
    "Infectious and parasitic",
    "Neoplasms",
    "Blood and immune",
    "Endocrine and metabolic",
    "Mental health",
    "Nervous system",
    "Eye",
    "Ear",
    "Circulatory",
    "Respiratory",
    "Digestive",
    "Skin",
    "Musculoskeletal",
    "Genitourinary",
    "Pregnancy and childbirth",
    "Perinatal",
    "Congenital",
    "Symptoms and signs",
    "Injury and poisoning",
    "External causes",
    "Health services factors",
    "Special purpose codes",
    "Other or unmapped",
]

PALETTE = [
    "#264653",
    "#2a9d8f",
    "#e9c46a",
    "#f4a261",
    "#e76f51",
    "#457b9d",
    "#8d6ab8",
    "#5c946e",
    "#bc4749",
    "#6d597a",
    "#118ab2",
    "#ef476f",
]


def configure_style() -> None:
    plt.rcParams.update(
        {
            "figure.facecolor": "#fbfaf7",
            "axes.facecolor": "#fbfaf7",
            "savefig.facecolor": "#fbfaf7",
            "axes.edgecolor": "#ded8cf",
            "axes.labelcolor": "#2f3136",
            "xtick.color": "#4a4d52",
            "ytick.color": "#4a4d52",
            "text.color": "#202226",
            "font.family": "DejaVu Sans",
            "font.size": 10,
            "axes.titleweight": "bold",
            "axes.titlesize": 14,
            "axes.labelsize": 10,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "grid.color": "#e9e2d8",
            "grid.linewidth": 0.8,
        }
    )


def millions(x: float, _pos: int | None = None) -> str:
    if pd.isna(x):
        return ""
    if abs(x) >= 1_000_000:
        return f"{x / 1_000_000:.1f}m"
    if abs(x) >= 1_000:
        return f"{x / 1_000:.0f}k"
    return f"{x:.0f}"


def pct(x: float, _pos: int | None = None) -> str:
    if pd.isna(x):
        return ""
    return f"{x:.0%}"


def fiscal_label(start_year: int) -> str:
    return f"{start_year}-{str(start_year + 1)[-2:]}"


def clean_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


def clean_header(value: object) -> str:
    return clean_text(value).lower()


def to_number(value: object) -> float:
    text = clean_text(value)
    if not text or text in {".", "-", ".."}:
        return np.nan
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return np.nan


def wrap_label(text: str, width: int = 34) -> str:
    text = clean_text(text)
    return "\n".join(textwrap.wrap(text, width=width, break_long_words=False))


def short_category(row: pd.Series, width: int = 42) -> str:
    code = clean_text(row.get("code", ""))
    desc = clean_text(row.get("description", ""))
    label = f"{code} {desc}".strip()
    return wrap_label(label, width)


def col_to_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0
    total = 0
    for char in match.group(1):
        total = total * 26 + ord(char) - 64
    return total - 1


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall(f"{XLS_NS}si"):
        strings.append("".join(t.text or "" for t in item.iter(f"{XLS_NS}t")))
    return strings


def xlsx_sheet_paths(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels
            if rel.attrib.get("Type", "").endswith("/worksheet")
        }
        result = {}
        for sheet in workbook.findall(f"{XLS_NS}sheets/{XLS_NS}sheet"):
            name = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{REL_NS}id"]
            target = rel_targets[rel_id].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            result[name] = target
        return result


def read_xlsx_rows(path: Path, sheet_matcher: Callable[[str], bool]) -> list[list[object]]:
    sheet_paths = xlsx_sheet_paths(path)
    matches = [name for name in sheet_paths if sheet_matcher(name)]
    if not matches:
        raise ValueError(f"No matching sheet found in {path.name}")

    sheet_path = sheet_paths[matches[0]]
    rows: list[list[object]] = []

    with zipfile.ZipFile(path) as archive:
        strings = shared_strings(archive)
        root = ET.fromstring(archive.read(sheet_path))
        for row in root.findall(f".//{XLS_NS}sheetData/{XLS_NS}row"):
            values: list[object] = []
            for cell in row.findall(f"{XLS_NS}c"):
                idx = col_to_index(cell.attrib.get("r", ""))
                while len(values) <= idx:
                    values.append("")

                cell_type = cell.attrib.get("t")
                if cell_type == "s":
                    value_node = cell.find(f"{XLS_NS}v")
                    value = strings[int(value_node.text)] if value_node is not None and value_node.text else ""
                elif cell_type == "inlineStr":
                    value = "".join(t.text or "" for t in cell.iter(f"{XLS_NS}t"))
                else:
                    value_node = cell.find(f"{XLS_NS}v")
                    value = value_node.text if value_node is not None and value_node.text else ""
                values[idx] = value
            rows.append(values)
    return rows


def find_header_row(rows: list[list[object]]) -> int:
    for idx, row in enumerate(rows):
        lowered = [clean_header(value) for value in row]
        combined = " ".join(lowered)
        has_diagnosis = "diagnosis" in combined and (
            "summary" in combined or "table 2. diagnosis" in combined
        )
        has_metric = "finished consultant" in combined or "admissions" in combined
        if has_diagnosis and has_metric:
            return idx
    raise ValueError("Could not find a diagnosis summary header row.")


def find_index(headers: list[str], *needles: str, start: int = 0) -> int | None:
    for idx, header in enumerate(headers[start:], start=start):
        if header and all(needle in header for needle in needles):
            return idx
    return None


def build_header_map(headers: list[str]) -> dict[str, int | None]:
    def first(predicate: Callable[[str], bool], start: int = 0) -> int | None:
        for idx, header in enumerate(headers[start:], start=start):
            if header and predicate(header):
                return idx
        return None

    return {
        "fce": first(lambda h: "finished consultant episodes" in h),
        "admissions": first(lambda h: "finished admission episodes" in h or h == "admissions"),
        "male": first(lambda h: h.startswith("male")),
        "female": first(lambda h: h.startswith("female")),
        "gender_unknown": first(lambda h: "gender unknown" in h),
        "emergency": first(lambda h: h.startswith("emergency")),
        "waiting_list": first(lambda h: "waiting list" in h),
        "planned": first(lambda h: h.startswith("planned")),
        "other_admission": first(lambda h: "other admission" in h or h == "other"),
        "mean_wait": first(lambda h: "mean time waited" in h or "mean waiting time" in h),
        "median_wait": first(lambda h: "median time waited" in h or "median waiting time" in h),
        "mean_los": first(lambda h: "mean length" in h),
        "median_los": first(lambda h: "median length" in h),
        "mean_age": first(lambda h: "mean age" in h),
        "day_case": first(lambda h: h.startswith("day case")),
        "bed_days": first(lambda h: "bed day" in h or "fce bed days" in h),
    }


def parse_age_range(header: str) -> tuple[int, int] | None:
    header = clean_header(header)
    if not header.startswith("age ") or "mean age" in header:
        return None
    match = re.match(r"age\s+(\d+)(?:-(\d+)|\+)?", header)
    if not match:
        return None
    start = int(match.group(1))
    if header.endswith("+") or "+" in header:
        return start, 200
    end = int(match.group(2)) if match.group(2) else start
    return start, end


def broad_age_columns(headers: list[str]) -> dict[str, list[int]]:
    buckets = {
        "age_0_14": [],
        "age_15_59": [],
        "age_60_74": [],
        "age_75_plus": [],
    }

    direct = {
        "age 0-14": "age_0_14",
        "age 15-59": "age_15_59",
        "age 60-74": "age_60_74",
        "age 75+": "age_75_plus",
    }

    for idx, header in enumerate(headers):
        if header in direct:
            buckets[direct[header]].append(idx)
            continue

        age_range = parse_age_range(header)
        if not age_range:
            continue
        start, end = age_range
        if start >= 0 and end <= 14:
            buckets["age_0_14"].append(idx)
        elif start >= 15 and end <= 59:
            buckets["age_15_59"].append(idx)
        elif start >= 60 and end <= 74:
            buckets["age_60_74"].append(idx)
        elif start >= 75:
            buckets["age_75_plus"].append(idx)
    return buckets


def clean_code(value: object) -> str:
    text = clean_text(value)
    text = re.sub(r"^[^A-Z0-9]+", "", text)
    text = text.replace(";", ",")
    if text.lower().startswith("total"):
        return "Total"
    return text


def split_code_description(value: object) -> tuple[str, str] | None:
    text = clean_code(value)
    if not text:
        return None
    if text == "Total":
        return "Total", "Total"

    match = re.match(
        r"^([A-Z]\d{2}(?:-[A-Z]?\d{2})?(?:,\s*[A-Z]\d{2}(?:-[A-Z]?\d{2})?)*)\s+(.+)$",
        text,
    )
    if match:
        return match.group(1).strip(), clean_text(match.group(2))
    if re.match(r"^[A-Z]\d{2}", text):
        return text, ""
    return None


def normalize_row_code(row: list[object], separate_description: bool) -> tuple[str, str] | None:
    first = clean_text(row[0]) if len(row) > 0 else ""
    second = clean_text(row[1]) if len(row) > 1 else ""

    if separate_description:
        if clean_code(first) == "Total" or clean_code(second) == "Total":
            return "Total", "Total"
        if re.search(r"[A-Z]\d{2}", first):
            return clean_code(first), second
        if re.search(r"[A-Z]\d{2}", second):
            return clean_code(second), ""
        return None

    return split_code_description(first)


def year_from_path(path: Path) -> int:
    if path.parent.name.isdigit():
        return int(path.parent.name)
    match = re.search(r"(19|20)\d{2}", path.name)
    if not match:
        raise ValueError(f"Could not infer year from {path}")
    return int(match.group(0))


def discover_summary_files(data_dir: Path) -> list[Path]:
    files: list[Path] = []

    for year_dir in sorted([p for p in data_dir.iterdir() if p.is_dir() and p.name.isdigit()]):
        xls_files = sorted(year_dir.glob("*.xls"))
        sum_files = [p for p in xls_files if "sum" in p.name.lower()]
        if sum_files:
            files.append(sum_files[0])
            continue
        possible = [
            p
            for p in xls_files
            if "4cha" not in p.name.lower()
            and "4char" not in p.name.lower()
            and "3cha" not in p.name.lower()
            and "3char" not in p.name.lower()
        ]
        if possible:
            files.append(possible[0])

    for xlsx in sorted(data_dir.glob("hosp-epis-stat-admi-diag-20*.xlsx")):
        if not xlsx.name.startswith("~$"):
            files.append(xlsx)

    unique_by_year: dict[int, Path] = {}
    for path in files:
        unique_by_year[year_from_path(path)] = path
    return [unique_by_year[year] for year in sorted(unique_by_year)]


def parse_summary_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".xlsx":
        rows = read_xlsx_rows(
            path,
            lambda name: "primary" in name.lower() and "summary" in name.lower(),
        )
    else:
        workbook = pd.ExcelFile(path)
        sheet_name = next(
            (name for name in workbook.sheet_names if "summary" in name.lower()),
            workbook.sheet_names[0],
        )
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        rows = raw.where(pd.notna(raw), "").values.tolist()

    header_idx = find_header_row(rows)
    headers = [clean_header(value) for value in rows[header_idx]]
    header_map = build_header_map(headers)
    age_map = broad_age_columns(headers)
    fce_idx = header_map.get("fce")
    if fce_idx is None:
        raise ValueError(f"Finished consultant episodes column not found in {path.name}")

    separate_description = fce_idx >= 2
    year_start = year_from_path(path)
    records = []

    for row in rows[header_idx + 1 :]:
        if not any(clean_text(value) for value in row):
            continue

        code_desc = normalize_row_code(row, separate_description)
        if not code_desc:
            continue
        code, description = code_desc
        if code != "Total" and not re.search(r"[A-Z]\d{2}", code):
            continue

        record = {
            "year_start": year_start,
            "fiscal_year": fiscal_label(year_start),
            "source_file": path.name,
            "code": code,
            "description": description or code,
        }
        for field, idx in header_map.items():
            if idx is not None and idx < len(row):
                record[field] = to_number(row[idx])

        for bucket, indices in age_map.items():
            record[bucket] = np.nansum([to_number(row[idx]) for idx in indices if idx < len(row)])

        records.append(record)

    return pd.DataFrame.from_records(records)


def first_icd_code(code: str) -> tuple[str, int] | None:
    match = re.search(r"([A-Z])(\d{2})", code)
    if not match:
        return None
    return match.group(1), int(match.group(2))


def chapter_for_code(code: str) -> str:
    parsed = first_icd_code(code)
    if not parsed:
        return "Other or unmapped"
    letter, number = parsed

    if letter in {"A", "B"}:
        return "Infectious and parasitic"
    if letter == "C" or (letter == "D" and number <= 48):
        return "Neoplasms"
    if letter == "D":
        return "Blood and immune"
    if letter == "E":
        return "Endocrine and metabolic"
    if letter == "F":
        return "Mental health"
    if letter == "G":
        return "Nervous system"
    if letter == "H" and number <= 59:
        return "Eye"
    if letter == "H":
        return "Ear"
    if letter == "I":
        return "Circulatory"
    if letter == "J":
        return "Respiratory"
    if letter == "K":
        return "Digestive"
    if letter == "L":
        return "Skin"
    if letter == "M":
        return "Musculoskeletal"
    if letter == "N":
        return "Genitourinary"
    if letter == "O":
        return "Pregnancy and childbirth"
    if letter == "P":
        return "Perinatal"
    if letter == "Q":
        return "Congenital"
    if letter == "R":
        return "Symptoms and signs"
    if letter in {"S", "T"}:
        return "Injury and poisoning"
    if letter in {"V", "W", "X", "Y"}:
        return "External causes"
    if letter == "Z":
        return "Health services factors"
    if letter == "U":
        return "Special purpose codes"
    return "Other or unmapped"


def load_summary_data() -> pd.DataFrame:
    files = discover_summary_files(DATA_DIR)
    frames = [parse_summary_file(path) for path in files]
    summary = pd.concat(frames, ignore_index=True)

    for field in METRIC_FIELDS + ["mean_wait", "median_wait", "mean_los", "median_los", "mean_age"]:
        if field not in summary:
            summary[field] = np.nan
        summary[field] = pd.to_numeric(summary[field], errors="coerce")

    summary["chapter"] = np.where(
        summary["code"].eq("Total"),
        "Total",
        summary["code"].map(chapter_for_code),
    )
    summary["category"] = summary["code"] + " " + summary["description"].fillna("")
    return summary.sort_values(["year_start", "code"]).reset_index(drop=True)


def weighted_average(values: pd.Series, weights: pd.Series) -> float:
    mask = values.notna() & weights.notna() & (weights > 0)
    if not mask.any():
        return np.nan
    return float(np.average(values[mask], weights=weights[mask]))


def make_chapter_data(summary: pd.DataFrame) -> pd.DataFrame:
    data = summary[summary["code"] != "Total"].copy()
    grouped = []
    for (year_start, fiscal_year, chapter), group in data.groupby(
        ["year_start", "fiscal_year", "chapter"], sort=False
    ):
        record = {
            "year_start": year_start,
            "fiscal_year": fiscal_year,
            "chapter": chapter,
        }
        for field in METRIC_FIELDS:
            record[field] = group[field].sum(min_count=1)
        record["mean_age"] = weighted_average(group["mean_age"], group["fce"])
        record["mean_los"] = weighted_average(group["mean_los"], group["fce"])
        grouped.append(record)
    chapter = pd.DataFrame(grouped)
    chapter["chapter"] = pd.Categorical(chapter["chapter"], CHAPTER_ORDER, ordered=True)
    return chapter.sort_values(["year_start", "chapter"]).reset_index(drop=True)


def save_tables(summary: pd.DataFrame, chapter: pd.DataFrame) -> None:
    summary.to_csv(OUTPUT_DIR / "cleaned_summary_categories.csv", index=False)
    chapter.to_csv(OUTPUT_DIR / "chapter_timeseries.csv", index=False)


def records_for_json(frame: pd.DataFrame, columns: list[str]) -> list[dict[str, object]]:
    available = [column for column in columns if column in frame.columns]
    return json.loads(frame[available].to_json(orient="records"))


def build_dashboard(summary: pd.DataFrame, chapter: pd.DataFrame) -> Path:
    category = summary[summary["code"] != "Total"].copy()
    category["label"] = category["code"] + " " + category["description"].fillna("")

    payload = {
        "years": (
            chapter[["year_start", "fiscal_year"]]
            .drop_duplicates()
            .sort_values("year_start")
            .to_dict(orient="records")
        ),
        "chapter": records_for_json(
            chapter,
            [
                "year_start",
                "fiscal_year",
                "chapter",
                "admissions",
                "emergency",
                "fce",
                "bed_days",
                "mean_age",
                "mean_los",
                "age_0_14",
                "age_15_59",
                "age_60_74",
                "age_75_plus",
            ],
        ),
        "category": records_for_json(
            category,
            [
                "year_start",
                "fiscal_year",
                "code",
                "description",
                "label",
                "chapter",
                "admissions",
                "emergency",
                "fce",
                "bed_days",
                "mean_los",
                "mean_age",
            ],
        ),
    }

    template = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Interactive Hospital Admissions Dashboard</title>
  <style>
    :root {
      --bg: #fbfaf7;
      --panel: #fffdf8;
      --ink: #202226;
      --muted: #5f6368;
      --line: #ded8cf;
      --teal: #2a9d8f;
      --navy: #264653;
      --gold: #e9c46a;
      --orange: #f4a261;
      --red: #e76f51;
      --purple: #6d597a;
      --blue: #457b9d;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Arial, Helvetica, sans-serif;
    }
    main {
      max-width: 1480px;
      margin: 0 auto;
      padding: 24px;
    }
    header {
      display: flex;
      justify-content: space-between;
      gap: 24px;
      align-items: flex-end;
      margin-bottom: 18px;
    }
    h1 {
      margin: 0 0 6px;
      font-size: 30px;
      line-height: 1.12;
    }
    .subtle {
      color: var(--muted);
      font-size: 14px;
      line-height: 1.45;
    }
    .toolbar {
      display: grid;
      grid-template-columns: minmax(260px, 1fr) auto;
      gap: 14px;
      align-items: start;
      padding: 16px;
      background: var(--panel);
      border: 1px solid var(--line);
      margin-bottom: 16px;
    }
    .year-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(92px, 1fr));
      gap: 7px;
    }
    label.year {
      display: flex;
      align-items: center;
      gap: 6px;
      min-height: 30px;
      padding: 5px 8px;
      border: 1px solid var(--line);
      background: #ffffff;
      font-size: 13px;
      user-select: none;
    }
    input[type="checkbox"] {
      width: 14px;
      height: 14px;
      accent-color: var(--teal);
    }
    select, button {
      height: 34px;
      border: 1px solid var(--line);
      background: white;
      color: var(--ink);
      font: inherit;
      padding: 0 10px;
    }
    button {
      cursor: pointer;
      font-weight: 700;
    }
    .controls {
      display: grid;
      grid-template-columns: 1fr;
      gap: 8px;
      min-width: 210px;
    }
    .button-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    .kpi {
      padding: 14px;
      background: var(--panel);
      border: 1px solid var(--line);
    }
    .kpi strong {
      display: block;
      font-size: 24px;
      line-height: 1.1;
      margin-bottom: 4px;
    }
    .kpi span {
      color: var(--muted);
      font-size: 13px;
    }
    .grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      align-items: start;
    }
    .wide {
      grid-column: 1 / -1;
    }
    section {
      background: var(--panel);
      border: 1px solid var(--line);
      padding: 14px;
      min-width: 0;
    }
    h2 {
      margin: 0 0 10px;
      font-size: 18px;
      line-height: 1.25;
    }
    svg {
      display: block;
      width: 100%;
      height: 390px;
      overflow: visible;
    }
    .short svg { height: 330px; }
    .axis, .tick {
      stroke: var(--line);
      stroke-width: 1;
      shape-rendering: crispEdges;
    }
    .label {
      fill: var(--muted);
      font-size: 12px;
    }
    .chart-title {
      fill: var(--ink);
      font-size: 13px;
      font-weight: 700;
    }
    .legend {
      display: flex;
      flex-wrap: wrap;
      gap: 10px 16px;
      margin-top: 10px;
      color: var(--muted);
      font-size: 13px;
    }
    .swatch {
      display: inline-block;
      width: 11px;
      height: 11px;
      margin-right: 6px;
      vertical-align: -1px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th, td {
      border-bottom: 1px solid var(--line);
      padding: 8px 6px;
      text-align: right;
    }
    th:first-child, td:first-child {
      text-align: left;
    }
    th {
      color: var(--muted);
      font-weight: 700;
    }
    @media (max-width: 900px) {
      header, .toolbar, .grid, .kpis {
        grid-template-columns: 1fr;
        display: grid;
      }
      .controls { min-width: 0; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>Interactive Hospital Admissions Dashboard</h1>
        <div class="subtle">Select years to show or hide data. Change the metric to redraw every chart.</div>
      </div>
      <div class="subtle" id="status"></div>
    </header>

    <div class="toolbar">
      <div>
        <div class="subtle" style="margin-bottom:8px;">Years</div>
        <div class="year-grid" id="yearControls"></div>
      </div>
      <div class="controls">
        <select id="metricSelect" aria-label="Metric">
          <option value="admissions">Admissions</option>
          <option value="emergency">Emergency admissions</option>
          <option value="fce">Finished consultant episodes</option>
          <option value="bed_days">Bed days</option>
        </select>
        <div class="button-row">
          <button type="button" data-action="all">All years</button>
          <button type="button" data-action="none">None</button>
          <button type="button" data-action="lockdown">2019-24</button>
          <button type="button" data-action="latest">Latest</button>
        </div>
      </div>
    </div>

    <div class="kpis">
      <div class="kpi"><strong id="kpiTotal">-</strong><span id="kpiTotalLabel">Selected total</span></div>
      <div class="kpi"><strong id="kpiEmergency">-</strong><span>Emergency share</span></div>
      <div class="kpi"><strong id="kpiTopChapter">-</strong><span>Largest chapter</span></div>
      <div class="kpi"><strong id="kpiYears">-</strong><span>Years selected</span></div>
    </div>

    <div class="grid">
      <section>
        <h2>Total trend</h2>
        <svg id="trendChart"></svg>
      </section>
      <section>
        <h2>Chapter burden</h2>
        <svg id="chapterBar"></svg>
      </section>
      <section class="wide">
        <h2>Chapter by year heatmap</h2>
        <svg id="heatmap"></svg>
      </section>
      <section>
        <h2>Age profile</h2>
        <svg id="ageChart"></svg>
        <div class="legend" id="ageLegend"></div>
      </section>
      <section>
        <h2>Top detailed categories</h2>
        <svg id="categoryBar"></svg>
      </section>
      <section class="wide">
        <h2>Top category table</h2>
        <table>
          <thead>
            <tr><th>Category</th><th>Chapter</th><th id="metricHead">Admissions</th><th>Emergency share</th><th>Mean age</th></tr>
          </thead>
          <tbody id="categoryTable"></tbody>
        </table>
      </section>
    </div>
  </main>

  <script>
    const DATA = %%DATA%%;
    const palette = ["#264653", "#2a9d8f", "#e9c46a", "#f4a261", "#e76f51", "#457b9d", "#8d6ab8", "#5c946e", "#bc4749", "#6d597a"];
    const ageColors = {
      age_0_14: "#2a9d8f",
      age_15_59: "#e9c46a",
      age_60_74: "#f4a261",
      age_75_plus: "#e76f51"
    };
    const ageLabels = {
      age_0_14: "0-14",
      age_15_59: "15-59",
      age_60_74: "60-74",
      age_75_plus: "75+"
    };
    const metricLabels = {
      admissions: "Admissions",
      emergency: "Emergency admissions",
      fce: "Finished consultant episodes",
      bed_days: "Bed days"
    };

    const yearControls = document.getElementById("yearControls");
    const metricSelect = document.getElementById("metricSelect");
    const status = document.getElementById("status");

    function fmt(value) {
      if (!Number.isFinite(value)) return "-";
      const abs = Math.abs(value);
      if (abs >= 1e6) return (value / 1e6).toFixed(1) + "m";
      if (abs >= 1e3) return Math.round(value / 1e3) + "k";
      return Math.round(value).toLocaleString();
    }
    function pct(value) {
      return Number.isFinite(value) ? Math.round(value * 100) + "%" : "-";
    }
    function getSelectedYears() {
      return [...document.querySelectorAll(".year-check:checked")].map(input => Number(input.value)).sort((a, b) => a - b);
    }
    function setYears(predicate) {
      document.querySelectorAll(".year-check").forEach(input => {
        input.checked = predicate(Number(input.value));
      });
      update();
    }
    function sumBy(rows, key, metric) {
      const map = new Map();
      for (const row of rows) {
        const name = row[key] || "Unknown";
        const current = map.get(name) || 0;
        map.set(name, current + (Number(row[metric]) || 0));
      }
      return [...map.entries()].map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    }
    function clear(svg) {
      while (svg.firstChild) svg.removeChild(svg.firstChild);
    }
    function node(name, attrs = {}, text = "") {
      const el = document.createElementNS("http://www.w3.org/2000/svg", name);
      for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, value);
      if (text) el.textContent = text;
      return el;
    }
    function size(svg) {
      const box = svg.getBoundingClientRect();
      return { width: Math.max(320, box.width), height: Math.max(260, box.height || 360) };
    }
    function scaleLinear(domainMin, domainMax, rangeMin, rangeMax) {
      const span = domainMax - domainMin || 1;
      return value => rangeMin + ((value - domainMin) / span) * (rangeMax - rangeMin);
    }
    function wrapSvgText(text, maxChars) {
      const words = String(text).split(/\s+/);
      const lines = [];
      let line = "";
      for (const word of words) {
        const next = line ? line + " " + word : word;
        if (next.length > maxChars && line) {
          lines.push(line);
          line = word;
        } else {
          line = next;
        }
      }
      if (line) lines.push(line);
      return lines.slice(0, 2);
    }
    function drawAxes(svg, plot, xTicks, yTicks, xScale, yScale) {
      svg.appendChild(node("line", { class: "axis", x1: plot.left, y1: plot.bottom, x2: plot.right, y2: plot.bottom }));
      svg.appendChild(node("line", { class: "axis", x1: plot.left, y1: plot.top, x2: plot.left, y2: plot.bottom }));
      for (const tick of yTicks) {
        const y = yScale(tick);
        svg.appendChild(node("line", { class: "tick", x1: plot.left, y1: y, x2: plot.right, y2: y }));
        svg.appendChild(node("text", { class: "label", x: plot.left - 8, y: y + 4, "text-anchor": "end" }, fmt(tick)));
      }
      for (const tick of xTicks) {
        const x = xScale(tick.year_start);
        const text = node("text", { class: "label", x, y: plot.bottom + 18, "text-anchor": "middle", transform: `rotate(-35 ${x} ${plot.bottom + 18})` }, tick.fiscal_year);
        svg.appendChild(text);
      }
    }
    function niceTicks(maxValue, count = 5) {
      if (!Number.isFinite(maxValue) || maxValue <= 0) return [0];
      const stepRaw = maxValue / count;
      const pow = Math.pow(10, Math.floor(Math.log10(stepRaw)));
      const step = [1, 2, 5, 10].find(x => x * pow >= stepRaw) * pow;
      const ticks = [];
      for (let value = 0; value <= maxValue + step * 0.5; value += step) ticks.push(value);
      return ticks;
    }

    function drawTrend(years, metric) {
      const svg = document.getElementById("trendChart");
      clear(svg);
      const { width, height } = size(svg);
      const plot = { left: 72, right: width - 22, top: 20, bottom: height - 72 };
      const rows = DATA.years
        .filter(year => years.includes(year.year_start))
        .map(year => {
          const value = DATA.chapter.filter(r => r.year_start === year.year_start).reduce((s, r) => s + (Number(r[metric]) || 0), 0);
          return { ...year, value };
        });
      if (!rows.length) return;
      const maxY = Math.max(...rows.map(r => r.value)) * 1.08;
      const x = scaleLinear(Math.min(...years), Math.max(...years), plot.left, plot.right);
      const y = scaleLinear(0, maxY, plot.bottom, plot.top);
      drawAxes(svg, plot, rows, niceTicks(maxY), x, y);
      const path = rows.map((row, idx) => `${idx ? "L" : "M"} ${x(row.year_start)} ${y(row.value)}`).join(" ");
      svg.appendChild(node("path", { d: path, fill: "none", stroke: "#264653", "stroke-width": 3 }));
      for (const row of rows) {
        svg.appendChild(node("circle", { cx: x(row.year_start), cy: y(row.value), r: 4.5, fill: "#264653" }));
      }
    }

    function drawHorizontalBars(svgId, rows, options) {
      const svg = document.getElementById(svgId);
      clear(svg);
      const { width, height } = size(svg);
      const plot = { left: options.left || 210, right: width - 30, top: 12, bottom: height - 26 };
      const maxValue = Math.max(...rows.map(r => r.value), 1);
      const x = scaleLinear(0, maxValue * 1.08, plot.left, plot.right);
      const barH = Math.max(9, Math.min(22, (plot.bottom - plot.top) / rows.length - 4));
      rows.forEach((row, idx) => {
        const y = plot.top + idx * ((plot.bottom - plot.top) / rows.length);
        const color = options.color ? options.color(row, idx) : palette[idx % palette.length];
        svg.appendChild(node("rect", { x: plot.left, y, width: x(row.value) - plot.left, height: barH, fill: color }));
        const label = node("text", { class: "label", x: plot.left - 8, y: y + barH * 0.72, "text-anchor": "end" });
        const lines = wrapSvgText(row.name, options.labelChars || 28);
        lines.forEach((line, lineIdx) => {
          const tspan = node("tspan", { x: plot.left - 8, dy: lineIdx ? 13 : 0 }, line);
          label.appendChild(tspan);
        });
        svg.appendChild(label);
        svg.appendChild(node("text", { class: "label", x: x(row.value) + 6, y: y + barH * 0.72 }, fmt(row.value)));
      });
      for (const tick of niceTicks(maxValue, 4)) {
        const tx = x(tick);
        svg.appendChild(node("line", { class: "tick", x1: tx, y1: plot.top, x2: tx, y2: plot.bottom }));
        svg.appendChild(node("text", { class: "label", x: tx, y: height - 6, "text-anchor": "middle" }, fmt(tick)));
      }
    }

    function drawHeatmap(years, metric) {
      const svg = document.getElementById("heatmap");
      clear(svg);
      const { width, height } = size(svg);
      const chapters = sumBy(DATA.chapter.filter(r => years.includes(r.year_start)), "chapter", metric).slice(0, 14).map(r => r.name);
      const plot = { left: 190, right: width - 22, top: 10, bottom: height - 54 };
      const cellW = (plot.right - plot.left) / Math.max(1, years.length);
      const cellH = (plot.bottom - plot.top) / Math.max(1, chapters.length);
      const maxValue = Math.max(...DATA.chapter.filter(r => years.includes(r.year_start) && chapters.includes(r.chapter)).map(r => Number(r[metric]) || 0), 1);
      const color = value => {
        const t = Math.sqrt((value || 0) / maxValue);
        const r = Math.round(251 + (38 - 251) * t);
        const g = Math.round(250 + (70 - 250) * t);
        const b = Math.round(247 + (83 - 247) * t);
        return `rgb(${r},${g},${b})`;
      };
      chapters.forEach((chapter, rowIdx) => {
        const y = plot.top + rowIdx * cellH;
        svg.appendChild(node("text", { class: "label", x: plot.left - 8, y: y + cellH * 0.62, "text-anchor": "end" }, chapter));
        years.forEach((year, colIdx) => {
          const value = DATA.chapter.find(r => r.year_start === year && r.chapter === chapter)?.[metric] || 0;
          const x = plot.left + colIdx * cellW;
          svg.appendChild(node("rect", { x, y, width: Math.max(1, cellW - 1), height: Math.max(1, cellH - 1), fill: color(value) }));
        });
      });
      years.forEach((year, idx) => {
        const yearInfo = DATA.years.find(y => y.year_start === year);
        const x = plot.left + idx * cellW + cellW / 2;
        svg.appendChild(node("text", { class: "label", x, y: plot.bottom + 18, "text-anchor": "middle", transform: `rotate(-35 ${x} ${plot.bottom + 18})` }, yearInfo?.fiscal_year || year));
      });
    }

    function drawAge(years) {
      const svg = document.getElementById("ageChart");
      clear(svg);
      const ageKeys = Object.keys(ageLabels);
      const rows = ageKeys.map(key => ({
        key,
        name: ageLabels[key],
        value: DATA.chapter.filter(r => years.includes(r.year_start)).reduce((sum, row) => sum + (Number(row[key]) || 0), 0)
      }));
      const total = rows.reduce((sum, row) => sum + row.value, 0) || 1;
      const { width, height } = size(svg);
      const plot = { left: 72, right: width - 30, top: 34, bottom: height - 44 };
      let x0 = plot.left;
      rows.forEach(row => {
        const w = (row.value / total) * (plot.right - plot.left);
        svg.appendChild(node("rect", { x: x0, y: plot.top, width: w, height: plot.bottom - plot.top, fill: ageColors[row.key] }));
        if (w > 56) {
          svg.appendChild(node("text", { x: x0 + w / 2, y: (plot.top + plot.bottom) / 2, "text-anchor": "middle", fill: "#202226", "font-size": 13, "font-weight": 700 }, `${row.name} ${pct(row.value / total)}`));
        }
        x0 += w;
      });
      document.getElementById("ageLegend").innerHTML = rows.map(row => `<span><i class="swatch" style="background:${ageColors[row.key]}"></i>${row.name}: ${fmt(row.value)}</span>`).join("");
    }

    function updateTable(rows, metric) {
      document.getElementById("metricHead").textContent = metricLabels[metric];
      const tbody = document.getElementById("categoryTable");
      tbody.innerHTML = rows.slice(0, 12).map(row => {
        const emergencyShare = row.admissions ? row.emergency / row.admissions : NaN;
        return `<tr><td>${row.name}</td><td>${row.chapter}</td><td>${fmt(row.value)}</td><td>${pct(emergencyShare)}</td><td>${Number.isFinite(row.meanAge) ? row.meanAge.toFixed(1) : "-"}</td></tr>`;
      }).join("");
    }

    function update() {
      const years = getSelectedYears();
      const metric = metricSelect.value;
      const chapterRows = DATA.chapter.filter(row => years.includes(row.year_start));
      const categoryRows = DATA.category.filter(row => years.includes(row.year_start));
      const total = chapterRows.reduce((sum, row) => sum + (Number(row[metric]) || 0), 0);
      const admissions = chapterRows.reduce((sum, row) => sum + (Number(row.admissions) || 0), 0);
      const emergency = chapterRows.reduce((sum, row) => sum + (Number(row.emergency) || 0), 0);
      const chapters = sumBy(chapterRows, "chapter", metric);
      const catMap = new Map();
      for (const row of categoryRows) {
        const key = row.code + " " + row.description;
        const current = catMap.get(key) || { name: key, chapter: row.chapter, value: 0, admissions: 0, emergency: 0, ageWeight: 0, meanAgeSum: 0 };
        const value = Number(row[metric]) || 0;
        const adm = Number(row.admissions) || 0;
        const emg = Number(row.emergency) || 0;
        const meanAge = Number(row.mean_age);
        current.value += value;
        current.admissions += adm;
        current.emergency += emg;
        if (Number.isFinite(meanAge) && adm > 0) {
          current.meanAgeSum += meanAge * adm;
          current.ageWeight += adm;
        }
        catMap.set(key, current);
      }
      const categories = [...catMap.values()]
        .map(row => ({ ...row, meanAge: row.ageWeight ? row.meanAgeSum / row.ageWeight : NaN }))
        .sort((a, b) => b.value - a.value);

      document.getElementById("kpiTotal").textContent = fmt(total);
      document.getElementById("kpiTotalLabel").textContent = `Selected ${metricLabels[metric].toLowerCase()}`;
      document.getElementById("kpiEmergency").textContent = pct(emergency / admissions);
      document.getElementById("kpiTopChapter").textContent = chapters[0]?.name || "-";
      document.getElementById("kpiYears").textContent = years.length;
      status.textContent = years.length ? `${DATA.years.find(y => y.year_start === years[0])?.fiscal_year} to ${DATA.years.find(y => y.year_start === years[years.length - 1])?.fiscal_year}` : "No years selected";

      drawTrend(years, metric);
      drawHorizontalBars("chapterBar", chapters.slice(0, 12), { left: 190, labelChars: 24, color: (_row, idx) => palette[idx % palette.length] });
      drawHeatmap(years, metric);
      drawAge(years);
      drawHorizontalBars("categoryBar", categories.slice(0, 12), { left: 250, labelChars: 32, color: (_row, idx) => idx < 3 ? "#e76f51" : "#457b9d" });
      updateTable(categories, metric);
    }

    function init() {
      yearControls.innerHTML = DATA.years.map(year => `
        <label class="year">
          <input class="year-check" type="checkbox" value="${year.year_start}" checked>
          <span>${year.fiscal_year}</span>
        </label>
      `).join("");
      document.querySelectorAll(".year-check").forEach(input => input.addEventListener("change", update));
      metricSelect.addEventListener("change", update);
      document.querySelectorAll("button[data-action]").forEach(button => {
        button.addEventListener("click", () => {
          const action = button.dataset.action;
          const maxYear = Math.max(...DATA.years.map(y => y.year_start));
          if (action === "all") setYears(() => true);
          if (action === "none") setYears(() => false);
          if (action === "lockdown") setYears(year => year >= 2019);
          if (action === "latest") setYears(year => year === maxYear);
        });
      });
      window.addEventListener("resize", update);
      update();
    }
    init();
  </script>
</body>
</html>
"""
    html = template.replace("%%DATA%%", json.dumps(payload, ensure_ascii=False))
    path = OUTPUT_DIR / "dashboard.html"
    path.write_text(html, encoding="utf-8")
    return path


def build_advanced_dashboard(summary: pd.DataFrame, chapter: pd.DataFrame) -> Path:
    category = summary[summary["code"] != "Total"].copy()
    category["label"] = category["code"] + " " + category["description"].fillna("")

    payload = {
        "years": (
            chapter[["year_start", "fiscal_year"]]
            .drop_duplicates()
            .sort_values("year_start")
            .to_dict(orient="records")
        ),
        "chapter": records_for_json(
            chapter,
            [
                "year_start",
                "fiscal_year",
                "chapter",
                "admissions",
                "emergency",
                "fce",
                "bed_days",
                "mean_age",
                "mean_los",
                "age_0_14",
                "age_15_59",
                "age_60_74",
                "age_75_plus",
            ],
        ),
        "category": records_for_json(
            category,
            [
                "year_start",
                "fiscal_year",
                "code",
                "description",
                "label",
                "chapter",
                "admissions",
                "emergency",
                "fce",
                "bed_days",
                "mean_los",
                "mean_age",
                "age_0_14",
                "age_15_59",
                "age_60_74",
                "age_75_plus",
            ],
        ),
    }

    template = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Advanced Hospital Admissions Visual Analysis</title>
  <style>
    :root {
      --bg: #fbfaf7;
      --panel: #fffdf8;
      --ink: #202226;
      --muted: #5f6368;
      --line: #ded8cf;
      --teal: #2a9d8f;
      --navy: #264653;
      --gold: #e9c46a;
      --orange: #f4a261;
      --red: #e76f51;
      --purple: #6d597a;
      --blue: #457b9d;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Arial, Helvetica, sans-serif;
    }
    main {
      max-width: 1500px;
      margin: 0 auto;
      padding: 24px;
    }
    header {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 18px;
      align-items: end;
      margin-bottom: 16px;
    }
    h1 { margin: 0 0 6px; font-size: 31px; line-height: 1.12; }
    h2 { margin: 0 0 10px; font-size: 18px; line-height: 1.2; }
    p { margin: 0; }
    .subtle { color: var(--muted); font-size: 14px; line-height: 1.45; }
    .toolbar {
      display: grid;
      grid-template-columns: minmax(320px, 1fr) 230px;
      gap: 14px;
      padding: 15px;
      background: var(--panel);
      border: 1px solid var(--line);
      margin-bottom: 16px;
    }
    .year-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(92px, 1fr));
      gap: 7px;
    }
    label.year {
      display: flex;
      align-items: center;
      gap: 6px;
      min-height: 30px;
      padding: 5px 8px;
      border: 1px solid var(--line);
      background: white;
      font-size: 13px;
      user-select: none;
    }
    input[type="checkbox"] { width: 14px; height: 14px; accent-color: var(--teal); }
    select, button {
      height: 34px;
      border: 1px solid var(--line);
      background: white;
      color: var(--ink);
      font: inherit;
      padding: 0 10px;
    }
    button { cursor: pointer; font-weight: 700; }
    .controls { display: grid; grid-template-columns: 1fr; gap: 8px; }
    .button-row { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
    .insights {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    .insight {
      padding: 14px;
      background: var(--panel);
      border: 1px solid var(--line);
      min-height: 112px;
    }
    .insight strong {
      display: block;
      font-size: 16px;
      margin-bottom: 7px;
    }
    .scope {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    .scope-box {
      padding: 12px 14px;
      background: var(--panel);
      border: 1px solid var(--line);
    }
    .scope-box strong {
      display: block;
      margin-bottom: 6px;
      font-size: 14px;
    }
    .level-tag {
      display: inline-block;
      margin-bottom: 8px;
      padding: 4px 8px;
      border: 1px solid var(--line);
      color: var(--muted);
      background: #fff;
      font-size: 12px;
      font-weight: 700;
    }
    .grid {
      display: grid;
      grid-template-columns: 1.05fr 1fr;
      gap: 16px;
      align-items: start;
    }
    section {
      background: var(--panel);
      border: 1px solid var(--line);
      padding: 14px;
      min-width: 0;
    }
    .wide { grid-column: 1 / -1; }
    svg {
      display: block;
      width: 100%;
      height: 500px;
      overflow: visible;
    }
    .medium svg { height: 430px; }
    .tall svg { height: 620px; }
    .axis, .gridline {
      stroke: var(--line);
      stroke-width: 1;
      shape-rendering: crispEdges;
    }
    .label { fill: var(--muted); font-size: 12px; }
    .small-label { fill: var(--muted); font-size: 10.5px; }
    .title-label { fill: var(--ink); font-size: 13px; font-weight: 700; }
    .legend {
      display: flex;
      flex-wrap: wrap;
      gap: 9px 16px;
      margin-top: 10px;
      color: var(--muted);
      font-size: 13px;
    }
    .swatch {
      display: inline-block;
      width: 11px;
      height: 11px;
      margin-right: 6px;
      vertical-align: -1px;
    }
    .note {
      margin-top: 8px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
    }
    .panel-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
      gap: 12px;
    }
    .panel-card {
      border: 1px solid var(--line);
      background: #fff;
      padding: 10px;
      min-width: 0;
    }
    .panel-head {
      display: grid;
      gap: 4px;
      margin-bottom: 8px;
    }
    .panel-title {
      font-size: 13px;
      font-weight: 700;
      line-height: 1.25;
    }
    .panel-meta {
      color: var(--muted);
      font-size: 11.5px;
      line-height: 1.35;
    }
    .panel-svg {
      display: block;
      width: 100%;
      height: 196px;
      overflow: visible;
    }
    .panel-legend {
      margin-top: 10px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.4;
    }
    .parallel-layout {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 190px;
      gap: 12px;
      align-items: start;
    }
    .parallel-side {
      border: 1px solid var(--line);
      background: #fff;
      padding: 10px;
      min-height: 430px;
    }
    .parallel-side-title {
      font-size: 12px;
      font-weight: 700;
      margin-bottom: 8px;
    }
    .parallel-actions {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 6px;
      margin-bottom: 10px;
    }
    .parallel-actions button {
      height: 30px;
      padding: 0 8px;
      font-size: 12px;
    }
    .parallel-filters {
      display: grid;
      gap: 6px;
      max-height: 370px;
      overflow: auto;
      padding-right: 2px;
    }
    label.parallel-option {
      display: flex;
      align-items: flex-start;
      gap: 7px;
      padding: 6px 4px;
      border: 1px solid var(--line);
      background: #fffdf8;
      font-size: 12px;
      line-height: 1.3;
    }
    label.parallel-option input {
      margin-top: 2px;
    }
    .parallel-empty {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.4;
      padding: 12px 4px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th, td {
      border-bottom: 1px solid var(--line);
      padding: 8px 6px;
      text-align: right;
      vertical-align: top;
    }
    th:first-child, td:first-child { text-align: left; }
    th { color: var(--muted); font-weight: 700; }
    @media (max-width: 980px) {
      header, .toolbar, .grid, .insights {
        grid-template-columns: 1fr;
      }
      .scope { grid-template-columns: 1fr; }
      .parallel-layout { grid-template-columns: 1fr; }
      .parallel-side { min-height: 0; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>Advanced Hospital Admissions Visual Analysis</h1>
        <p class="subtle">The main view is a small-multiple age-year heatmap for diagnosis categories, supported by matrix, radial, parallel, and age-composition views. Year checkboxes dynamically change every view.</p>
      </div>
      <div class="subtle" id="status"></div>
    </header>

    <div class="toolbar">
      <div>
        <div class="subtle" style="margin-bottom:8px;">Years to include</div>
        <div class="year-grid" id="yearControls"></div>
      </div>
      <div class="controls">
        <select id="metricSelect" aria-label="Metric">
          <option value="admissions">Admissions</option>
          <option value="emergency">Emergency admissions</option>
          <option value="fce">Finished consultant episodes</option>
          <option value="bed_days">Bed days</option>
        </select>
        <div class="button-row">
          <button type="button" data-action="all">All years</button>
          <button type="button" data-action="none">None</button>
          <button type="button" data-action="lockdown">2019-24</button>
          <button type="button" data-action="latest">Latest</button>
        </div>
      </div>
    </div>

    <div class="grid">
      <section class="wide">
        <h2>Small multiples: age-group trends by fiscal year</h2>
        <div class="level-tag">Level: primary diagnosis summary · Columns: fiscal years · Rows: 0-14, 15-59, 60-74, 75+</div>
        <div class="note" id="ageYearSummary">Primary diagnosis summary categories are shown with code + description. Fiscal years run across columns, age groups run down rows, and 3-character / 4-character diagnosis levels are intentionally excluded from this view.</div>
        <div id="ageYearPanels" class="panel-grid"></div>
        <div class="panel-legend">This is the main view for research question C. Each panel is one primary diagnosis summary category labeled with both code and description. Cell color shows age-specific admissions counts within that category using a panel-specific scale, so age-pattern changes remain visible even when total burden differs sharply across categories. The source tables provide age splits for admissions, not age-specific FCE or bed days, so this view remains fixed to admissions while the supporting views switch metric.</div>
      </section>

      <section class="medium">
        <h2>Parallel coordinates: multi-attribute profile</h2>
        <div class="level-tag">Level: ICD-10 chapter</div>
        <div class="parallel-layout">
          <svg id="parallel"></svg>
          <aside class="parallel-side">
            <div class="parallel-side-title">Show or hide chapters</div>
            <div class="parallel-actions">
              <button type="button" id="parallelAll">All</button>
              <button type="button" id="parallelNone">None</button>
            </div>
            <div class="parallel-filters" id="parallelFilters"></div>
          </aside>
        </div>
        <div class="note">Each chapter becomes a path across burden, emergency share, mean age, length of stay, and 75+ share. Crossings reveal relationships and exceptions.</div>
      </section>

      <section class="wide">
        <h2>Matrix heatmap: chapter x fiscal year</h2>
        <div class="level-tag">Level: ICD-10 chapter · Columns: fiscal years</div>
        <svg id="matrix"></svg>
        <div class="note">Rows are ICD-10 chapters, columns are selected years, color is intensity. This is better than many lines because it shows all categories at once and highlights the lockdown break.</div>
      </section>

      <section class="medium">
        <h2>Radial year-ring matrix</h2>
        <div class="level-tag">Level: ICD-10 chapter · Rings: fiscal years</div>
        <svg id="radial"></svg>
        <div class="note">Rings are years, sectors are chapters. The 2020-21 ring makes system-wide disruption visible as a circular anomaly.</div>
      </section>

      <section class="medium">
        <h2>Age-composition matrix</h2>
        <div class="level-tag">Level: primary diagnosis summary · Columns: 0-14, 15-59, 60-74, 75+</div>
        <svg id="ageMatrix"></svg>
        <div class="note">Rows are high-burden detailed categories, columns are age groups. Color shows share within category, making age disparity visible without separate charts.</div>
      </section>

      <section class="wide">
        <h2>Discovery table: largest detailed categories for the selected years</h2>
        <div class="level-tag">Level: primary diagnosis summary category</div>
        <table>
          <thead>
            <tr><th>Category</th><th>Chapter</th><th id="metricHead">Admissions</th><th>Emergency share</th><th>Mean age</th><th>Mean stay</th></tr>
          </thead>
          <tbody id="categoryTable"></tbody>
        </table>
      </section>
    </div>
  </main>

  <script>
    const DATA = %%DATA%%;
    const palette = ["#264653", "#2a9d8f", "#e9c46a", "#f4a261", "#e76f51", "#457b9d", "#8d6ab8", "#5c946e", "#bc4749", "#6d597a", "#118ab2", "#ef476f"];
    const metricLabels = {
      admissions: "Admissions",
      emergency: "Emergency admissions",
      fce: "Finished consultant episodes",
      bed_days: "Bed days"
    };
    const ageKeys = ["age_0_14", "age_15_59", "age_60_74", "age_75_plus"];
    const ageLabels = ["0-14", "15-59", "60-74", "75+"];
    const ageColors = ["#d8f3dc", "#95d5b2", "#52b788", "#1b4332"];
    const parallelSelection = new Map();

    const yearControls = document.getElementById("yearControls");
    const metricSelect = document.getElementById("metricSelect");

    function fmt(value) {
      if (!Number.isFinite(value)) return "-";
      const abs = Math.abs(value);
      if (abs >= 1e6) return (value / 1e6).toFixed(1) + "m";
      if (abs >= 1e3) return Math.round(value / 1e3) + "k";
      return Math.round(value).toLocaleString();
    }
    function pct(value) {
      return Number.isFinite(value) ? Math.round(value * 100) + "%" : "-";
    }
    function cleanNumber(value) {
      const number = Number(value);
      return Number.isFinite(number) ? number : 0;
    }
    function selectedYears() {
      return [...document.querySelectorAll(".year-check:checked")].map(input => Number(input.value)).sort((a, b) => a - b);
    }
    function setYears(predicate) {
      document.querySelectorAll(".year-check").forEach(input => input.checked = predicate(Number(input.value)));
      update();
    }
    function node(name, attrs = {}, text = "") {
      const el = document.createElementNS("http://www.w3.org/2000/svg", name);
      for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, value);
      if (text) el.textContent = text;
      return el;
    }
    function clear(svg) {
      while (svg.firstChild) svg.removeChild(svg.firstChild);
    }
    function size(svg) {
      const box = svg.getBoundingClientRect();
      return { width: Math.max(360, box.width), height: Math.max(300, box.height || 420) };
    }
    function lerp(a, b, t) { return a + (b - a) * t; }
    function colorScale(value, maxValue, low = [251, 250, 247], high = [38, 70, 83]) {
      const t = Math.max(0, Math.min(1, Math.sqrt((value || 0) / (maxValue || 1))));
      return `rgb(${Math.round(lerp(low[0], high[0], t))},${Math.round(lerp(low[1], high[1], t))},${Math.round(lerp(low[2], high[2], t))})`;
    }
    function scale(min, max, a, b) {
      const span = max - min || 1;
      return value => a + ((value - min) / span) * (b - a);
    }
    function wrap(text, maxChars) {
      const words = String(text).split(/\s+/);
      const lines = [];
      let line = "";
      for (const word of words) {
        const next = line ? `${line} ${word}` : word;
        if (next.length > maxChars && line) {
          lines.push(line);
          line = word;
        } else {
          line = next;
        }
      }
      if (line) lines.push(line);
      return lines.slice(0, 2);
    }
    function fiscal(year) {
      return DATA.years.find(item => item.year_start === year)?.fiscal_year || String(year);
    }
    function aggregateChapters(years, metric) {
      const map = new Map();
      for (const row of DATA.chapter.filter(item => years.includes(item.year_start))) {
        const current = map.get(row.chapter) || {
          name: row.chapter,
          value: 0,
          admissions: 0,
          emergency: 0,
          fce: 0,
          bed_days: 0,
          age_0_14: 0,
          age_15_59: 0,
          age_60_74: 0,
          age_75_plus: 0,
          meanAgeSum: 0,
          meanLosSum: 0,
          weight: 0
        };
        const admissions = cleanNumber(row.admissions);
        current.value += cleanNumber(row[metric]);
        current.admissions += admissions;
        current.emergency += cleanNumber(row.emergency);
        current.fce += cleanNumber(row.fce);
        current.bed_days += cleanNumber(row.bed_days);
        for (const key of ageKeys) current[key] += cleanNumber(row[key]);
        if (admissions > 0) {
          current.meanAgeSum += cleanNumber(row.mean_age) * admissions;
          current.meanLosSum += cleanNumber(row.mean_los) * admissions;
          current.weight += admissions;
        }
        map.set(row.chapter, current);
      }
      return [...map.values()].map(row => ({
        ...row,
        emergencyShare: row.admissions ? row.emergency / row.admissions : 0,
        meanAge: row.weight ? row.meanAgeSum / row.weight : 0,
        meanLos: row.weight ? row.meanLosSum / row.weight : 0,
        olderShare: (row.age_0_14 + row.age_15_59 + row.age_60_74 + row.age_75_plus) ? row.age_75_plus / (row.age_0_14 + row.age_15_59 + row.age_60_74 + row.age_75_plus) : 0
      })).sort((a, b) => b.value - a.value);
    }
    function aggregateCategories(years, metric) {
      const map = new Map();
      for (const row of DATA.category.filter(item => years.includes(item.year_start))) {
        const key = row.label || `${row.code} ${row.description}`;
        const current = map.get(key) || {
          name: key,
          chapter: row.chapter,
          value: 0,
          admissions: 0,
          emergency: 0,
          meanAgeSum: 0,
          meanLosSum: 0,
          weight: 0,
          age_0_14: 0,
          age_15_59: 0,
          age_60_74: 0,
          age_75_plus: 0
        };
        const admissions = cleanNumber(row.admissions);
        current.value += cleanNumber(row[metric]);
        current.admissions += admissions;
        current.emergency += cleanNumber(row.emergency);
        current.meanAgeSum += cleanNumber(row.mean_age) * admissions;
        current.meanLosSum += cleanNumber(row.mean_los) * admissions;
        current.weight += admissions;
        for (const age of ageKeys) current[age] += cleanNumber(row[age]);
        map.set(key, current);
      }
      return [...map.values()].map(row => ({
        ...row,
        emergencyShare: row.admissions ? row.emergency / row.admissions : 0,
        meanAge: row.weight ? row.meanAgeSum / row.weight : 0,
        meanLos: row.weight ? row.meanLosSum / row.weight : 0
      })).sort((a, b) => b.value - a.value);
    }
    function drawMatrix(years, metric) {
      const svg = document.getElementById("matrix");
      clear(svg);
      const { width, height } = size(svg);
      const chapters = aggregateChapters(years, metric).slice(0, 16).map(row => row.name);
      const plot = { left: 190, top: 12, right: width - 26, bottom: height - 68 };
      const cellW = (plot.right - plot.left) / Math.max(1, years.length);
      const cellH = (plot.bottom - plot.top) / Math.max(1, chapters.length);
      const allValues = DATA.chapter.filter(row => years.includes(row.year_start) && chapters.includes(row.chapter)).map(row => cleanNumber(row[metric]));
      const maxValue = Math.max(...allValues, 1);
      chapters.forEach((chapter, rowIdx) => {
        const y = plot.top + rowIdx * cellH;
        svg.appendChild(node("text", { class: "label", x: plot.left - 8, y: y + cellH * 0.62, "text-anchor": "end" }, chapter));
        years.forEach((year, colIdx) => {
          const x = plot.left + colIdx * cellW;
          const value = cleanNumber(DATA.chapter.find(row => row.year_start === year && row.chapter === chapter)?.[metric]);
          svg.appendChild(node("rect", { x, y, width: Math.max(1, cellW - 1), height: Math.max(1, cellH - 1), fill: colorScale(value, maxValue) }));
          if (cellW > 54 && cellH > 21) svg.appendChild(node("text", { x: x + cellW / 2, y: y + cellH * 0.62, "text-anchor": "middle", fill: value > maxValue * 0.45 ? "white" : "#202226", "font-size": 10 }, fmt(value)));
        });
      });
      years.forEach((year, idx) => {
        const x = plot.left + idx * cellW + cellW / 2;
        svg.appendChild(node("text", { class: "label", x, y: plot.bottom + 18, "text-anchor": "middle", transform: `rotate(-35 ${x} ${plot.bottom + 18})` }, fiscal(year)));
      });
    }

    function arcPath(cx, cy, r0, r1, a0, a1) {
      const p = (r, a) => [cx + r * Math.cos(a), cy + r * Math.sin(a)];
      const [x00, y00] = p(r0, a0), [x01, y01] = p(r0, a1), [x10, y10] = p(r1, a0), [x11, y11] = p(r1, a1);
      const large = a1 - a0 > Math.PI ? 1 : 0;
      return `M ${x10} ${y10} A ${r1} ${r1} 0 ${large} 1 ${x11} ${y11} L ${x01} ${y01} A ${r0} ${r0} 0 ${large} 0 ${x00} ${y00} Z`;
    }
    function drawRadial(years, metric) {
      const svg = document.getElementById("radial");
      clear(svg);
      const { width, height } = size(svg);
      const cx = width / 2, cy = height / 2 + 6;
      const chapters = aggregateChapters(years, metric).slice(0, 12).map(row => row.name);
      const maxRadius = Math.min(width, height) * 0.43;
      const inner = Math.min(width, height) * 0.12;
      const ring = (maxRadius - inner) / Math.max(1, years.length);
      const maxValue = Math.max(...DATA.chapter.filter(row => years.includes(row.year_start) && chapters.includes(row.chapter)).map(row => cleanNumber(row[metric])), 1);
      years.forEach((year, ringIdx) => {
        const r0 = inner + ringIdx * ring;
        const r1 = r0 + ring - 1;
        chapters.forEach((chapter, sectorIdx) => {
          const gap = 0.012;
          const a0 = -Math.PI / 2 + sectorIdx * 2 * Math.PI / chapters.length + gap;
          const a1 = -Math.PI / 2 + (sectorIdx + 1) * 2 * Math.PI / chapters.length - gap;
          const value = cleanNumber(DATA.chapter.find(row => row.year_start === year && row.chapter === chapter)?.[metric]);
          svg.appendChild(node("path", { d: arcPath(cx, cy, r0, r1, a0, a1), fill: colorScale(value, maxValue, [255, 253, 248], [231, 111, 81]), stroke: "#fbfaf7", "stroke-width": 0.8 }));
        });
      });
      chapters.forEach((chapter, idx) => {
        const a = -Math.PI / 2 + (idx + 0.5) * 2 * Math.PI / chapters.length;
        const x = cx + (maxRadius + 20) * Math.cos(a);
        const y = cy + (maxRadius + 20) * Math.sin(a);
        svg.appendChild(node("text", { class: "small-label", x, y, "text-anchor": x < cx ? "end" : "start" }, chapter.split(" ")[0]));
      });
      svg.appendChild(node("text", { class: "title-label", x: cx, y: cy - 4, "text-anchor": "middle" }, "Years"));
      svg.appendChild(node("text", { class: "small-label", x: cx, y: cy + 12, "text-anchor": "middle" }, `${fiscal(years[0])}`));
      if (years.length > 1) svg.appendChild(node("text", { class: "small-label", x: cx, y: cy + 27, "text-anchor": "middle" }, `to ${fiscal(years[years.length - 1])}`));
    }

    function drawParallel(chapters, metric) {
      const svg = document.getElementById("parallel");
      clear(svg);
      const { width, height } = size(svg);
      const rows = chapters.slice(0, 14);
      const filterHost = document.getElementById("parallelFilters");
      rows.forEach(row => {
        if (!parallelSelection.has(row.name)) parallelSelection.set(row.name, true);
      });
      for (const key of [...parallelSelection.keys()]) {
        if (!rows.some(row => row.name === key)) parallelSelection.delete(key);
      }
      filterHost.innerHTML = rows.map((row, idx) => `
        <label class="parallel-option">
          <input type="checkbox" data-parallel-name="${row.name}" ${parallelSelection.get(row.name) !== false ? "checked" : ""}>
          <span><span class="swatch" style="background:${palette[idx % palette.length]}"></span>${row.name}</span>
        </label>
      `).join("");
      const visibleRows = rows.filter(row => parallelSelection.get(row.name) !== false);
      if (!visibleRows.length) {
        svg.appendChild(node("text", { class: "label", x: width / 2, y: height / 2, "text-anchor": "middle" }, "Select at least one chapter on the right"));
        filterHost.querySelectorAll("input").forEach(input => {
          input.addEventListener("change", () => {
            parallelSelection.set(input.dataset.parallelName, input.checked);
            update();
          });
        });
        return;
      }
      const dims = [
        { key: "value", label: metricLabels[metric] },
        { key: "emergencyShare", label: "Emergency share", format: pct },
        { key: "meanAge", label: "Mean age", format: v => v.toFixed(0) },
        { key: "meanLos", label: "Mean stay", format: v => v.toFixed(1) },
        { key: "olderShare", label: "75+ share", format: pct }
      ];
      const plot = { left: 70, right: width - 22, top: 44, bottom: height - 42 };
      const x = idx => plot.left + idx * (plot.right - plot.left) / (dims.length - 1);
      const scales = {};
      dims.forEach(dim => {
        const values = visibleRows.map(row => cleanNumber(row[dim.key]));
        scales[dim.key] = scale(Math.min(...values), Math.max(...values), plot.bottom, plot.top);
      });
      dims.forEach((dim, idx) => {
        const xx = x(idx);
        svg.appendChild(node("line", { class: "axis", x1: xx, y1: plot.top, x2: xx, y2: plot.bottom }));
        svg.appendChild(node("text", { class: "label", x: xx, y: 18, "text-anchor": "middle" }, dim.label));
        const values = visibleRows.map(row => cleanNumber(row[dim.key]));
        const min = Math.min(...values), max = Math.max(...values);
        svg.appendChild(node("text", { class: "small-label", x: xx + 4, y: plot.top + 4 }, dim.format ? dim.format(max) : fmt(max)));
        svg.appendChild(node("text", { class: "small-label", x: xx + 4, y: plot.bottom }, dim.format ? dim.format(min) : fmt(min)));
      });
      visibleRows.forEach((row, idx) => {
        const colorIdx = rows.findIndex(item => item.name === row.name);
        const d = dims.map((dim, dimIdx) => `${dimIdx ? "L" : "M"} ${x(dimIdx)} ${scales[dim.key](cleanNumber(row[dim.key]))}`).join(" ");
        const color = palette[colorIdx % palette.length];
        svg.appendChild(node("path", { d, fill: "none", stroke: color, "stroke-width": 2.2, opacity: 0.78 }));
      });
      filterHost.querySelectorAll("input").forEach(input => {
        input.addEventListener("change", () => {
          parallelSelection.set(input.dataset.parallelName, input.checked);
          update();
        });
      });
    }

    function drawAgeMatrix(categories) {
      const svg = document.getElementById("ageMatrix");
      clear(svg);
      const { width, height } = size(svg);
      const rows = categories.filter(row => ageKeys.reduce((sum, key) => sum + cleanNumber(row[key]), 0) > 0).slice(0, 14);
      const plot = { left: 260, top: 30, right: width - 28, bottom: height - 26 };
      const cellW = (plot.right - plot.left) / ageKeys.length;
      const cellH = (plot.bottom - plot.top) / Math.max(1, rows.length);
      ageLabels.forEach((label, idx) => {
        svg.appendChild(node("text", { class: "label", x: plot.left + idx * cellW + cellW / 2, y: 16, "text-anchor": "middle" }, label));
      });
      rows.forEach((row, rowIdx) => {
        const y = plot.top + rowIdx * cellH;
        const total = ageKeys.reduce((sum, key) => sum + cleanNumber(row[key]), 0) || 1;
        const label = node("text", { class: "label", x: plot.left - 8, y: y + cellH * 0.62, "text-anchor": "end" });
        wrap(row.name, 34).forEach((line, lineIdx) => label.appendChild(node("tspan", { x: plot.left - 8, dy: lineIdx ? 12 : 0 }, line)));
        svg.appendChild(label);
        ageKeys.forEach((key, idx) => {
          const share = cleanNumber(row[key]) / total;
          const x = plot.left + idx * cellW;
          svg.appendChild(node("rect", { x, y, width: Math.max(1, cellW - 1), height: Math.max(1, cellH - 1), fill: colorScale(share, 1, [251, 250, 247], [27, 67, 50]) }));
          if (cellH > 20) svg.appendChild(node("text", { x: x + cellW / 2, y: y + cellH * 0.62, "text-anchor": "middle", fill: share > 0.45 ? "white" : "#202226", "font-size": 10 }, pct(share)));
        });
      });
    }

    function updateTable(categories, metric) {
      document.getElementById("metricHead").textContent = metricLabels[metric];
      document.getElementById("categoryTable").innerHTML = categories.slice(0, 12).map(row => `
        <tr>
          <td>${row.name}</td>
          <td>${row.chapter}</td>
          <td>${fmt(row.value)}</td>
          <td>${pct(row.emergencyShare)}</td>
          <td>${row.meanAge ? row.meanAge.toFixed(1) : "-"}</td>
          <td>${row.meanLos ? row.meanLos.toFixed(1) : "-"}</td>
        </tr>
      `).join("");
    }

    function signedFmt(value) {
      return `${value > 0 ? "+" : value < 0 ? "-" : ""}${fmt(Math.abs(value))}`;
    }

    function buildAgeYearPanels(years) {
      const shortlist = aggregateCategories(years, "admissions")
        .filter(row => ageKeys.reduce((sum, key) => sum + cleanNumber(row[key]), 0) > 0)
        .slice(0, 24);
      const panels = new Map(shortlist.map(row => [row.name, {
        name: row.name,
        chapter: row.chapter,
        total: row.admissions,
        yearly: {}
      }]));

      for (const row of DATA.category.filter(item => years.includes(item.year_start))) {
        const name = row.label || `${row.code} ${row.description}`;
        if (!panels.has(name)) continue;
        panels.get(name).yearly[row.year_start] = ageKeys.map(key => cleanNumber(row[key]));
      }

      const firstYear = years[0];
      const lastYear = years[years.length - 1];
      return [...panels.values()].map(panel => {
        let maxCell = 0;
        let dominantAge = { label: "-", value: 0 };
        let largestShift = { label: "-", delta: 0, start: 0, end: 0, absDelta: 0 };

        ageKeys.forEach((key, idx) => {
          const label = ageLabels[idx];
          const start = cleanNumber(panel.yearly[firstYear]?.[idx]);
          const end = cleanNumber(panel.yearly[lastYear]?.[idx]);
          const delta = end - start;
          if (Math.abs(delta) > largestShift.absDelta) {
            largestShift = { label, delta, start, end, absDelta: Math.abs(delta) };
          }

          let ageTotal = 0;
          years.forEach(year => {
            const value = cleanNumber(panel.yearly[year]?.[idx]);
            ageTotal += value;
            maxCell = Math.max(maxCell, value);
          });
          if (ageTotal > dominantAge.value) dominantAge = { label, value: ageTotal };
        });

        return {
          ...panel,
          maxCell,
          dominantAge: dominantAge.label,
          largestShift,
          score: largestShift.absDelta
        };
      }).sort((a, b) => b.score - a.score || b.total - a.total).slice(0, 8);
    }

    function drawAgeYearPanels(years) {
      const host = document.getElementById("ageYearPanels");
      const summary = document.getElementById("ageYearSummary");
      host.innerHTML = "";
      const panels = buildAgeYearPanels(years);
      if (!panels.length) {
        summary.textContent = "Primary diagnosis summary categories are shown with code + description. Fiscal years run across columns, age groups run down rows, and 3-character / 4-character diagnosis levels are intentionally excluded from this view.";
        host.innerHTML = `<div class="subtle">No age-split primary diagnosis summary categories are available for the selected fiscal years.</div>`;
        return panels;
      }
      const leadPanel = panels[0];
      summary.textContent = `${leadPanel.name} currently shows the strongest age-pattern shift in the selected fiscal years. The biggest change is in age ${leadPanel.largestShift.label}, moving ${signedFmt(leadPanel.largestShift.delta)} admissions from ${fiscal(years[0])} to ${fiscal(years[years.length - 1])}. Categories are shown with code + description, and 3-character / 4-character diagnosis levels are intentionally excluded here.`;

      panels.forEach(panel => {
        const card = document.createElement("article");
        card.className = "panel-card";

        const head = document.createElement("div");
        head.className = "panel-head";

        const title = document.createElement("div");
        title.className = "panel-title";
        title.textContent = panel.name;

        const meta = document.createElement("div");
        meta.className = "panel-meta";
        meta.textContent = `${panel.chapter} · total admissions ${fmt(panel.total)} · largest shift ${panel.largestShift.label} ${signedFmt(panel.largestShift.delta)} from ${fiscal(years[0])} to ${fiscal(years[years.length - 1])}`;

        head.appendChild(title);
        head.appendChild(meta);
        card.appendChild(head);

        const svg = node("svg", { class: "panel-svg", viewBox: "0 0 360 196", preserveAspectRatio: "xMidYMid meet" });
        const plot = { left: 62, top: 12, right: 348, bottom: 136 };
        const cellW = (plot.right - plot.left) / Math.max(1, years.length);
        const cellH = (plot.bottom - plot.top) / ageKeys.length;
        const yearStep = Math.max(1, Math.ceil(years.length / 6));

        ageLabels.forEach((label, rowIdx) => {
          const y = plot.top + rowIdx * cellH;
          svg.appendChild(node("text", { class: "small-label", x: plot.left - 8, y: y + cellH * 0.62, "text-anchor": "end" }, label));
        });

        years.forEach((year, colIdx) => {
          const x = plot.left + colIdx * cellW;
          ageKeys.forEach((key, rowIdx) => {
            const y = plot.top + rowIdx * cellH;
            const value = cleanNumber(panel.yearly[year]?.[rowIdx]);
            svg.appendChild(node("rect", {
              x,
              y,
              width: Math.max(1, cellW - 1),
              height: Math.max(1, cellH - 1),
              fill: colorScale(value, panel.maxCell || 1, [252, 248, 241], [31, 78, 121])
            }));
          });

          if (colIdx % yearStep === 0 || colIdx === years.length - 1) {
            const label = node("text", {
              class: "small-label",
              x: x + cellW / 2,
              y: plot.bottom + 16,
              "text-anchor": "middle",
              transform: `rotate(-35 ${x + cellW / 2} ${plot.bottom + 16})`
            }, fiscal(year));
            svg.appendChild(label);
          }
        });

        const caption = node("text", { class: "small-label", x: 12, y: 182 }, `Panel scale: darker = more age-specific admissions within this category`);
        svg.appendChild(caption);
        card.appendChild(svg);
        host.appendChild(card);
      });

      return panels;
    }

    function update() {
      const years = selectedYears();
      const metric = metricSelect.value;
      if (!years.length) {
        document.querySelectorAll("svg").forEach(clear);
        document.getElementById("ageYearPanels").innerHTML = "";
        document.getElementById("ageYearSummary").textContent = "Primary diagnosis summary categories are shown with code + description. Fiscal years run across columns, age groups run down rows, and 3-character / 4-character diagnosis levels are intentionally excluded from this view.";
        document.getElementById("status").textContent = "No years selected";
        return;
      }
      const chapters = aggregateChapters(years, metric);
      const categories = aggregateCategories(years, metric);
      drawAgeYearPanels(years);
      drawParallel(chapters, metric);
      drawMatrix(years, metric);
      drawRadial(years, metric);
      drawAgeMatrix(categories);
      updateTable(categories, metric);
      document.getElementById("status").textContent = `${fiscal(years[0])} to ${fiscal(years[years.length - 1])} · ${years.length} selected`;
    }

    function init() {
      yearControls.innerHTML = DATA.years.map(year => `
        <label class="year">
          <input class="year-check" type="checkbox" value="${year.year_start}" checked>
          <span>${year.fiscal_year}</span>
        </label>
      `).join("");
      document.querySelectorAll(".year-check").forEach(input => input.addEventListener("change", update));
      metricSelect.addEventListener("change", update);
      document.getElementById("parallelAll").addEventListener("click", () => {
        for (const key of parallelSelection.keys()) parallelSelection.set(key, true);
        update();
      });
      document.getElementById("parallelNone").addEventListener("click", () => {
        for (const key of parallelSelection.keys()) parallelSelection.set(key, false);
        update();
      });
      document.querySelectorAll("button[data-action]").forEach(button => {
        button.addEventListener("click", () => {
          const action = button.dataset.action;
          const maxYear = Math.max(...DATA.years.map(year => year.year_start));
          if (action === "all") setYears(() => true);
          if (action === "none") setYears(() => false);
          if (action === "lockdown") setYears(year => year >= 2019);
          if (action === "latest") setYears(year => year === maxYear);
        });
      });
      window.addEventListener("resize", update);
      update();
    }
    init();
  </script>
</body>
</html>
"""
    path = OUTPUT_DIR / "dashboard.html"
    path.write_text(template.replace("%%DATA%%", json.dumps(payload, ensure_ascii=False)), encoding="utf-8")
    return path


def plot_overview_heatmap(chapter: pd.DataFrame) -> Path:
    metric = "admissions"
    totals = chapter.groupby(["year_start", "fiscal_year"], as_index=False)[metric].sum()
    latest_year = int(chapter["year_start"].max())
    latest_top = (
        chapter[chapter["year_start"].eq(latest_year)]
        .sort_values(metric, ascending=False)
        .head(16)["chapter"]
        .astype(str)
        .tolist()
    )
    heat = chapter[chapter["chapter"].astype(str).isin(latest_top)].pivot(
        index="chapter", columns="fiscal_year", values=metric
    )
    heat = heat.reindex(latest_top)
    indexed = heat.div(heat.replace(0, np.nan).bfill(axis=1).iloc[:, 0], axis=0)
    log_index = np.log2(indexed.replace(0, np.nan))

    fig = plt.figure(figsize=(18, 11), constrained_layout=True)
    gs = fig.add_gridspec(2, 1, height_ratios=[1.05, 3.0])
    ax_top = fig.add_subplot(gs[0, 0])
    ax_heat = fig.add_subplot(gs[1, 0])

    ax_top.plot(
        totals["fiscal_year"],
        totals[metric],
        color="#264653",
        linewidth=2.5,
        marker="o",
        markersize=4,
    )
    pandemic = totals[totals["fiscal_year"].eq("2020-21")]
    if not pandemic.empty:
        ax_top.scatter(pandemic["fiscal_year"], pandemic[metric], s=90, color="#e76f51", zorder=5)
        ax_top.annotate(
            "2020-21 lockdown year",
            xy=(pandemic["fiscal_year"].iloc[0], pandemic[metric].iloc[0]),
            xytext=(15, 18),
            textcoords="offset points",
            arrowprops={"arrowstyle": "->", "color": "#7a4b3b"},
            fontsize=10,
            color="#7a4b3b",
        )
    ax_top.set_title("England hospital admissions, 1998-99 to 2023-24")
    ax_top.set_ylabel("Admissions")
    ax_top.yaxis.set_major_formatter(FuncFormatter(millions))
    ax_top.grid(axis="y")
    ax_top.tick_params(axis="x", rotation=45)

    cmap = LinearSegmentedColormap.from_list(
        "growth",
        ["#8c2d04", "#f1c27d", "#fbfaf7", "#98d8c8", "#1b6f78"],
    )
    image = ax_heat.imshow(log_index, aspect="auto", cmap=cmap, vmin=-1.0, vmax=1.0)
    ax_heat.set_title("Chapter-level change, indexed to each chapter's first observed year")
    ax_heat.set_xticks(np.arange(len(log_index.columns)))
    ax_heat.set_xticklabels(log_index.columns, rotation=45, ha="right")
    ax_heat.set_yticks(np.arange(len(log_index.index)))
    ax_heat.set_yticklabels([wrap_label(str(label), 26) for label in log_index.index])
    ax_heat.tick_params(length=0)
    for spine in ax_heat.spines.values():
        spine.set_visible(False)

    colorbar = fig.colorbar(image, ax=ax_heat, fraction=0.025, pad=0.015)
    colorbar.set_label("log2 index: -1 is half, +1 is double")

    path = OUTPUT_DIR / "01_overview_trends_heatmap.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def detailed_modern_categories(summary: pd.DataFrame, start_year: int = 2012) -> pd.DataFrame:
    data = summary[(summary["code"] != "Total") & (summary["year_start"] >= start_year)].copy()
    data["label"] = data["code"] + " " + data["description"]
    return data


def plot_lockdown_impact(summary: pd.DataFrame) -> Path:
    modern = detailed_modern_categories(summary, 2019)
    before = modern[modern["fiscal_year"].eq("2019-20")][["code", "description", "emergency"]]
    during = modern[modern["fiscal_year"].eq("2020-21")][["code", "emergency"]]
    merged = before.merge(during, on="code", suffixes=("_2019", "_2020"))
    merged = merged[merged["emergency_2019"].fillna(0) >= 1_000].copy()
    merged["change"] = merged["emergency_2020"] - merged["emergency_2019"]
    merged["pct_change"] = merged["change"] / merged["emergency_2019"].replace(0, np.nan)

    decreases = merged.nsmallest(12, "change")
    increases = merged.nlargest(12, "change")
    selected = pd.concat([decreases, increases]).drop_duplicates("code")
    selected = selected.sort_values("change")

    colors = np.where(selected["change"] < 0, "#457b9d", "#e76f51")
    y = np.arange(len(selected))

    fig, ax = plt.subplots(figsize=(14, 11))
    ax.barh(y, selected["change"], color=colors, alpha=0.9)
    ax.axvline(0, color="#2f3136", linewidth=1)
    ax.set_yticks(y)
    ax.set_yticklabels(
        [wrap_label(f"{r.code} {r.description}", 42) for r in selected.itertuples()],
        fontsize=9,
    )
    ax.xaxis.set_major_formatter(FuncFormatter(millions))
    ax.set_xlabel("Change in emergency admissions")
    ax.set_title("Lockdown impact: emergency admissions changed sharply by category")
    ax.grid(axis="x")

    for idx, row in enumerate(selected.itertuples()):
        label = f"{row.pct_change:+.0%}"
        x = row.change
        ax.text(
            x + (8_000 if x >= 0 else -8_000),
            idx,
            label,
            va="center",
            ha="left" if x >= 0 else "right",
            fontsize=9,
            color="#2f3136",
        )
    ax.text(
        0.01,
        0.02,
        "Comparison: 2020-21 versus 2019-20. Categories filtered to at least 1,000 emergency admissions in 2019-20.",
        transform=ax.transAxes,
        fontsize=9,
        color="#5f6368",
    )

    path = OUTPUT_DIR / "02_lockdown_emergency_change.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def plot_emergency_focus(summary: pd.DataFrame) -> Path:
    latest_year = int(summary["year_start"].max())
    latest = summary[(summary["year_start"].eq(latest_year)) & (summary["code"] != "Total")].copy()
    latest["emergency_share"] = latest["emergency"] / latest["admissions"].replace(0, np.nan)
    latest = latest.nlargest(22, "emergency").sort_values("emergency")

    norm = Normalize(vmin=0.15, vmax=0.95)
    cmap = LinearSegmentedColormap.from_list("emergency", ["#7cc8bd", "#f1c27d", "#bc4749"])
    colors = cmap(norm(latest["emergency_share"].fillna(0.15)))

    fig, ax = plt.subplots(figsize=(14, 11))
    y = np.arange(len(latest))
    ax.barh(y, latest["emergency"], color=colors, edgecolor="#fbfaf7", linewidth=0.8)
    ax.set_yticks(y)
    ax.set_yticklabels([short_category(row, 42) for _, row in latest.iterrows()], fontsize=9)
    ax.xaxis.set_major_formatter(FuncFormatter(millions))
    ax.set_xlabel("Emergency admissions")
    ax.set_title(f"Where to focus: largest emergency admission burdens in {fiscal_label(latest_year)}")
    ax.grid(axis="x")

    for idx, row in enumerate(latest.itertuples()):
        ax.text(
            row.emergency + latest["emergency"].max() * 0.012,
            idx,
            f"{row.emergency_share:.0%} emergency",
            va="center",
            fontsize=8.5,
            color="#4a4d52",
        )

    sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([])
    colorbar = fig.colorbar(sm, ax=ax, fraction=0.025, pad=0.012)
    colorbar.set_label("Emergency share of admissions")
    colorbar.ax.yaxis.set_major_formatter(FuncFormatter(pct))

    path = OUTPUT_DIR / "03_emergency_focus_categories.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def plot_age_disparity(summary: pd.DataFrame) -> Path:
    latest_year = int(summary["year_start"].max())
    latest = summary[(summary["year_start"].eq(latest_year)) & (summary["code"] != "Total")].copy()
    age_cols = ["age_0_14", "age_15_59", "age_60_74", "age_75_plus"]
    latest["age_total"] = latest[age_cols].sum(axis=1)
    latest = latest[(latest["age_total"] > 20_000) & latest[age_cols].notna().any(axis=1)].copy()
    shares = latest[age_cols].div(latest["age_total"], axis=0)
    latest["age_disparity"] = shares.max(axis=1) - shares.min(axis=1)
    latest["dominant_age"] = shares.idxmax(axis=1)
    order_map = {"age_0_14": 0, "age_15_59": 1, "age_60_74": 2, "age_75_plus": 3}
    latest["dominant_order"] = latest["dominant_age"].map(order_map)
    selected = latest.sort_values(["dominant_order", "age_disparity"], ascending=[True, False]).head(26)
    matrix = selected[age_cols].div(selected[age_cols].sum(axis=1), axis=0)

    fig, ax = plt.subplots(figsize=(11, 13))
    cmap = LinearSegmentedColormap.from_list(
        "age",
        ["#fbfaf7", "#b8d8d8", "#5ca4a9", "#264653"],
    )
    image = ax.imshow(matrix, aspect="auto", cmap=cmap, vmin=0, vmax=1)
    ax.set_xticks(np.arange(len(age_cols)))
    ax.set_xticklabels(["0-14", "15-59", "60-74", "75+"])
    ax.set_yticks(np.arange(len(selected)))
    ax.set_yticklabels([short_category(row, 45) for _, row in selected.iterrows()], fontsize=8.5)
    ax.set_title(f"Age disparities by admission category in {fiscal_label(latest_year)}")
    ax.tick_params(length=0)
    for spine in ax.spines.values():
        spine.set_visible(False)

    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            value = matrix.iloc[i, j]
            if pd.notna(value):
                ax.text(
                    j,
                    i,
                    f"{value:.0%}",
                    ha="center",
                    va="center",
                    color="#ffffff" if value > 0.55 else "#2f3136",
                    fontsize=8,
                )

    colorbar = fig.colorbar(image, ax=ax, fraction=0.035, pad=0.015)
    colorbar.set_label("Share of category FCEs")
    colorbar.ax.yaxis.set_major_formatter(FuncFormatter(pct))

    path = OUTPUT_DIR / "04_age_disparity_heatmap.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def plot_long_term_change(chapter: pd.DataFrame) -> Path:
    metric = "admissions"
    first_year = int(chapter["year_start"].min())
    latest_year = int(chapter["year_start"].max())
    first = chapter[chapter["year_start"].eq(first_year)][["chapter", metric]]
    latest = chapter[chapter["year_start"].eq(latest_year)][["chapter", metric]]
    merged = first.merge(latest, on="chapter", suffixes=("_first", "_latest"))
    merged = merged[(merged[f"{metric}_first"] > 10_000) | (merged[f"{metric}_latest"] > 10_000)].copy()
    merged["change"] = merged[f"{metric}_latest"] - merged[f"{metric}_first"]
    merged["pct_change"] = merged["change"] / merged[f"{metric}_first"].replace(0, np.nan)
    merged = merged.sort_values("change")

    colors = np.where(merged["change"] < 0, "#457b9d", "#e76f51")
    y = np.arange(len(merged))

    fig, ax = plt.subplots(figsize=(13, 9.5))
    ax.barh(y, merged["change"], color=colors)
    ax.axvline(0, color="#2f3136", linewidth=1)
    ax.set_yticks(y)
    ax.set_yticklabels([wrap_label(str(label), 28) for label in merged["chapter"]])
    ax.xaxis.set_major_formatter(FuncFormatter(millions))
    ax.set_xlabel("Change in admissions")
    ax.set_title(f"Long-term change by ICD-10 chapter: {fiscal_label(first_year)} to {fiscal_label(latest_year)}")
    ax.grid(axis="x")

    for idx, row in enumerate(merged.itertuples()):
        ax.text(
            row.change + (80_000 if row.change >= 0 else -80_000),
            idx,
            f"{row.pct_change:+.0%}",
            va="center",
            ha="left" if row.change >= 0 else "right",
            fontsize=9,
            color="#4a4d52",
        )

    path = OUTPUT_DIR / "05_long_term_chapter_change.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def plot_mental_health(summary: pd.DataFrame, chapter: pd.DataFrame) -> Path:
    mental_total = chapter[chapter["chapter"].astype(str).eq("Mental health")].copy()
    mental_categories = summary[
        (summary["code"] != "Total") & summary["chapter"].eq("Mental health")
    ].copy()
    all_years = sorted(summary["year_start"].unique())
    tick_labels = [fiscal_label(int(year)) for year in all_years]
    detail_years = [year for year in all_years if year >= 2019]
    detail_tick_labels = [fiscal_label(int(year)) for year in detail_years]

    latest_year = int(summary["year_start"].max())
    top_codes = (
        mental_categories[mental_categories["year_start"].eq(latest_year)]
        .nlargest(7, "admissions")["code"]
        .tolist()
    )
    detail = mental_categories[mental_categories["code"].isin(top_codes)].copy()

    fig = plt.figure(figsize=(15, 10), constrained_layout=True)
    gs = fig.add_gridspec(1, 2, width_ratios=[1.15, 1.35])
    ax_total = fig.add_subplot(gs[0, 0])
    ax_detail = fig.add_subplot(gs[0, 1])

    ax_total.fill_between(
        mental_total["year_start"],
        mental_total["admissions"],
        color="#8d6ab8",
        alpha=0.25,
    )
    ax_total.plot(
        mental_total["year_start"],
        mental_total["admissions"],
        color="#6d597a",
        linewidth=2.8,
        marker="o",
        markersize=4,
    )
    ax_total.set_title("Mental health admissions over time")
    ax_total.set_ylabel("Admissions")
    ax_total.yaxis.set_major_formatter(FuncFormatter(millions))
    ax_total.set_xticks(all_years)
    ax_total.set_xticklabels(tick_labels)
    ax_total.tick_params(axis="x", rotation=45)
    ax_total.grid(axis="y")

    for idx, (code, group) in enumerate(detail.groupby("code")):
        group = group.sort_values("year_start")
        desc = clean_text(group["description"].iloc[-1])
        label = f"{code} {desc}"
        series = group.set_index("year_start")["admissions"].reindex(detail_years)
        ax_detail.plot(
            detail_years,
            series,
            linewidth=2.0,
            marker="o",
            markersize=3,
            color=PALETTE[idx % len(PALETTE)],
            label=wrap_label(label, 28),
        )
    ax_detail.set_title("Largest mental health categories since 2019-20")
    ax_detail.yaxis.set_major_formatter(FuncFormatter(millions))
    ax_detail.set_xticks(detail_years)
    ax_detail.set_xticklabels(detail_tick_labels)
    ax_detail.tick_params(axis="x", rotation=45)
    ax_detail.grid(axis="y")
    ax_detail.legend(loc="upper left", bbox_to_anchor=(1.01, 1.0), frameon=False, fontsize=8.5)

    path = OUTPUT_DIR / "06_mental_health_trends.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def code_in_range(code: str, letter: str, start: int, end: int) -> bool:
    parsed = first_icd_code(code)
    if not parsed:
        return False
    code_letter, number = parsed
    return code_letter == letter and start <= number <= end


def infection_proxy_group(row: pd.Series) -> str | None:
    code = clean_code(row["code"])
    description = clean_text(row["description"]).lower()

    if code_in_range(code, "U", 0, 49):
        return "COVID/special infection codes"
    if code_in_range(code, "T", 80, 88) or "complications of surgical" in description:
        return "Hospital/internal-care proxy"
    if (
        chapter_for_code(code) == "Infectious and parasitic"
        or code_in_range(code, "J", 0, 22)
        or code_in_range(code, "L", 0, 8)
        or code_in_range(code, "N", 30, 39)
    ):
        return "Community/external infection proxy"
    return None


def plot_infection_proxy_lockdown(summary: pd.DataFrame) -> Path:
    modern = detailed_modern_categories(summary, 2019)
    modern["infection_proxy"] = modern.apply(infection_proxy_group, axis=1)
    proxy = modern[modern["infection_proxy"].notna()].copy()
    grouped = (
        proxy.groupby(["year_start", "fiscal_year", "infection_proxy"], as_index=False)
        .agg(admissions=("admissions", "sum"), emergency=("emergency", "sum"))
        .sort_values(["infection_proxy", "year_start"])
    )
    baseline = grouped[grouped["fiscal_year"].eq("2019-20")][
        ["infection_proxy", "admissions"]
    ].rename(columns={"admissions": "baseline"})
    grouped = grouped.merge(baseline, on="infection_proxy", how="left")
    grouped["index_2019"] = grouped["admissions"] / grouped["baseline"].replace(0, np.nan)

    order = [
        "Community/external infection proxy",
        "COVID/special infection codes",
        "Hospital/internal-care proxy",
    ]
    colors = {
        "Community/external infection proxy": "#2a9d8f",
        "COVID/special infection codes": "#e76f51",
        "Hospital/internal-care proxy": "#6d597a",
    }

    fig = plt.figure(figsize=(15, 9.2), constrained_layout=False)
    gs = fig.add_gridspec(1, 2, width_ratios=[1.2, 1.0])
    ax_abs = fig.add_subplot(gs[0, 0])
    ax_index = fig.add_subplot(gs[0, 1])
    fig.subplots_adjust(top=0.84, bottom=0.22, left=0.07, right=0.97, wspace=0.14)

    for group_name in order:
        group = grouped[grouped["infection_proxy"].eq(group_name)]
        if group.empty:
            continue
        ax_abs.plot(
            group["year_start"],
            group["admissions"],
            marker="o",
            linewidth=2.6,
            color=colors[group_name],
            label=wrap_label(group_name, 28),
        )
        ax_index.plot(
            group["year_start"],
            group["index_2019"],
            marker="o",
            linewidth=2.6,
            color=colors[group_name],
            label=wrap_label(group_name, 28),
        )

    years = sorted(grouped["year_start"].unique())
    labels = [fiscal_label(int(year)) for year in years]
    for ax in [ax_abs, ax_index]:
        ax.set_xticks(years)
        ax.set_xticklabels(labels, rotation=45, ha="right")
        ax.grid(axis="y")
        ax.axvspan(2020 - 0.35, 2020 + 0.35, color="#e76f51", alpha=0.08, linewidth=0)

    ax_abs.set_title("Admissions by infection-source proxy")
    ax_abs.set_ylabel("Admissions")
    ax_abs.yaxis.set_major_formatter(FuncFormatter(millions))
    ax_abs.legend(frameon=False, loc="upper left")

    ax_index.axhline(1, color="#4a4d52", linewidth=1)
    ax_index.set_title("Indexed to 2019-20")
    ax_index.set_ylabel("2019-20 = 1.0")
    ax_index.yaxis.set_major_formatter(FuncFormatter(lambda x, _pos: f"{x:.1f}x"))

    fig.suptitle(
        "Lockdown infection comparison: external risk fell, COVID surged, internal-care proxy persisted",
        y=0.96,
        fontsize=16,
    )
    note = (
        "Proxy definitions: community/external = A-B infectious, J00-J22 respiratory infection, "
        "L00-L08 skin infection, N30-N39 urinary; hospital/internal-care = T80-T88 and "
        "surgical/medical-care complications. The data does not directly record acquisition source."
    )
    fig.text(
        0.07,
        0.07,
        wrap_label(note, 175),
        fontsize=8.5,
        color="#5f6368",
    )

    path = OUTPUT_DIR / "08_lockdown_infection_proxy_comparison.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def worst_ratio(row: list[float], side: float) -> float:
    if not row:
        return math.inf
    total = sum(row)
    maximum = max(row)
    minimum = min(row)
    if minimum == 0:
        return math.inf
    side_squared = side * side
    return max((side_squared * maximum) / (total * total), (total * total) / (side_squared * minimum))


def layout_row(row: list[float], x: float, y: float, dx: float, dy: float) -> tuple[list[dict[str, float]], float, float, float, float]:
    rects = []
    covered = sum(row)
    if dx >= dy:
        row_height = covered / dx
        rx = x
        for size in row:
            width = size / row_height
            rects.append({"x": rx, "y": y, "dx": width, "dy": row_height})
            rx += width
        y += row_height
        dy -= row_height
    else:
        row_width = covered / dy
        ry = y
        for size in row:
            height = size / row_width
            rects.append({"x": x, "y": ry, "dx": row_width, "dy": height})
            ry += height
        x += row_width
        dx -= row_width
    return rects, x, y, dx, dy


def squarify(sizes: list[float], x: float, y: float, dx: float, dy: float) -> list[dict[str, float]]:
    total = sum(sizes)
    if total <= 0:
        return []
    sizes = [size * dx * dy / total for size in sizes if size > 0]
    rects: list[dict[str, float]] = []
    row: list[float] = []
    remaining = list(sizes)
    while remaining:
        candidate = remaining[0]
        side = min(dx, dy)
        if not row or worst_ratio(row + [candidate], side) <= worst_ratio(row, side):
            row.append(candidate)
            remaining.pop(0)
        else:
            new_rects, x, y, dx, dy = layout_row(row, x, y, dx, dy)
            rects.extend(new_rects)
            row = []
    if row:
        new_rects, x, y, dx, dy = layout_row(row, x, y, dx, dy)
        rects.extend(new_rects)
    return rects


def plot_relative_burden(chapter: pd.DataFrame) -> Path:
    latest_year = int(chapter["year_start"].max())
    latest = chapter[chapter["year_start"].eq(latest_year)].copy()
    latest = latest[latest["admissions"].fillna(0) > 0].sort_values("admissions", ascending=False)
    latest["share"] = latest["admissions"] / latest["admissions"].sum()
    rects = squarify(latest["admissions"].tolist(), 0, 0, 100, 60)

    fig, ax = plt.subplots(figsize=(16, 9))
    ax.set_title(f"Relative burden of admissions by ICD-10 chapter in {fiscal_label(latest_year)}")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 60)
    ax.axis("off")

    for idx, (rect, row) in enumerate(zip(rects, latest.itertuples())):
        color = PALETTE[idx % len(PALETTE)]
        patch = patches.Rectangle(
            (rect["x"], rect["y"]),
            rect["dx"],
            rect["dy"],
            facecolor=color,
            edgecolor="#fbfaf7",
            linewidth=2,
            alpha=0.9,
        )
        ax.add_patch(patch)
        area = rect["dx"] * rect["dy"]
        if area > 115:
            label = f"{row.chapter}\n{row.share:.1%}\n{millions(row.admissions)}"
            ax.text(
                rect["x"] + rect["dx"] * 0.04,
                rect["y"] + rect["dy"] * 0.08,
                wrap_label(label, 20),
                ha="left",
                va="bottom",
                fontsize=9 if area > 230 else 8,
                color="#ffffff",
                weight="bold" if area > 230 else "normal",
            )

    path = OUTPUT_DIR / "07_relative_burden_treemap.png"
    fig.savefig(path, bbox_inches="tight", dpi=220)
    plt.close(fig)
    return path


def build_visuals(summary: pd.DataFrame, chapter: pd.DataFrame) -> list[Path]:
    return [
        plot_overview_heatmap(chapter),
        plot_lockdown_impact(summary),
        plot_emergency_focus(summary),
        plot_age_disparity(summary),
        plot_long_term_change(chapter),
        plot_mental_health(summary, chapter),
        plot_infection_proxy_lockdown(summary),
    ]


def main() -> None:
    configure_style()
    OUTPUT_DIR.mkdir(exist_ok=True)

    summary = load_summary_data()
    chapter = make_chapter_data(summary)
    save_tables(summary, chapter)
    paths = build_visuals(summary, chapter)
    dashboard_path = build_advanced_dashboard(summary, chapter)

    print(f"Loaded {summary['year_start'].nunique()} years and {len(summary):,} category-year rows.")
    print(f"Saved cleaned tables and {len(paths)} visualizations to: {OUTPUT_DIR}")
    for path in paths:
        print(f"- {path.name}")
    print(f"- {dashboard_path.name}")


if __name__ == "__main__":
    main()

