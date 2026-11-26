#!/usr/bin/env python3
"""Scottish Highlands Weather Monitor

Collects selected public weather observations and appends them to an Excel
workbook in OneDrive. Designed for Windows 11 + Microsoft 365 desktop Excel.

Holfuy page collection is enabled on the user's confirmation that permission
for this low-frequency, personal collection has been granted. The collector
reads the latest visible Temp. (°C) sample from each public station page and
records the page's refresh time and sample time where available.
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import sys
import time
from dataclasses import dataclass, asdict
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup

UK = ZoneInfo("Europe/London")
APP_NAME = "Scottish Highlands Weather Monitor"
USER_AGENT = "ScottishHighlandsWeatherMonitor/1.9.2 (personal weather log; low-frequency polling)"
TIMEOUT = 30

URL_TRAFFIC = "https://www.traffic.gov.scot/traffic-information/weather-stations"
URL_CAIRNGORM = "https://www.cairngormmountain.co.uk/mountain-report/"
URL_CIC = "https://holfuy.com/en/weather/1746"
URL_AONACH_650 = "https://holfuy.com/en/weather/296"
URL_AONACH_1185 = "https://holfuy.com/en/weather/1365"
URL_DRUMOCHTER = "https://holfuy.com/en/weather/1531"
URL_AVIEMORE = "https://wstatmap.com/stations/aviemore-automatic-w-03063"

BASE_ITEMS = [
    ("Tyndrum – Air", "A82 Tyndrum", "Air Temperature", URL_TRAFFIC),
    ("Tyndrum – Road", "A82 Tyndrum", "Road Surface Temperature", URL_TRAFFIC),
    ("Spean Bridge – Air", "A82 Spean Bridge", "Air Temperature", URL_TRAFFIC),
    ("Spean Bridge – Road", "A82 Spean Bridge", "Road Surface Temperature", URL_TRAFFIC),
    ("Cairngorm – Base", "Cairngorm Mountain", "Base Station Temperature", URL_CAIRNGORM),
    ("CIC Hut", "CIC Hut", "Air Temperature", URL_CIC),
    ("Aonach Mor (650m)", "Aonach Mor (650m)", "Air Temperature", URL_AONACH_650),
    ("Aonach Mor (1185m)", "Aonach Mor (1185m)", "Air Temperature", URL_AONACH_1185),
    ("Drumochter Pass", "Drumochter Pass", "Air Temperature", URL_DRUMOCHTER),
    ("Aviemore", "Aviemore", "Air Temperature", URL_AVIEMORE),
]

ELEVATIONS = dict(zip(["A82 Tyndrum", "A82 Spean Bridge", "Cairngorm Mountain", "CIC Hut", "Aonach Mor (650m)", "Aonach Mor (1185m)", "Drumochter Pass", "Aviemore"], [340,60,1245,650,650,1185,460,318]))
EXTRA = ["Elevation (m)", "Wind Speed (mph)", "Wind Direction", "Wind Gust (mph)", "Wind Label", "Wind Status"]
COMPASS = "N NNE NE ENE E ESE SE SSE S SSW SW WSW W WNW NW NNW".split()
def elevation(loc):
    return ELEVATIONS.get("Aonach Mor (650m)" if loc == "Aonach Mor" else loc)
def titled(loc):
    return loc if loc.endswith("m)") else f"{loc} ({elevation(loc)}m)"
def wind_label(speed, direction):
    return "0" if speed == 0 else f"{int(speed+0.5)}{direction}" if speed is not None and direction else None

def _parse_wind(reading, html, kind):
    soup = BeautifulSoup(html, "html.parser")
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    reading.wind_speed = reading.wind_direction = reading.wind_gust = None
    reading.wind_status = "Unavailable"
    d = r"(" + "|".join(sorted(COMPASS,key=len,reverse=True)) + r")\b"
    n = r"(\d+(?:\.\d+)?)"
    if kind == "holfuy":
        if reading.temperature_c is None:
            return reading
        for table in soup.find_all("table"):
            rows = {}
            for row in table.find_all("tr"):
                cells = [c.get_text(" ", strip=True) for c in row.find_all(["td", "th"], recursive=False)]
                if cells:
                    rows[_normalise_holfuy_label(cells[0])] = cells[1:]
            temps = next((v for k,v in rows.items() if re.fullmatch(r"temp\.?\s*\(°c\)", k)), None)
            if temps is None:
                continue
            # Only the first (15-minute) table; never fall back to older hourly data.
            speeds = rows.get("speed", [])
            gusts = rows.get("gust (mph)", [])
            directions = rows.get("direction deg.", [])
            times = rows.get("time", [])
            if not (len(temps)==len(speeds)==len(gusts)==len(directions)==len(times)) or not temps:
                return reading
            # Explicit mph unit in the table, and the latest temperature column only.
            i = len(temps)-1
            if parse_temp_value(temps[i]) != reading.temperature_c:
                return reading
            if not re.fullmatch(n, speeds[i]):
                return reading
            direction = re.match(d, directions[i], re.I)
            speed = float(speeds[i])
            if not direction and speed != 0:
                return reading
            # Reject old or ambiguous observation times relative to collection clock.
            tm = re.fullmatch(r"(\d{1,2})(?:h|:(\d{2}))", times[i])
            updated = re.search(r"Averages updated:\s*(\d{2}):(\d{2}):(\d{2})\s+\w+\s+(GMT|BST)\b", text)
            if not tm or not updated:
                return reading
            hour, minute = int(tm[1]), int(tm[2] or 0)
            if hour > 23 or minute > 59:
                return reading
            now = local_now()
            zone = timezone(timedelta(hours=1 if updated[4]=="BST" else 0))
            clock = now.astimezone(zone)
            try:
                page_clock = clock.replace(hour=int(updated[1]), minute=int(updated[2]), second=int(updated[3]), microsecond=0)
            except ValueError:
                return reading
            if page_clock > clock + timedelta(minutes=2):
                page_clock -= timedelta(days=1)
            if not -120 <= (clock-page_clock).total_seconds() <= 3600:
                reading.wind_status = "Stale"; return reading
            sample = clock.replace(hour=hour, minute=minute, second=0, microsecond=0)
            if sample > clock + timedelta(minutes=2):
                sample -= timedelta(days=1)
            if not -120 <= (clock-sample).total_seconds() <= 3600:
                reading.wind_status = "Stale"; return reading
            reading.wind_speed = speed
            reading.wind_direction = direction[1].upper() if direction else None
            reading.wind_gust = float(gusts[i]) if re.fullmatch(n, gusts[i]) else None
            break
    elif kind == "cairngorm":
        marker = re.search(r"Summit Weather Conditions", text, re.I)
        section = text[marker.end():marker.end()+500] if marker else ""
        m = re.search(n+r"\s*mph\s*Wind Speed",section,re.I)
        direction = re.search(d+r"\s*Wind Direction",section,re.I)
        gust = re.search(n+r"\s*mph\s*Wind Gust",section,re.I)
        if m and direction and reading.temperature_c is not None:
            reading.wind_speed,reading.wind_direction=float(m[1]),direction[1].upper()
            reading.wind_gust=float(gust[1]) if gust else None
    elif kind == "aviemore":
        m = re.search(r"\bSpeed\s*"+n+r"\s*km/h",text,re.I)
        direction = re.search(r"\bFrom\s*[▲△▼▴]?\s*"+d,text,re.I)
        if m and direction:
            reading.wind_speed,reading.wind_direction=float(m[1])/1.609344,direction[1].upper()
    if reading.wind_speed is not None and 0 <= reading.wind_speed <= 250:
        reading.wind_status = "OK"
    else:
        reading.wind_speed = reading.wind_direction = reading.wind_gust = None
    return reading

def add_wind(reading, html, kind):
    if reading.location in ("Aonach Mor (650m)", "Aonach Mor (1185m)"):
        reading.wind_speed = reading.wind_direction = reading.wind_gust = None
        reading.wind_status = "Not available"
        return reading
    try:
        result = _parse_wind(reading, html, kind)
        logging.getLogger("highlands_weather").info("Wind %s: %s speed=%s direction=%s gust=%s", reading.location, result.wind_status, result.wind_speed, result.wind_direction, result.wind_gust)
        return result
    except Exception:
        logging.getLogger("highlands_weather").exception("Wind parsing failed for %s", reading.location)
        reading.wind_speed = reading.wind_direction = reading.wind_gust = None
        reading.wind_status = "Error"
        return reading

def extras(r):
    return [None,r.wind_speed,r.wind_direction,r.wind_gust,wind_label(r.wind_speed,r.wind_direction),r.wind_status]

@dataclass
class Reading:
    display_label: str
    location: str
    measurement: str
    temperature_c: float | None
    source_updated_at: str
    status: str
    source_url: str
    notes: str = ""
    wind_speed: float | None = None
    wind_direction: str | None = None
    wind_gust: float | None = None
    wind_status: str = "Unavailable"


def local_now() -> datetime:
    return datetime.now(UK)


def as_excel_datetime(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    if value.tzinfo is not None:
        return value.astimezone(UK).replace(tzinfo=None)
    return value


def setup_logging(output_dir: Path) -> logging.Logger:
    output_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("highlands_weather")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = logging.FileHandler(output_dir / "collector.log", encoding="utf-8")
    file_handler.setFormatter(formatter)
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(console)
    return logger


def load_config() -> dict[str, Any]:
    cfg_path = Path(__file__).with_name("config.json")
    if not cfg_path.exists():
        raise FileNotFoundError(
            f"Missing {cfg_path}. Run INSTALL.bat first so the OneDrive workbook location is configured."
        )
    with cfg_path.open("r", encoding="utf-8") as f:
        config = json.load(f)
    if not config.get("workbook_path"):
        raise ValueError("config.json has no workbook_path. Run setup_local.py again.")
    return config


def get_output_dir(config: dict[str, Any]) -> Path:
    workbook = Path(config["workbook_path"]).expanduser()
    return workbook.parent


def http_get(url: str, params: dict[str, str] | None = None) -> requests.Response:
    headers = {"User-Agent": USER_AGENT, "Accept": "text/html,application/json,*/*"}
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            response = requests.get(url, params=params, headers=headers, timeout=TIMEOUT)
            response.raise_for_status()
            return response
        except Exception as exc:
            last_error = exc
            if attempt < 2:
                time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Request failed after 3 attempts: {last_error}")


def parse_temp_value(raw: str) -> float | None:
    cleaned = raw.strip().replace("−", "-")
    if not cleaned or "not available" in cleaned.lower() or cleaned in {"--", "—", "-"}:
        return None
    match = re.search(r"[-+]?\d+(?:\.\d+)?", cleaned)
    return float(match.group(0)) if match else None


def parse_traffic_scotland(html: str, collected_at: datetime) -> list[Reading]:
    soup = BeautifulSoup(html, "html.parser")
    text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    results: list[Reading] = []
    for display, location, measurement, url in BASE_ITEMS[:4]:
        station = "A82 Tyndrum" if "Tyndrum" in location else "A82 Spean Bridge"
        start = text.find(station)
        if start < 0:
            seg = ""
        else:
            seg = text[start:start + 1400]
        updated_match = re.search(r"Last updated:\s*(.+?)\s+Road\s+Road surface state", seg, re.I)
        if not updated_match:
            updated_match = re.search(r"Last updated:\s*(.+?)\s+Road surface state", seg, re.I)
        updated = updated_match.group(1).strip() if updated_match else ""
        if measurement == "Road Surface Temperature":
            temp_match = re.search(
                r"Road surface temperature\s+(Not Available|[-+]?\d+(?:\.\d+)?)\s*°?\s*C", seg, re.I
            )
        else:
            temp_match = re.search(
                r"Air temperature\s+(Not Available|[-+]?\d+(?:\.\d+)?)\s*°?\s*C", seg, re.I
            )
        temp = parse_temp_value(temp_match.group(1)) if temp_match else None
        status = "OK" if temp is not None else "Unavailable"
        note = "" if temp is not None else "Station value was not found or is marked Not Available on the source page."
        results.append(Reading(display, location, measurement, temp, updated, status, url, note))
    return results


def parse_cairngorm(html: str) -> Reading:
    soup = BeautifulSoup(html, "html.parser")
    raw = str(soup)
    marker = re.search(r"Summit Weather Conditions", raw, re.I)
    tail = raw[marker.end():marker.end() + 8000] if marker else ""
    # Cairngorm's weather icon alt text usually contains e.g. "Cloudy (-5°)".
    temp_match = re.search(r"\(\s*([-+]?\d+(?:\.\d+)?)\s*°\s*\)", tail)
    if not temp_match:
        temp_match = re.search(r"([-+]?\d+(?:\.\d+)?)\s*°\s*(?:Sunny|Cloudy|Rain|Snow|Overcast|Clear)", tail, re.I)
    temp = float(temp_match.group(1)) if temp_match else None
    updated_match = re.search(r"Last updated\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4}\s+[0-9:]+\s*[ap]m)", raw, re.I)
    updated = updated_match.group(1).strip() if updated_match else ""
    status = "OK" if temp is not None else "Unavailable"
    note = "" if temp is not None else "Could not identify the Summit temperature in the report HTML."
    return Reading("Cairngorm – Base", "Cairngorm Mountain", "Base Station Temperature", temp,
                   updated, status, URL_CAIRNGORM, note)


def parse_aviemore(html: str) -> Reading:
    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    # The WStatMap station page's SEO/page title currently publishes a "... -3.8°C now" value.
    match = re.search(r"([-+]?\d+(?:\.\d+)?)\s*°\s*C\s*now\b", title, re.I)
    if not match:
        # Best-effort fallback if the page moves the live value into visible text.
        text = soup.get_text(" ", strip=True)
        match = re.search(r"([-+]?\d+(?:\.\d+)?)\s*°\s*C\s*now\b", text, re.I)
    temp = float(match.group(1)) if match else None
    status = "OK" if temp is not None else "Unavailable"
    note = "" if temp is not None else "No live temperature was found in the page title/text; the site may have changed its rendering."
    return Reading("Aviemore", "Aviemore", "Air Temperature", temp, "", status, URL_AVIEMORE, note)


def _normalise_holfuy_label(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("−", "-")).strip().lower()


def parse_holfuy_page(html: str, label: str, location: str, url: str) -> Reading:
    """Read the most recent visible sample in the first Holfuy Temp. (°C) row.

    Holfuy exposes a recent observations table on the station page. The first
    temperature row is the latest time window; values run left-to-right from
    older to newer, so the last populated numeric cell is the latest sample.
    """
    soup = BeautifulSoup(html, "html.parser")
    page_text = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))

    # Holfuy currently labels the observation row "Temp. (°C)". Accept minor
    # punctuation/spacing variations without matching Wetbulb temperature.
    temp_label_re = re.compile(
        r"^(?:temp(?:erature)?\.?\s*\(?\s*°\s*c\s*\)?|temp\.?\s*\(?\s*c\s*\)?)$",
        re.I,
    )
    temp_row = None
    time_row = None
    latest_value: float | None = None
    sample_time = ""

    for row in soup.find_all("tr"):
        cells = row.find_all(["th", "td"])
        if not cells:
            continue
        first_cell = _normalise_holfuy_label(cells[0].get_text(" ", strip=True))
        if temp_label_re.fullmatch(first_cell):
            temp_row = row
            parent_table = row.find_parent("table")
            if parent_table is not None:
                for candidate in parent_table.find_all("tr"):
                    candidate_cells = candidate.find_all(["th", "td"])
                    if candidate_cells and _normalise_holfuy_label(
                        candidate_cells[0].get_text(" ", strip=True)
                    ) == "time":
                        time_row = candidate
                        break
            values = [parse_temp_value(cell.get_text(" ", strip=True)) for cell in cells[1:]]
            populated = [value for value in values if value is not None]
            if populated:
                latest_value = populated[-1]
            break

    if temp_row is not None and time_row is not None:
        time_cells = time_row.find_all(["th", "td"])[1:]
        times = [cell.get_text(" ", strip=True) for cell in time_cells]
        nonempty_times = [value for value in times if value]
        if nonempty_times:
            sample_time = nonempty_times[-1]

    # Fallback for page markup where the observation rows are represented by
    # divs rather than table rows. Read numeric samples after the visible label.
    if latest_value is None and temp_row is None:
        marker = re.search(r"Temp\.?\s*\(\s*°\s*C\s*\)", page_text, re.I)
        if marker:
            tail = page_text[marker.end():]
            # The newest observations section ends before the next historical
            # block or the graph section. The last number in this short region
            # is the latest sample in the row.
            tail = re.split(r"\s(?:\*\s*){2,}|\sGraphs\b", tail, maxsplit=1, flags=re.I)[0]
            token_matches = re.findall(r"(?<![A-Za-z])[-+]?(?:\d+\.\d+|\d+)(?![A-Za-z])", tail)
            if token_matches:
                latest_value = parse_temp_value(token_matches[-1])

    # Avoid accepting an unrelated number if the site markup changes radically.
    if latest_value is not None and not (-80.0 <= latest_value <= 60.0):
        latest_value = None

    updated_match = re.search(
        r"Averages updated:\s*(.+?)(?=\s+Time\b|\s+Speed\b|$)", page_text, re.I
    )
    page_updated = updated_match.group(1).strip() if updated_match else ""
    updated_parts = []
    if page_updated:
        updated_parts.append(f"Averages updated: {page_updated}")
    if sample_time:
        updated_parts.append(f"Latest sample: {sample_time}")
    source_updated_at = "; ".join(updated_parts)

    status = "OK" if latest_value is not None else "Unavailable"
    if latest_value is not None:
        note = "Latest populated value in the visible Holfuy Temp. (°C) observations row."
    else:
        note = "Could not locate a populated Temp. (°C) observations row on the Holfuy page."
    return Reading(label, location, "Air Temperature", latest_value,
                   source_updated_at, status, url, note)


def holfuy_records(collected_at: datetime, logger: logging.Logger) -> list[Reading]:
    stations = [
        ("CIC Hut", "CIC Hut", URL_CIC),
        ("Aonach Mor (650m)", "Aonach Mor (650m)", URL_AONACH_650),
        ("Aonach Mor (1185m)", "Aonach Mor (1185m)", URL_AONACH_1185),
        ("Drumochter Pass", "Drumochter Pass", URL_DRUMOCHTER),
    ]
    results: list[Reading] = []
    for label, location, url in stations:
        try:
            response = http_get(url)
            reading = parse_holfuy_page(response.text, label, location, url)
            results.append(add_wind(reading, response.text, "holfuy"))
            logger.info("Holfuy %s: %s%s", label, reading.status,
                        f" ({reading.temperature_c:.1f}°C)" if reading.temperature_c is not None else "")
        except Exception as exc:
            logger.exception("Holfuy collection failed for %s", label)
            results.append(missing_reading(label, location, "Air Temperature", url, exc))
        # Keep the polling polite and avoid firing all four requests at once.
        time.sleep(1)
    logger.info("Holfuy pages returned %s of %s requested temperatures.",
                sum(1 for item in results if item.temperature_c is not None), len(results))
    return results

def missing_reading(display: str, location: str, measurement: str, url: str, exc: Exception) -> Reading:
    return Reading(display, location, measurement, None, "", "Error", url, str(exc)[:500])


def collect_all(logger: logging.Logger) -> list[Reading]:
    readings: list[Reading] = []
    collected_at = local_now()

    # Traffic Scotland: one page fetch contains both requested A82 stations.
    try:
        response = http_get(URL_TRAFFIC)
        traffic = parse_traffic_scotland(response.text, collected_at)
        readings.extend(traffic)
        logger.info("Traffic Scotland: %s/%s requested readings found.",
                    sum(1 for x in traffic if x.temperature_c is not None), len(traffic))
    except Exception as exc:
        logger.exception("Traffic Scotland collection failed")
        for display, loc, measure, url in BASE_ITEMS[:4]:
            readings.append(missing_reading(display, loc, measure, url, exc))

    # Cairngorm Mountain.
    try:
        response = http_get(URL_CAIRNGORM)
        item = add_wind(parse_cairngorm(response.text), response.text, "cairngorm")
        readings.append(item)
        logger.info("Cairngorm Mountain: %s", item.status)
    except Exception as exc:
        logger.exception("Cairngorm collection failed")
        readings.append(missing_reading("Cairngorm – Base", "Cairngorm Mountain",
                                        "Base Station Temperature", URL_CAIRNGORM, exc))

    # Holfuy: the user confirmed permission for low-frequency page collection.
    readings.extend(holfuy_records(collected_at, logger))

    # WStatMap Aviemore station page.
    try:
        response = http_get(URL_AVIEMORE)
        item = add_wind(parse_aviemore(response.text), response.text, "aviemore")
        readings.append(item)
        logger.info("WStatMap Aviemore: %s", item.status)
    except Exception as exc:
        logger.exception("Aviemore collection failed")
        readings.append(missing_reading("Aviemore", "Aviemore", "Air Temperature", URL_AVIEMORE, exc))

    for reading in readings:
        if reading.location in ("Aonach Mor (650m)", "Aonach Mor (1185m)"):
            reading.wind_speed = reading.wind_direction = reading.wind_gust = None
            reading.wind_status = "Not available"

    # Keep the fixed row order used by the Excel dashboard.
    order = {(loc, measure): idx for idx, (_, loc, measure, _) in enumerate(BASE_ITEMS)}
    readings.sort(key=lambda r: order.get((r.location, r.measurement), 99))
    return readings


CSV_FIELDS = ["Collected At (UK)", "Location", "Measurement", "Temperature (°C)",
              "Source Updated At", "Status", "Source URL", "Notes"] + EXTRA
NEW_CSV = "Temperature_Raw_Log_v1_9.csv"

def append_rows(path, rows, logger):
    existing = set()
    if path.exists() and path.stat().st_size:
        with path.open(encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames != CSV_FIELDS:
                raise RuntimeError("Unexpected CSV header: " + str(path))
            for row in reader:
                existing.add(_history_key(row[CSV_FIELDS[0]], row[CSV_FIELDS[1]], row[CSV_FIELDS[2]]))
    new = not path.exists() or not path.stat().st_size
    count = 0
    with path.open("a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
        writer = csv.writer(f)
        if new: writer.writerow(CSV_FIELDS)
        for row in rows:
            key = _history_key(*row[:3])
            if key is None or not key[1] or not key[2] or key in existing: continue
            writer.writerow(row)
            existing.add(key)
            count += 1
    logger.info("CSV backup: %s new rows appended to %s", count, path)

def append_csv(output_dir, readings, collected_at, logger):
    append_rows(output_dir / NEW_CSV, ([collected_at.strftime("%Y-%m-%d %H:%M:%S %Z"), r.location, r.measurement, r.temperature_c, r.source_updated_at, r.status, r.source_url, r.notes, *extras(r)] for r in readings), logger)

def recover_csv(table, path, logger):
    rows = []
    for row in _as_com_rows(table.DataBodyRange.Value2):
        if len(row) < 8: continue
        dt = _normalise_history_datetime(row[0])
        if dt is None: continue
        rows.append([dt.strftime("%Y-%m-%d %H:%M:%S"), *row[1:8], *([None] + list(row[9:14]))])
    append_rows(path.parent / NEW_CSV, rows, logger)


def _try_set_excel_property(excel_app: Any, property_name: str, value: Any,
                            logger: logging.Logger) -> None:
    """Set a non-essential Excel COM property without aborting a collection.

    Some Excel COM instances returned by GetActiveObject expose a dynamic
    dispatch wrapper that rejects writes to properties such as DisplayAlerts.
    These settings are helpful but must not prevent valid readings being saved.
    """
    try:
        setattr(excel_app, property_name, value)
    except Exception as exc:
        logger.warning(
            "Could not set Excel.Application.%s; continuing with Excel defaults: %s",
            property_name,
            exc,
        )


def _normalise_history_datetime(value: Any) -> datetime | None:
    """Convert CSV/Excel timestamp values to a naive UK-local datetime."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.astimezone(UK).replace(tzinfo=None) if value.tzinfo else value
    if isinstance(value, (int, float)):
        # Excel's 1900 date system uses 1899-12-30 as its serial-date epoch.
        try:
            return datetime(1899, 12, 30) + timedelta(seconds=round(float(value) * 86400))
        except (OverflowError, ValueError):
            return None
    text = str(value).strip()
    if not text:
        return None
    # Collector CSV timestamps are written as YYYY-MM-DD HH:MM:SS GMT/BST.
    if len(text) >= 19:
        try:
            return datetime.strptime(text[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return parsed.astimezone(UK).replace(tzinfo=None) if parsed.tzinfo else parsed
    except ValueError:
        return None


def _history_key(timestamp: Any, location: Any, measurement: Any) -> tuple[str, str, str] | None:
    dt = _normalise_history_datetime(timestamp)
    if dt is None:
        return None
    return (
        dt.strftime("%Y-%m-%d %H:%M:%S"),
        str(location or "").strip(),
        str(measurement or "").strip(),
    )


def _as_com_rows(value: Any) -> list[tuple[Any, ...]]:
    """Normalise an Excel COM range value to a list of row tuples."""
    if value is None:
        return []
    if isinstance(value, tuple):
        if not value:
            return []
        if isinstance(value[0], tuple):
            return [tuple(row) for row in value]
        return [tuple(value)]
    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], (tuple, list)):
            return [tuple(row) for row in value]
        return [tuple(value)]
    return [(value,)]


def _table_next_row(target_table: Any) -> Any:
    """Reuse a blank seed row when present; otherwise append a table row."""
    if target_table.ListRows.Count == 1:
        seed = target_table.ListRows(1)
        current = seed.Range.Value
        rows = _as_com_rows(current)
        if rows and all(value in (None, "") for value in rows[0]):
            return seed
    return target_table.ListRows.Add()


def _csv_history_records(csv_path: Path, logger: logging.Logger) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    if not csv_path.exists():
        logger.warning("Raw CSV history file not found for chart backfill: %s", csv_path)
        return records
    try:
        with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            for item in reader:
                timestamp = _normalise_history_datetime(item.get("Collected At (UK)"))
                location = (item.get("Location") or "").strip()
                measurement = (item.get("Measurement") or "").strip()
                if timestamp is None or not location or not measurement:
                    continue
                raw_temp = (item.get("Temperature (°C)") or "").strip()
                temperature = parse_temp_value(raw_temp) if raw_temp else None
                status = (item.get("Status") or "").strip()
                if not status:
                    status = "OK" if temperature is not None else "Unavailable"
                records.append({
                    "timestamp": timestamp,
                    "location": location,
                    "measurement": measurement,
                    "temperature": temperature,
                    "source_updated_at": item.get("Source Updated At") or "",
                    "status": status,
                    "source_url": item.get("Source URL") or "",
                    "notes": item.get("Notes") or "",
                    "extra": [None, parse_temp_value(item.get(EXTRA[1]) or ""), item.get(EXTRA[2]) or None, parse_temp_value(item.get(EXTRA[3]) or ""), item.get(EXTRA[4]) or None, item.get(EXTRA[5]) or None],
                })
    except Exception:
        logger.exception("Could not read historical CSV for Excel/chart synchronisation: %s", csv_path)
    return records


def _sync_history_into_table(
    table: Any,
    readings: list[Reading],
    collected_at: datetime,
    csv_path: Path,
    logger: logging.Logger,
) -> int:
    """Backfill CSV rows missing from Excel, then ensure the current run is present.

    The CSV is the recovery ledger: early runs whose Excel save failed are
    restored here. A timestamp/location/measurement key prevents duplicates.
    """
    existing_keys: set[tuple[str, str, str]] = set()
    if table.ListRows.Count:
        try:
            current_rows = _as_com_rows(table.DataBodyRange.Value2)
        except Exception:
            raise RuntimeError("Cannot read existing Temperature Log; refusing to append duplicate history")
        for row in current_rows:
            if len(row) >= 3:
                key = _history_key(row[0], row[1], row[2])
                if key:
                    existing_keys.add(key)

    pending = _csv_history_records(csv_path, logger) + _csv_history_records(csv_path.parent / NEW_CSV, logger)
    for reading in readings:
        pending.append({
            "timestamp": as_excel_datetime(collected_at),
            "location": reading.location,
            "measurement": reading.measurement,
            "temperature": reading.temperature_c,
            "source_updated_at": reading.source_updated_at,
            "status": reading.status,
            "source_url": reading.source_url,
            "notes": reading.notes,
            "extra": extras(reading),
        })

    added = 0
    for item in pending:
        timestamp = _normalise_history_datetime(item.get("timestamp"))
        if timestamp is None:
            continue
        key = _history_key(timestamp, item.get("location"), item.get("measurement"))
        if key is None or key in existing_keys:
            continue
        row = _table_next_row(table)
        row.Range.Value2 = ((
            (timestamp - datetime(1899, 12, 30)).total_seconds() / 86400,
            item.get("location", ""),
            item.get("measurement", ""),
            item.get("temperature"),
            item.get("source_updated_at", ""),
            item.get("status", "Unavailable"),
            item.get("source_url", ""),
            item.get("notes", ""),
            *item["extra"],
        ),)
        existing_keys.add(key)
        added += 1

    if added:
        logger.info("Reconciled %s missing historical/current row(s) from CSV into Temperature Log.", added)
    else:
        logger.info("Temperature Log already contains all rows found in the CSV/current run.")
    return added


def _read_history_table(table: Any) -> list[dict[str, Any]]:
    history: list[dict[str, Any]] = []
    if not table.ListRows.Count:
        return history
    try:
        values = _as_com_rows(table.DataBodyRange.Value2)
    except Exception:
        return history
    for row in values:
        if len(row) < 4:
            continue
        timestamp = _normalise_history_datetime(row[0])
        location = str(row[1] or "").strip()
        measurement = str(row[2] or "").strip()
        if timestamp is None or not location or not measurement:
            continue
        try:
            temp = float(row[3]) if row[3] not in (None, "") else None
        except (TypeError, ValueError):
            temp = None
        history.append({
            "timestamp": timestamp,
            "location": location,
            "measurement": measurement,
            "temperature": temp,
            "wind_label": row[12] if len(row)>12 else None,
            "status": str(row[5] or "") if len(row) > 5 else "",
        })
    return history


def _excel_column_letter(column_number: int) -> str:
    """Convert a 1-based Excel column number to its A1 column letters."""
    result = ""
    number = int(column_number)
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _chart_date_interval(span_days: float) -> int:
    """Return the scatter-axis major unit in days (Excel serial units)."""
    for limit, interval in ((7, 1), (14, 2), (30, 5), (90, 10), (180, 30), (365, 60)):
        if span_days <= limit:
            return interval
    return max(60, 30 * int((span_days + 179) // 180))


def refresh_location_charts(workbook: Any, table: Any, logger: logging.Logger) -> None:
    """Rebuild eight location charts from the complete Temperature Log history.

    Chart series names are passed as plain Python strings rather than values
    fetched through Excel COM. Some Excel/pywin32 dynamic-dispatch combinations
    return a VARIANT that Excel rejects when assigned to Series.Name (the
    Type mismatch seen in the user's v1.5 chart rebuild). Each chart/series is
    also guarded so one problematic chart cannot prevent the remaining charts
    or the workbook save.
    """
    history = _read_history_table(table)
    if not history:
        logger.warning("No history rows are available yet; location charts were not created.")
        return

    def canonical_location(location: str) -> str:
        # Earlier logs used 'Aonach Mor' before the 650m/1185m split.
        return "Aonach Mor (650m)" if location == "Aonach Mor" else location

    chart_definitions = [
        ("A82 Tyndrum", ["Air Temperature", "Road Surface Temperature"]),
        ("A82 Spean Bridge", ["Air Temperature", "Road Surface Temperature"]),
        ("Cairngorm Mountain", ["Base Station Temperature"]),
        ("CIC Hut", ["Air Temperature"]),
        ("Aonach Mor (650m)", ["Air Temperature"]),
        ("Aonach Mor (1185m)", ["Air Temperature"]),
        ("Drumochter Pass", ["Air Temperature"]),
        ("Aviemore", ["Air Temperature"]),
    ]
    configured_locations = {location for location, _ in chart_definitions}

    grouped: dict[str, dict[datetime, dict[str, float | None]]] = {}
    for item in history:
        location = canonical_location(item["location"])
        if location not in configured_locations:
            continue
        stamp = item["timestamp"]
        slot = grouped.setdefault(location, {}).setdefault(stamp, {})
        measure = item["measurement"]
        slot[measure] = item["temperature"]
        if item.get("wind_label") not in (None, "") and item["temperature"] is not None: slot["wind_label"] = str(item["wind_label"])

    # Create the visible source-data sheet. Keep one data row for each historic
    # timestamp; unavailable observations stay blank and therefore plot as gaps.
    try:
        data_sheet = workbook.Worksheets("Chart Data")
    except Exception:
        data_sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        data_sheet.Name = "Chart Data"
    data_sheet.Cells.ClearContents()

    chart_blocks: list[dict[str, Any]] = []
    for index, (location, measures) in enumerate(chart_definitions):
        points = grouped.get(location, {})
        timestamps = sorted(points)
        start_col = 1 + index * 5  # up to 5 columns: time, 2 measures, and zero line
        measure_headers = [
            "Road surface (°C)" if measure == "Road Surface Temperature"
            else "Air temperature (°C)" if measure == "Air Temperature"
            else "Base station (°C)"
            for measure in measures
        ]
        headers = ["Collection time"] + measure_headers + ["0°C reference"]
        first_col = _excel_column_letter(start_col)
        end_col = _excel_column_letter(start_col + len(headers) - 1)
        data_sheet.Range(f"{first_col}1:{end_col}1").Value2 = (tuple(headers),)

        # Even an empty station gets a titled chart with an empty point, so the
        # Location Charts sheet always has all eight locations represented.
        rows_to_write: list[list[Any]] = []
        if timestamps:
            for stamp in timestamps:
                row: list[Any] = [(stamp - datetime(1899, 12, 30)).total_seconds() / 86400]
                row.extend(points[stamp].get(measure) for measure in measures)
                row.append(0.0)
                rows_to_write.append(row)
        else:
            rows_to_write.append([None] + [None] * len(measures) + [None])

        last_data_row = 1 + len(rows_to_write)
        target = data_sheet.Range(f"{first_col}2:{end_col}{last_data_row}")
        target.Value2 = tuple(tuple(row) for row in rows_to_write)
        actual = _as_com_rows(target.Value2)
        if len(actual) != len(rows_to_write) or any(
            len(got) != len(expected) or any(
                (abs(float(g)-float(e)) > 1e-7 if isinstance(e, (int,float)) and isinstance(g,(int,float)) else g != e)
                for g,e in zip(got,expected))
            for got,expected in zip(actual,rows_to_write)):
            raise RuntimeError(f"Chart Data write verification failed for {location}")
        data_sheet.Range(f"{first_col}2:{first_col}{last_data_row}").NumberFormat = "dd/mm/yyyy hh:mm"
        data_sheet.Range(f"{_excel_column_letter(start_col+1)}2:{end_col}{last_data_row}").NumberFormat = "0.0"
        logger.info("Verified chart data for %s: %s timestamps, %s numeric temperatures", location, len(timestamps), sum(v is not None for row in rows_to_write for v in row[1:-1]))
        chart_blocks.append({
            "location": location,
            "start_col": start_col,
            "last_row": last_data_row,
            "measures": measures,
            "measure_headers": measure_headers,
            "zero_col": start_col + len(headers) - 1,
            "timestamp_count": len(timestamps),
            "labels": [(i+1, stamp, points[stamp].get("wind_label")) for i,stamp in enumerate(timestamps) if points[stamp].get("wind_label")],
            "span_days": (timestamps[-1] - timestamps[0]).total_seconds() / 86400 if timestamps else 0,
        })

    try:
        data_sheet.Rows(1).Font.Bold = True
        data_sheet.Rows(1).WrapText = True
        data_sheet.Range("A:AN").Columns.AutoFit()
    except Exception as fmt_exc:
        logger.warning("Could not format Chart Data sheet: %s", fmt_exc)

    try:
        chart_sheet = workbook.Worksheets("Location Charts")
    except Exception:
        chart_sheet = workbook.Worksheets.Add(Before=data_sheet)
        chart_sheet.Name = "Location Charts"

    # Remove old/partially-created charts before rebuilding the complete set.
    try:
        chart_objects = chart_sheet.ChartObjects()
        for chart_index in range(chart_objects.Count, 0, -1):
            chart_objects.Item(chart_index).Delete()
    except Exception as clear_exc:
        logger.warning("Could not remove previous location charts: %s", clear_exc)
    chart_sheet.Cells.ClearContents()
    chart_sheet.Range("A1").Value = "Scottish Highlands – Temperature History"
    chart_sheet.Range("A2").Value = (
        "Historical readings from Temperature Log and CSV recovery; blank gaps mean unavailable readings. "
        "Grey line: 0°C. Wind labels: rounded mph + compass direction. Charts refresh after each collection."
    )
    try:
        chart_sheet.Range("A1").Font.Bold = True
        chart_sheet.Range("A1").Font.Size = 18
        chart_sheet.Range("A2").Font.Italic = True
        chart_sheet.Range("A2").Font.Size = 10
        chart_sheet.Columns("A:S").ColumnWidth = 12
    except Exception:
        pass

    # Two charts per row, four rows total.
    chart_width, chart_height = 510, 270
    left_positions = [18, 548]
    top_start, row_gap = 58, 292
    created = 0
    failed = 0
    for chart_index, block in enumerate(chart_blocks):
        row_index, col_index = divmod(chart_index, 2)
        left = left_positions[col_index]
        top = top_start + row_index * row_gap
        try:
            chart_obj = chart_sheet.ChartObjects().Add(left, top, chart_width, chart_height)
            chart = chart_obj.Chart
            chart.ChartType = 74  # xlXYScatterLines
            chart.HasTitle = True
            chart.ChartTitle.Text = titled(block["location"])
            chart.HasLegend = True
            try:
                chart.DisplayBlanksAs = 1  # xlNotPlotted: retain gaps for unavailable readings
            except Exception:
                pass

            while chart.SeriesCollection().Count:
                chart.SeriesCollection(1).Delete()
            # Add measured temperature series. Use literal Python strings for
            # Series.Name to avoid the COM VARIANT type mismatch in v1.5.
            for series_index, header in enumerate(block["measure_headers"]):
                value_col = block["start_col"] + 1 + series_index
                series = chart.SeriesCollection().NewSeries()
                x_col_letter = _excel_column_letter(block["start_col"])
                value_col_letter = _excel_column_letter(value_col)
                sheet_ref = "'" + str(data_sheet.Name).replace("'", "''") + "'!"
                series_name = str(header).replace('"', '""')
                series_formula = (
                    f'=SERIES("{series_name}",{sheet_ref}${x_col_letter}$2:${x_col_letter}${block["last_row"]},'
                    f'{sheet_ref}${value_col_letter}$2:${value_col_letter}${block["last_row"]},{series_index + 1})'
                )
                series.Formula = series_formula
                if value_col_letter not in str(series.Formula):
                    raise RuntimeError("Chart series formula was not retained")
                try:
                    series.Format.Line.Weight = 1.75
                except Exception:
                    pass

                if series_index == 0:
                    span = block["span_days"]
                    hours = 4 if span <= 1 else 12 if span <= 7 else 24 if span <= 14 else 48 if span <= 30 else max(72, span*24/20)
                    last = None
                    for point_index, stamp, label in block["labels"]:
                        if last is not None and (stamp-last).total_seconds() < hours*3600: continue
                        try:
                            point = series.Points(point_index)
                            point.HasDataLabel = True
                            point.DataLabel.Text = label
                            point.DataLabel.Position = 0  # xlLabelPositionAbove
                            point.DataLabel.Font.Size = 8
                            last = stamp
                        except Exception as exc:
                            logger.warning("Wind label failed for %s: %s",block["location"],exc)
                            break

            # Add a distinct 0°C baseline to every location chart.
            zero_series = chart.SeriesCollection().NewSeries()
            zero_col_letter = _excel_column_letter(block["zero_col"])
            zero_formula = (
                f'=SERIES("0°C reference",{sheet_ref}${x_col_letter}$2:${x_col_letter}${block["last_row"]},'
                f'{sheet_ref}${zero_col_letter}$2:${zero_col_letter}${block["last_row"]},{len(block["measures"]) + 1})'
            )
            zero_series.Formula = zero_formula
            try:
                zero_series.Format.Line.Weight = 2.25
                zero_series.Format.Line.ForeColor.RGB = 0x707070
                zero_series.MarkerStyle = -4142  # xlMarkerStyleNone
            except Exception as line_exc:
                logger.warning("Could not format 0°C reference for %s: %s", block["location"], line_exc)

            try:
                # Place the legend under the chart where supported.
                chart.Legend.Position = -4107  # xlLegendPositionBottom
            except Exception:
                pass
            try:
                x_axis = chart.Axes(1)
                x_axis.HasTitle = True
                x_axis.AxisTitle.Text = "Collection time"
                x_axis.TickLabels.NumberFormat = "dd mmm" if block["span_days"] <= 365 else "mmm yy"
                x_axis.MajorUnit = float(_chart_date_interval(block["span_days"]))
                x_axis.TickLabelPosition = -4134  # low
                x_axis.TickLabels.Orientation = 0
                y_axis = chart.Axes(2)
                y_axis.HasTitle = True
                y_axis.AxisTitle.Text = "Temperature (°C)"
                y_axis.TickLabels.NumberFormat = "0.0"
            except Exception as axis_exc:
                logger.warning("Could not fully format axes for %s: %s", block["location"], axis_exc)
            created += 1
            logger.info(
                "Built chart for %s using %s historical timestamp(s).",
                block["location"], block["timestamp_count"],
            )
        except Exception as chart_exc:
            failed += 1
            logger.exception("Could not build chart for %s: %s", block["location"], chart_exc)

    logger.info(
        "Location Charts rebuild finished: %s of %s chart(s) built; %s historical reading row(s); %s chart failure(s).",
        created, len(chart_definitions), len(history), failed,
    )
    if created != len(chart_definitions):
        raise RuntimeError("One or more location charts could not be built; check preceding chart-specific errors.")


def ensure_columns(table, base):
    for i, name in enumerate(EXTRA,base+1):
        if table.ListColumns.Count < i:
            table.ListColumns.Add().Name = name
        elif str(table.ListColumns(i).Name) != name:
            raise RuntimeError(f"Unexpected table column {i}; refusing to overwrite")

def update_excel(workbook_path: Path, readings: list[Reading], collected_at: datetime,
                 started_at: datetime, logger: logging.Logger,
                 csv_path: Path | None = None, charts_only: bool = False) -> None:
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("Excel automation requires pywin32. Run INSTALL.bat again or install requirements.") from exc

    pythoncom.CoInitialize()
    excel = None
    own_excel = False
    workbook = None
    opened_by_script = False
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            logger.info("Connected to the existing Microsoft Excel instance.")
        except Exception:
            excel = win32com.client.DispatchEx("Excel.Application")
            own_excel = True
            _try_set_excel_property(excel, "Visible", False, logger)
            logger.info("Started a separate Microsoft Excel automation instance.")

        _try_set_excel_property(excel, "DisplayAlerts", False, logger)
        _try_set_excel_property(excel, "AskToUpdateLinks", False, logger)

        target = str(workbook_path.resolve())
        for idx in range(1, excel.Workbooks.Count + 1):
            candidate = excel.Workbooks(idx)
            try:
                if os.path.normcase(os.path.abspath(candidate.FullName)) == os.path.normcase(os.path.abspath(target)):
                    workbook = candidate
                    break
            except Exception:
                continue
        if workbook is None:
            workbook = excel.Workbooks.Open(
                target,
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
            )
            opened_by_script = True
        if workbook.ReadOnly:
            raise RuntimeError("The Excel workbook opened as read-only. Close any conflicting copy and retry.")

        log_sheet = workbook.Worksheets("Temperature Log")
        table = log_sheet.ListObjects("TemperatureLog")
        if csv_path is None:
            csv_path = workbook_path.parent / "Temperature_Raw_Log.csv"
        ensure_columns(table, 8)
        _sync_history_into_table(table, readings, collected_at, csv_path, logger)
        table.ListColumns(9).Range.EntireColumn.Hidden = True
        try:
            recover_csv(table, csv_path, logger)
        except Exception:
            logger.exception("CSV recovery failed; continuing with Excel save")

        # This mode lets the user build/backfill the chart tabs without making
        # another live collection or adding a new run-history row.
        if charts_only:
            refresh_location_charts(workbook, table, logger)
            workbook.Save()
            logger.info("Existing history synchronised and location charts saved: %s", workbook_path)
            return

        latest_sheet = workbook.Worksheets("Latest Readings")
        latest_table = workbook.Worksheets("Latest Readings").ListObjects("LatestReadingsTable")
        while latest_table.ListRows.Count < len(BASE_ITEMS):
            latest_table.ListRows.Add()

        ensure_columns(latest_table, 9)
        latest_table.ListColumns(10).Range.EntireColumn.Hidden = True
        metadata = [[label if label.endswith("m)") else f"{label} ({elevation(loc)}m)", loc, measure] for label, loc, measure, _url in BASE_ITEMS]
        latest_sheet.Range(f"A2:C{len(BASE_ITEMS) + 1}").Value = metadata

        key_map = {(r.location, r.measurement): r for r in readings}
        matrix = []
        for label, loc, measure, url in BASE_ITEMS:
            r = key_map.get((loc, measure))
            if r is None:
                matrix.append([None, "", as_excel_datetime(collected_at), "Unavailable", url, "No reading row returned."])
            else:
                matrix.append([
                    r.temperature_c,
                    r.source_updated_at,
                    as_excel_datetime(collected_at),
                    r.status,
                    r.source_url,
                    r.notes,
                ])
        latest_sheet.Range(f"D2:I{len(BASE_ITEMS) + 1}").Value = matrix
        latest_sheet.Range(f"J2:O{len(BASE_ITEMS) + 1}").Value2 = tuple(tuple(extras(key_map[(loc,measure)])) if (loc,measure) in key_map else (None,None,None,None,None,None) for _,loc,measure,_ in BASE_ITEMS)

        last_row = len(BASE_ITEMS) + 1
        dashboard = workbook.Worksheets("Dashboard")
        dashboard.Range("C3").Formula = (
            f'=IF(COUNT(\'Latest Readings\'!F2:F{last_row})=0,"",'
            f'MAX(\'Latest Readings\'!F2:F{last_row}))'
        )
        dashboard.Range("F3").Formula = f'=COUNT(\'Latest Readings\'!D2:D{last_row})'
        dashboard.Range("H3").Formula = (
            f'=IF(COUNT(\'Latest Readings\'!D2:D{last_row})=0,"",'
            f'MIN(\'Latest Readings\'!D2:D{last_row}))'
        )
        try:
            if dashboard.ChartObjects().Count:
                chart_series = dashboard.ChartObjects(1).Chart.SeriesCollection(1)
                chart_series.XValues = latest_sheet.Range(f"A2:A{last_row}")
                chart_series.Values = latest_sheet.Range(f"D2:D{last_row}")
        except Exception as chart_exc:
            logger.warning("Could not refresh dashboard summary chart ranges: %s", chart_exc)

        # Refresh all station-level time-series charts from the full historical table.
        try:
            refresh_location_charts(workbook, table, logger)
        except Exception as chart_exc:
            logger.exception("Location chart refresh failed; continuing with temperature data save: %s", chart_exc)

        successes = sum(1 for r in readings if r.temperature_c is not None and r.status == "OK")
        skipped = len(readings) - successes
        run_sheet = workbook.Worksheets("Run History")
        run_table = run_sheet.ListObjects("RunHistory")
        run_row = _table_next_row(run_table)
        run_details = "; ".join(f"{r.location}: {r.status}" for r in readings if r.status != "OK")
        run_row.Range.Value = [[
            as_excel_datetime(started_at),
            as_excel_datetime(local_now()),
            len(readings),
            successes,
            skipped,
            "Complete" if successes else "No successful readings",
            run_details[:30000],
        ]]

        dashboard.Range("C3").Value = as_excel_datetime(collected_at)
        try:
            excel.CalculateFull()
        except Exception as calc_exc:
            logger.warning("Excel full recalculation could not be completed: %s", calc_exc)

        workbook.Save()
        logger.info("Excel workbook updated: %s", workbook_path)
    finally:
        try:
            if workbook is not None and opened_by_script:
                workbook.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if excel is not None and own_excel:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--scheduled", action="store_true", help="Marker used by Windows Task Scheduler")
    parser.add_argument("--refresh-charts-only", action="store_true", help="Backfill CSV history and rebuild charts without fetching new readings")
    args = parser.parse_args()

    try:
        config = load_config()
        workbook_path = Path(config["workbook_path"]).expanduser()
        output_dir = get_output_dir(config)
        logger = setup_logging(output_dir)
    except Exception as exc:
        print(f"SETUP ERROR: {exc}", file=sys.stderr)
        return 2

    if not workbook_path.exists():
        logger.error("Workbook not found: %s. Run INSTALL.bat again or restore the workbook.", workbook_path)
        return 2

    if args.refresh_charts_only:
        logger.info("Chart-only mode: no new live temperature collection will be made.")
        now = local_now()
        try:
            update_excel(
                workbook_path, [], now, now, logger,
                csv_path=output_dir / "Temperature_Raw_Log.csv",
                charts_only=True,
            )
            return 0
        except Exception as exc:
            logger.exception("Chart rebuild failed: %s", exc)
            return 1

    started_at = local_now()
    run_mode = "scheduled" if args.scheduled else "manual"
    logger.info("Execution mode: %s", run_mode)
    logger.info("Starting temperature collection at %s", started_at.strftime("%Y-%m-%d %H:%M:%S %Z"))
    readings = collect_all(logger)
    collected_at = local_now()

    # CSV is written first as a safety net, even if Excel is temporarily locked.
    try:
        append_csv(output_dir, readings, collected_at, logger)
    except Exception:
        logger.exception("Could not append the CSV backup")

    try:
        update_excel(
            workbook_path, readings, collected_at, started_at, logger,
            csv_path=output_dir / "Temperature_Raw_Log.csv",
        )
    except Exception as exc:
        logger.exception("Excel update failed")
        logger.error("The raw CSV backup may still contain this run. Excel error: %s", exc)
        return 1

    ok_count = sum(1 for r in readings if r.temperature_c is not None and r.status == "OK")
    logger.info("Collection finished: %s/%s temperature values captured.", ok_count, len(readings))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
