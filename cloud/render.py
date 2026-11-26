import csv, json, math, re
from pathlib import Path
from datetime import datetime, timezone, timedelta
ROOT=Path(__file__).resolve().parent
def build(path):
    def num(s):
        try:
            v=float(s); return v if math.isfinite(v) else None
        except (ValueError,TypeError): return None
    rows={}
    with path.open(encoding='utf-8-sig',newline='') as f:
        for r in csv.DictReader(f):
            t=r['Collected At (UK)']; t=t+' GMT' if len(t)==19 and t[:7]=='2026-11' else t; m=re.fullmatch(r'(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d) (GMT|BST)',t)
            if not m: raise ValueError('Unexpected timestamp: '+t)
            dt=datetime.fromisoformat(m[1]).replace(tzinfo=timezone(timedelta(hours=int(m[2]=='BST')))); stamp=dt.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
            row=[stamp,r['Location'],r['Measurement'],num(r['Temperature (°C)']),r['Status'],num(r['Wind Speed (mph)']),r['Wind Direction'],num(r['Wind Gust (mph)']),r['Wind Label'],'',r['Wind Status']]
            if row[1] in ("Aonach Mor (650m)", "Aonach Mor (1185m)"):
                row[5:9] = [None, "", None, ""]
                row[10] = "Not available"
            rows[tuple(row[:3])]=row
    data=sorted(rows.values(),key=lambda r:r[:3]);
    if not data: raise ValueError('No data')
    if len([r for r in data if r[0]==data[-1][0]])!=10: raise ValueError('Latest collection incomplete; retry later')
    return (ROOT/'template.html').read_text(encoding='utf-8').replace('__OBSERVATIONS__',json.dumps(data,ensure_ascii=True,separators=(',',':')).replace('<','\\u003c')),len(data)
