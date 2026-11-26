import argparse, logging
from pathlib import Path
import collector
from render import build
ROOT = Path(__file__).resolve().parent
parser = argparse.ArgumentParser()
parser.add_argument('--render-only', action='store_true')
args = parser.parse_args()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('cloud')
if not args.render_only:
    readings = collector.collect_all(logger)
    expected = {(loc, measure) for _, loc, measure, _ in collector.BASE_ITEMS}
    if len(readings) != 10 or {(r.location, r.measurement) for r in readings} != expected:
        raise RuntimeError('Incomplete or unexpected station set')
    good = sum(r.temperature_c is not None and r.status == 'OK' for r in readings)
    print(f'Valid temperatures: {good}/10')
    for r in readings:
        print(r.location, r.measurement, r.temperature_c, r.status, r.wind_speed, r.wind_direction, r.wind_status)
    if not good:
        raise RuntimeError('All temperatures failed; preserving existing history')
    collector.append_csv(ROOT, readings, collector.local_now(), logger)
html, n = build(ROOT / collector.NEW_CSV)
out = ROOT / 'site'
out.mkdir(exist_ok=True)
(out / 'index.html').write_text(html, encoding='utf-8')
print(f'Rendered {n} rows')
