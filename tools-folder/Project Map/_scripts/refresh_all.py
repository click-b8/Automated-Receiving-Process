"""Runs both vault refreshers (file map + dashboard). Used by the scheduled task and refresh_dashboard.bat."""
import runpy, sys, pathlib, traceback, datetime
HERE = pathlib.Path(__file__).resolve().parent
LOG = HERE / ".refresh_log.txt"
msgs = []
for name in ("refresh_file_map.py", "refresh_dashboard.py"):
    try:
        sys.argv = [name]
        runpy.run_path(str(HERE / name), run_name="__main__")
        msgs.append(f"{name}: ok")
    except SystemExit:
        msgs.append(f"{name}: ok")
    except Exception:
        msgs.append(f"{name}: FAILED\n{traceback.format_exc()}")
with open(LOG, "a", encoding="utf-8") as f:
    f.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M} " + " | ".join(msgs) + "\n")
