#!/usr/bin/env python3
"""
lab_logger.py - Automatic data logger for bench instruments.

Supported instruments
  * OHAUS MB90 moisture analyzer      (serial over USB or RS232)
  * OHAUS pH meter a-AB33 / a-AB41    (serial over USB or RS232)
  * Linshang LS177C coffee colorimeter (watches a folder for Excel/CSV exports
                                        made with Linshang's PC software)

Everything is written as CSV files into OUTPUT_FOLDER. Point OUTPUT_FOLDER at a
OneDrive / Google Drive / Dropbox folder and the data is backed up to the cloud
automatically.

Usage
  python lab_logger.py                  start logging (creates config on first run)
  python lab_logger.py --list-ports     show the serial ports this computer can see
  python lab_logger.py --monitor COM5   show raw text arriving on one port (testing)
  python lab_logger.py --selftest       run the parsers on sample printouts

While running, type these commands in the window and press Enter:
  lot <number>     set the lot/batch number added to every new record
  op <name>        set the operator name added to every new record
  note <text>      set a free-text note added to every new record
  status           show connection status of each instrument
  quit             stop the logger

Requires:  pip install pyserial openpyxl
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import sys
import threading
import time
from pathlib import Path

try:
    import serial
    import serial.tools.list_ports
except ImportError:  # allow --selftest without pyserial
    serial = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "lab_logger_config.json"

DEFAULT_CONFIG = {
    "_help": "Edit this file, save it, then restart lab_logger.py. "
             "Use forward slashes or double backslashes in Windows paths.",
    "station_name": "Bench 1",
    "output_folder": "~/Lab Data Logs",
    "instruments": [
        {
            "name": "MB90",
            "type": "mb90",
            "enabled": True,
            "port": "COM3",
            "baudrate": 9600,
            "bytesize": 8,
            "parity": "N",
            "stopbits": 1,
            "xonxoff": False,
        },
        {
            "name": "pH_meter",
            "type": "ph",
            "enabled": True,
            "port": "COM4",
            "baudrate": 9600,
            "bytesize": 8,
            "parity": "N",
            "stopbits": 1,
            "xonxoff": False,
        },
        {
            "name": "LS177C",
            "type": "colorimeter_folder",
            "enabled": True,
            "watch_folder": "~/Lab Data Logs/Colorimeter Exports",
            "poll_seconds": 10,
        },
    ],
}

# Shared, operator-set context added to every record.
CONTEXT = {"lot": "", "operator": "", "note": ""}
CONTEXT_LOCK = threading.Lock()
STATUS = {}
STOP = threading.Event()


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def now_iso():
    return dt.datetime.now().isoformat(timespec="seconds")


def log(msg):
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def expand(path_str):
    return Path(os.path.expandvars(os.path.expanduser(path_str)))


def context_snapshot():
    with CONTEXT_LOCK:
        return dict(CONTEXT)


class CsvWriter:
    """Appends rows to CSV files. If a file is locked (e.g. open in Excel),
    rows are held in memory and retried, so nothing is lost."""

    def __init__(self):
        self.lock = threading.Lock()
        self.pending = []  # list of (path, rows)

    def append(self, path, rows):
        if not rows:
            return
        with self.lock:
            self.pending.append((Path(path), rows))
            self._flush_locked()

    def flush(self):
        with self.lock:
            self._flush_locked()

    def _flush_locked(self):
        still_pending = []
        for path, rows in self.pending:
            try:
                self._write(path, rows)
            except PermissionError:
                still_pending.append((path, rows))
        if still_pending and len(still_pending) != getattr(self, "_last_warn", 0):
            log(f"WARNING: {still_pending[0][0].name} is locked (open in Excel?). "
                f"Close it; {sum(len(r) for _, r in still_pending)} row(s) waiting.")
        self._last_warn = len(still_pending)
        self.pending = still_pending

    @staticmethod
    def _write(path, rows):
        path.parent.mkdir(parents=True, exist_ok=True)
        new_fields = []
        for r in rows:
            for k in r:
                if k not in new_fields:
                    new_fields.append(k)

        if path.exists() and path.stat().st_size > 0:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                header = next(reader, [])
            missing = [k for k in new_fields if k not in header]
            if missing:
                # Header grew: rewrite the file with the combined header.
                with open(path, newline="", encoding="utf-8-sig") as f:
                    old_rows = list(csv.DictReader(f))
                header = header + missing
                tmp = path.with_suffix(".tmp")
                with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.DictWriter(f, fieldnames=header)
                    w.writeheader()
                    w.writerows(old_rows)
                    w.writerows(rows)
                os.replace(tmp, path)
                return
            with open(path, "a", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=header, extrasaction="ignore")
                w.writerows(rows)
        else:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=new_fields)
                w.writeheader()
                w.writerows(rows)


WRITER = CsvWriter()


def monthly_csv(out_dir, name):
    return out_dir / name / f"{name}_{dt.date.today():%Y-%m}.csv"


def write_raw(out_dir, name, text):
    raw_dir = out_dir / name / "raw"
    raw_dir.mkdir(parents=True, exist_ok=True)
    with open(raw_dir / f"{name}_raw_{dt.date.today():%Y-%m-%d}.txt", "a",
              encoding="utf-8") as f:
        f.write(f"===== received {now_iso()} =====\n{text}\n\n")


def first_float(pattern, text, flags=re.I):
    m = re.search(pattern, text, flags)
    return m.group(1) if m else ""


# --------------------------------------------------------------------------
# Parsers
# --------------------------------------------------------------------------
MB90_LABELS = [
    ("serial_number", r"SNR\s*\(Drying Unit\)\s*(.+)"),
    ("software_version", r"SW\s*\(Drying Unit\)\s*(.+)"),
    ("method_name", r"Method Name\s+(.+)"),
    ("drying_program", r"Drying Prog\w*\s+(.+)"),
    ("drying_temp_C", r"Drying Temp\w*\s+(-?[\d.]+)"),
    ("switch_off", r"Switch Off\s+(.+)"),
    ("start_weight_g", r"Start Weight\s+(-?[\d.]+)"),
    ("total_time", r"Total Time\.?\s+([\d:]+)"),
    ("end_result", r"End Result\s*\.?\s+(-?[\d.]+)"),
    ("end_result_unit", r"End Result\s*\.?\s+-?[\d.]+\s*(%MC|%DC|%RG|g)"),
    ("nominal_weight_g", r"Nominal Weight\s+(-?[\d.]+)"),
    ("actual_weight_g", r"Actual Weight\s+(-?[\d.]+)"),
    ("difference_g", r"Difference\s+(-?[\d.]+)"),
    ("cell_temperature_C", r"Cell Temperature\s+(-?[\d.]+)"),
    ("temp1_target_C", r"Temp1 target\s+(-?[\d.]+)"),
    ("temp1_actual_C", r"Temp1 actual\s+(-?[\d.]+)"),
    ("temp2_target_C", r"Temp2 target\s+(-?[\d.]+)"),
    ("temp2_actual_C", r"Temp2 actual\s+(-?[\d.]+)"),
    ("adjustment", r"Adjustment\s+(Done|Failed|\w+)"),
    ("stat_sample_number", r"Sample Number\s+(\d+)"),
    ("stat_mean", r"Mean Value\s+(-?[\d.]+)"),
    ("stat_std_dev", r"Standard Deviation\s+(-?[\d.]+)"),
    ("stat_min", r"Minimum Value\s+(-?[\d.]+)"),
    ("stat_max", r"Maximum Value\s+(-?[\d.]+)"),
]
MB90_CURVE = re.compile(r"^\s*(\d{1,3}:\d{2})\s*min\s+(-?[\d.:]+)\s*(%MC|%DC|%RG|g)", re.I)


def parse_mb90(text):
    upper = text.upper()
    if "WEIGHT ADJUST" in upper:
        rtype = "weight_adjustment"
    elif "TEMPERATURE ADJUST" in upper:
        rtype = "temperature_adjustment"
    elif "STATISTICS" in upper:
        rtype = "statistics"
    elif "MOISTURE DETERMINATION" in upper:
        rtype = "moisture_test"
    else:
        rtype = "other"

    rec = {"record_type": rtype}
    for key, pat in MB90_LABELS:
        rec[key] = ""
        for line in text.splitlines():
            m = re.search(pat, line, re.I)
            if m:
                rec[key] = m.group(1).strip()
                break

    curve = []
    for line in text.splitlines():
        m = MB90_CURVE.match(line)
        if m:
            # "29:36" style typos in printouts -> treat ':' as '.'
            curve.append((m.group(1), m.group(2).replace(":", "."), m.group(3)))
    rec["curve_points"] = len(curve)
    if rtype == "moisture_test" and not rec["end_result"] and curve:
        rec["end_result"], rec["end_result_unit"] = curve[-1][1], curve[-1][2]
        rec["record_type"] = "moisture_test_incomplete"
    return rec, curve


def parse_ph(text):
    flat = " ".join(text.split())
    rec = {
        "meter_model": first_float(r"\b(A?-?AB\d{2}[A-Z0-9]+)\b", flat),
        "meter_datetime": first_float(r"(\d{2,4}/\d{1,2}/\d{1,4}\s+\d{1,2}:\d{2}(?::\d{2})?)", flat),
        "user_id": first_float(r"User\s*(?:ID)?\s*[:#]?\s*(\d+)", flat),
        "sample_id": first_float(r"Sample\s*(?:ID)?\s*[:#]?\s*(\w+)", flat),
        "mode": first_float(r"\b(ORP|mV|pH|Cond|TDS|SALT|RES)\b(?=\s+-?[\d.]+)", flat, 0),
        "pH": first_float(r"(-?\d+\.\d+)\s*pH\b", flat),
        "mV": first_float(r"(-?\d+\.\d+)\s*m[Vv]\b", flat),
        "ORP_RmV": first_float(r"(-?\d+\.\d+)\s*RmV\b", flat),
        "temperature": first_float(r"(-?\d+\.\d+)\s*°?\s*[℃℉CF]°?(?![A-Za-z/])", flat, 0),
        "temp_unit": first_float(r"-?\d+\.\d+\s*°?\s*([CF℃℉])°?(?![A-Za-z/])", flat, 0)
                     .replace("℃", "C").replace("℉", "F"),
        "temp_type": first_float(r"\b(ATC|MTC)\b", flat),
        "slope_pct": first_float(r"(-?\d+\.\d+)\s*%(?!/)", flat),
        "offset_mV": "",
        "conductivity": first_float(r"(-?\d+\.?\d*)\s*[uµm]S/cm", flat),
        "tds_mg_L": first_float(r"(-?\d+\.?\d*)\s*mg/L", flat),
        "calibration_record": "yes" if re.search(r"Cal\s*Point|Calibration|Offset\s*mV", flat, re.I) else "",
    }
    mvs = re.findall(r"(-?\d+\.\d+)\s*m[Vv]\b", flat)
    if len(mvs) >= 2:        # measurement printout: first mV is reading, last is offset
        rec["offset_mV"] = mvs[-1]
    sn = re.search(r"\b(\d{8,12})\b", flat)
    rec["serial_number"] = sn.group(1) if sn else ""
    return rec


# --------------------------------------------------------------------------
# Serial instrument worker
# --------------------------------------------------------------------------
class SerialInstrument(threading.Thread):
    START_MARKERS = {"ph": re.compile(r"\bA?-?AB(33|41)", re.I)}
    END_MARKERS = {"mb90": re.compile(r"-{3,}\s*END\s*-{3,}", re.I)}
    IDLE_FLUSH = {"mb90": 900.0, "ph": 2.5}   # seconds of silence before a record is closed

    def __init__(self, cfg, out_dir):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.name_ = cfg["name"]
        self.kind = cfg["type"]
        self.out_dir = out_dir
        self.buffer = []
        self.last_rx = 0.0
        STATUS[self.name_] = "starting"

    def open_port(self):
        parity = {"N": serial.PARITY_NONE, "E": serial.PARITY_EVEN,
                  "O": serial.PARITY_ODD}[self.cfg.get("parity", "N").upper()[0]]
        return serial.Serial(
            port=self.cfg["port"],
            baudrate=int(self.cfg.get("baudrate", 9600)),
            bytesize=int(self.cfg.get("bytesize", 8)),
            parity=parity,
            stopbits=float(self.cfg.get("stopbits", 1)),
            xonxoff=bool(self.cfg.get("xonxoff", False)),
            timeout=0.5,
        )

    def run(self):
        while not STOP.is_set():
            try:
                with self.open_port() as port:
                    STATUS[self.name_] = f"connected on {self.cfg['port']}"
                    log(f"{self.name_}: connected on {self.cfg['port']}")
                    partial = b""
                    while not STOP.is_set():
                        chunk = port.read(512)
                        if chunk:
                            self.last_rx = time.time()
                            partial += chunk
                            while b"\n" in partial or b"\r" in partial:
                                line, partial = re.split(rb"\r\n|\n|\r", partial, maxsplit=1)
                                self.handle_line(self.decode(line))
                        elif partial and time.time() - self.last_rx > 1.0:
                            self.handle_line(self.decode(partial))
                            partial = b""
                        self.check_idle()
            except Exception as e:  # port missing, unplugged, busy...
                if STATUS.get(self.name_, "").startswith("connected"):
                    log(f"{self.name_}: connection lost ({e}). Retrying...")
                elif STATUS.get(self.name_) != "waiting":
                    log(f"{self.name_}: cannot open {self.cfg['port']} ({e}). "
                        f"Will keep retrying every 5 s.")
                STATUS[self.name_] = "waiting"
                self.flush()
                STOP.wait(5)

    @staticmethod
    def decode(b):
        try:
            return b.decode("utf-8")
        except UnicodeDecodeError:
            return b.decode("latin-1")

    def handle_line(self, line):
        if not line.strip():
            if self.buffer:
                self.buffer.append("")
            return
        start = self.START_MARKERS.get(self.kind)
        if start and start.search(line) and any(l.strip() for l in self.buffer):
            self.flush()
        self.buffer.append(line.rstrip())
        end = self.END_MARKERS.get(self.kind)
        if end and end.search(line):
            self.flush()

    def check_idle(self):
        if self.buffer and time.time() - self.last_rx > self.IDLE_FLUSH.get(self.kind, 3.0):
            self.flush()

    def flush(self):
        text = "\n".join(self.buffer).strip()
        self.buffer = []
        if text:
            save_serial_record(self.name_, self.kind, text, self.out_dir)


def save_serial_record(name, kind, text, out_dir):
    ctx = context_snapshot()
    logged = now_iso()
    record_id = f"{name}-{dt.datetime.now():%Y%m%d-%H%M%S-%f}"[:-3]
    base = {"logged_at": logged, "record_id": record_id, "instrument": name,
            "lot": ctx["lot"], "operator": ctx["operator"], "note": ctx["note"]}
    write_raw(out_dir, name, text)

    if kind == "mb90":
        rec, curve = parse_mb90(text)
        row = {**base, **rec, "raw_text": text.replace("\n", " | ")}
        WRITER.append(monthly_csv(out_dir, name), [row])
        if curve:
            WRITER.append(out_dir / name / f"{name}_curves_{dt.date.today():%Y-%m}.csv",
                          [{"record_id": record_id, "logged_at": logged, "lot": ctx["lot"],
                            "elapsed": t, "value": v, "unit": u} for t, v, u in curve])
        summary = f"{rec['record_type']} {rec.get('end_result', '')} {rec.get('end_result_unit', '')}"
    else:
        rec = parse_ph(text)
        row = {**base, **rec, "raw_text": text.replace("\n", " | ")}
        WRITER.append(monthly_csv(out_dir, name), [row])
        summary = f"pH {rec['pH'] or '-'}  {rec['temperature']}{rec['temp_unit']}  mV {rec['mV'] or '-'}"
    log(f"{name}: saved {summary.strip()}  (lot '{ctx['lot']}')")


# --------------------------------------------------------------------------
# Colorimeter folder watcher
# --------------------------------------------------------------------------
class FolderWatcher(threading.Thread):
    EXTENSIONS = {".csv", ".xlsx", ".xlsm"}

    def __init__(self, cfg, out_dir):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.name_ = cfg["name"]
        self.out_dir = out_dir
        self.folder = expand(cfg["watch_folder"])
        self.done_dir = self.folder / "imported"
        self.poll = float(cfg.get("poll_seconds", 10))
        STATUS[self.name_] = f"watching {self.folder}"

    def run(self):
        self.folder.mkdir(parents=True, exist_ok=True)
        self.done_dir.mkdir(exist_ok=True)
        log(f"{self.name_}: watching folder {self.folder}")
        while not STOP.is_set():
            for f in sorted(self.folder.iterdir()):
                if f.is_file() and f.suffix.lower() in self.EXTENSIONS and not f.name.startswith("~$"):
                    if time.time() - f.stat().st_mtime < 3:
                        continue  # still being written
                    try:
                        self.import_file(f)
                    except PermissionError:
                        continue  # still open in the Linshang software / Excel
                    except Exception as e:
                        log(f"{self.name_}: could not read {f.name}: {e}")
                        self.move(f, self.folder / "failed")
            STOP.wait(self.poll)

    def read_rows(self, f):
        if f.suffix.lower() == ".csv":
            with open(f, newline="", encoding="utf-8-sig", errors="replace") as fh:
                table = [r for r in csv.reader(fh)]
            return [("csv", table)]
        if openpyxl is None:
            raise RuntimeError("openpyxl not installed (pip install openpyxl)")
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        sheets = []
        for ws in wb.worksheets:
            sheets.append((ws.title, [["" if c is None else str(c) for c in row]
                                      for row in ws.iter_rows(values_only=True)]))
        wb.close()
        return sheets

    def import_file(self, f):
        ctx = context_snapshot()
        rows_out = []
        for sheet, table in self.read_rows(f):
            table = [r for r in table if any(str(c).strip() for c in r)]
            if not table:
                continue
            # header = first row with at least 2 non-empty cells
            h_idx = next((i for i, r in enumerate(table)
                          if sum(1 for c in r if str(c).strip()) >= 2), 0)
            header, seen = [], {}
            for i, c in enumerate(table[h_idx]):
                c = str(c).strip() or f"col{i + 1}"
                seen[c] = seen.get(c, 0) + 1
                header.append(c if seen[c] == 1 else f"{c}_{seen[c]}")
            for n, r in enumerate(table[h_idx + 1:], start=1):
                row = {"logged_at": now_iso(), "instrument": self.name_,
                       "lot": ctx["lot"], "operator": ctx["operator"], "note": ctx["note"],
                       "source_file": f.name, "sheet": sheet, "row": n}
                for k, v in zip(header, r):
                    row[k] = v
                rows_out.append(row)
        WRITER.append(monthly_csv(self.out_dir, self.name_), rows_out)
        dest = self.move(f, self.done_dir)
        log(f"{self.name_}: imported {len(rows_out)} row(s) from {f.name} -> {dest.parent.name}/")

    @staticmethod
    def move(f, target_dir):
        target_dir.mkdir(exist_ok=True)
        dest = target_dir / f"{dt.datetime.now():%Y%m%d-%H%M%S}_{f.name}"
        shutil.move(str(f), dest)
        return dest


# --------------------------------------------------------------------------
# Config, console, entry points
# --------------------------------------------------------------------------
def load_config():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, indent=2), encoding="utf-8")
        print(f"Created {CONFIG_PATH}\nEdit the ports and output_folder, then run again.")
        sys.exit(0)
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def console_loop():
    print("Commands: lot <no.> | op <name> | note <text> | status | quit")
    while not STOP.is_set():
        try:
            line = input().strip()
        except (EOFError, KeyboardInterrupt):
            STOP.set()
            break
        cmd, _, arg = line.partition(" ")
        cmd = cmd.lower()
        if cmd in ("lot", "batch"):
            with CONTEXT_LOCK:
                CONTEXT["lot"] = arg.strip()
            log(f"Lot set to '{arg.strip()}'")
        elif cmd in ("op", "operator"):
            with CONTEXT_LOCK:
                CONTEXT["operator"] = arg.strip()
            log(f"Operator set to '{arg.strip()}'")
        elif cmd == "note":
            with CONTEXT_LOCK:
                CONTEXT["note"] = arg.strip()
            log(f"Note set to '{arg.strip()}'")
        elif cmd == "status":
            for k, v in STATUS.items():
                print(f"  {k:12s} {v}")
            print(f"  lot='{CONTEXT['lot']}' operator='{CONTEXT['operator']}' note='{CONTEXT['note']}'")
        elif cmd in ("quit", "exit", "q"):
            STOP.set()
        elif cmd:
            print("Unknown command. Use: lot, op, note, status, quit")


def run_logger():
    if serial is None:
        sys.exit("pyserial is not installed. Run:  pip install pyserial openpyxl")
    cfg = load_config()
    out_dir = expand(cfg["output_folder"])
    out_dir.mkdir(parents=True, exist_ok=True)
    log(f"Lab logger started at station '{cfg.get('station_name', '')}'")
    log(f"Saving data to {out_dir}")

    for inst in cfg["instruments"]:
        if not inst.get("enabled", True):
            continue
        if inst["type"] in ("mb90", "ph"):
            SerialInstrument(inst, out_dir).start()
        elif inst["type"] == "colorimeter_folder":
            FolderWatcher(inst, out_dir).start()
        else:
            log(f"Unknown instrument type '{inst['type']}' for {inst['name']}")

    threading.Thread(target=console_loop, daemon=True).start()
    try:
        while not STOP.is_set():
            WRITER.flush()
            STOP.wait(5)
    except KeyboardInterrupt:
        STOP.set()
    WRITER.flush()
    log("Logger stopped.")


def list_ports():
    if serial is None:
        sys.exit("pyserial is not installed. Run:  pip install pyserial")
    ports = list(serial.tools.list_ports.comports())
    if not ports:
        print("No serial ports found. Is the instrument plugged in and switched on?")
    for p in ports:
        print(f"{p.device:25s} {p.description}  [{p.hwid}]")


def monitor(port, baud):
    if serial is None:
        sys.exit("pyserial is not installed. Run:  pip install pyserial")
    print(f"Listening on {port} at {baud} baud. Press Print on the instrument. Ctrl+C to stop.")
    with serial.Serial(port, baud, timeout=0.5) as s:
        try:
            while True:
                data = s.read(512)
                if data:
                    sys.stdout.write(SerialInstrument.decode(data))
                    sys.stdout.flush()
        except KeyboardInterrupt:
            print("\nStopped.")


SAMPLE_MB90 = """MOISTURE DETERMINATION
Halogen Moisture Analyzer
Type MB90
SNR(Drying Unit) 1234567
SW(Drying Unit) 1.20
Method Name Method 1
Drying Prog Standard
Drying Temp 105°C
Switch Off A60(1mg/60s)
Start Weight 3.098 g
00:00 min 0.00%MC
00:30 min 9.17 %MC
01:00 min 12.35 %MC
05:21 min 31.94 %MC
Total Time. 05:21 min
End Result . 31.94 %MC
Sample ID:
1.Jan.15 15:35
---------------END----------------"""

SAMPLE_PH = """AB33PH 1234567890 1.00
0 2026/9/29 14:34 001 pH
pH 7.01 pH -2.3 mV
25.0 C ATC
98.5 % 3.2 mV"""


def selftest():
    import tempfile
    out = Path(tempfile.mkdtemp(prefix="lablogger_test_"))
    CONTEXT.update(lot="TEST-LOT", operator="selftest")
    save_serial_record("MB90", "mb90", SAMPLE_MB90, out)
    save_serial_record("pH_meter", "ph", SAMPLE_PH, out)
    colo = out / "colo_in"
    colo.mkdir()
    (colo / "export.csv").write_text("NO,Test Name,L*,a*,b*,SCAA\n1,Bean A,54.46,-25.92,15.35,96\n",
                                     encoding="utf-8")
    time.sleep(3.1)
    FolderWatcher({"name": "LS177C", "watch_folder": str(colo)}, out).import_file(colo / "export.csv")
    WRITER.flush()
    for f in sorted(out.rglob("*.csv")):
        print(f"\n--- {f.relative_to(out)} ---")
        print(f.read_text(encoding="utf-8-sig"))
    print(f"Self-test output folder: {out}")


def main():
    ap = argparse.ArgumentParser(description="Bench instrument data logger")
    ap.add_argument("--list-ports", action="store_true", help="list serial ports")
    ap.add_argument("--monitor", metavar="PORT", help="print raw data from a port")
    ap.add_argument("--baud", type=int, default=9600, help="baud rate for --monitor")
    ap.add_argument("--selftest", action="store_true", help="test parsers with sample data")
    a = ap.parse_args()
    if a.list_ports:
        list_ports()
    elif a.monitor:
        monitor(a.monitor, a.baud)
    elif a.selftest:
        selftest()
    else:
        run_logger()


if __name__ == "__main__":
    main()
