import os
import time
import pandas as pd
import threading
from pythonosc import udp_client
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


def send_sampler_osc(
    filename,
    host="127.0.0.1",
    port=57120,
    volume_db=0.0,
    rate=1.0,
    pitch=0.0,
    xfade=0.01,
    bpm_source=120.0,
    bpm_target=120.0,
    retrig_num_total=0.0,
    retrig_rate_change_beats=1.0,
    retrig_rate_start=1.0,
    retrig_rate_end=0.0,
    retrig_pitch_change=0.0,
    retrig_volume_change=0.0,
    slice_attack_beats=0.001,
    slice_duration_beats=1.0,
    slice_release_beats=0.001,
    slice_num=0.0,
    slice_count=32.0,
    effect_dry=1.0,
    effect_comb=0.0,
    effect_delay=0.0,
    effect_reverb=0.0,
    effect_reverse=0,
):
    try:
        client = udp_client.SimpleUDPClient(host, port)
        client.send_message(
            "/sampler",
            [
                os.path.abspath(filename),
                volume_db,
                rate,
                pitch,
                xfade,
                bpm_source,
                bpm_target,
                retrig_num_total,
                retrig_rate_change_beats,
                retrig_rate_start,
                retrig_rate_end,
                retrig_pitch_change,
                retrig_volume_change,
                slice_attack_beats,
                slice_duration_beats,
                slice_release_beats,
                slice_num,
                slice_count,
                effect_dry,
                effect_comb,
                effect_delay,
                effect_reverb,
                effect_reverse,
            ],
        )
        return True
    except:
        return False


def execute_sequence(df, master_bpm=120, host="127.0.0.1", port=57120):
    if df is None or df.empty:
        return

    beat_duration = 60.0 / master_bpm
    row_duration = beat_duration / 2
    while True:
        start_time = time.time()

        for index, row in df.iterrows():
            expected_time = start_time + (index * row_duration)
            current_time = time.time()

            if current_time < expected_time:
                time.sleep(expected_time - current_time)

            filename = str(row.get("Filename", "")).strip()
            if not filename or pd.isna(row.get("Filename")):
                continue

            params = {
                "filename": filename,
                "volume_db": float(row.get("Volume (dB)", 0.0)),
                "pitch": float(row.get("Pitch", 0.0)),
                "bpm_source": float(row.get("Source BPM", 120.0)),
                "bpm_target": float(row.get("Target BPM", 120.0)),
                "slice_num": float(row.get("Slice", 0.0)),
                "slice_count": float(row.get("Slice Count", 32.0)),
                "effect_dry": float(row.get("Dry", 1.0)),
                "effect_comb": float(row.get("Comb", 0.0)),
                "effect_delay": float(row.get("Delay", 0.0)),
                "effect_reverb": float(row.get("Reverb", 0.0)),
                "effect_reverse": int(row.get("Reverse", 0)),
                "retrig_num_total": float(row.get("Retrig Num", 0.0)),
                "retrig_rate_start": float(row.get("R-Rate Start", 1.0)),
                "retrig_rate_end": float(row.get("R-Rate End", 0.0)),
                "retrig_volume_change": float(row.get("R-Volume", 0.0)),
                "retrig_pitch_change": float(row.get("R-Pitch", 0.0)),
                "host": host,
                "port": port,
            }

            send_sampler_osc(**params)


class ExcelHandler(FileSystemEventHandler):
    def __init__(self, excel_file, master_bpm, host, port):
        self.excel_file = os.path.abspath(excel_file)
        self.master_bpm = master_bpm
        self.host = host
        self.port = port
        self.last_modified = 0
        self.is_running = False

    def on_modified(self, event):
        if event.is_directory or os.path.abspath(event.src_path) != self.excel_file:
            return

        current_time = time.time()
        if current_time - self.last_modified < 1.0 or self.is_running:
            return

        self.last_modified = current_time
        time.sleep(0.3)

        thread = threading.Thread(target=self.run_sequence)
        thread.daemon = True
        thread.start()

    def run_sequence(self):
        self.is_running = True
        try:
            df = pd.read_excel(self.excel_file)
            execute_sequence(df, self.master_bpm, self.host, self.port)
        except:
            pass
        finally:
            self.is_running = False


def start_monitor(excel_file, master_bpm=120, host="127.0.0.1", port=57120):
    if not os.path.exists(excel_file):
        print(f"File not found: {excel_file}")
        return

    try:
        df = pd.read_excel(excel_file)
        execute_sequence(df, master_bpm, host, port)
    except:
        pass

    handler = ExcelHandler(excel_file, master_bpm, host, port)
    observer = Observer()
    observer.schedule(
        handler, os.path.dirname(os.path.abspath(excel_file)) or ".", recursive=False
    )
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    EXCEL_FILE = "sampler.xlsx"
    MASTER_BPM = 180
    HOST = "127.0.0.1"
    PORT = 57120

    start_monitor(EXCEL_FILE, MASTER_BPM, HOST, PORT)
