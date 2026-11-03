import os
import sys
import time
import threading
from pathlib import Path
import psutil
from wdt.main import main as wdt_main

SRC_DOCX = Path(r"C:\Users\ChaoCheng成超\Desktop\benchmark_output\wdt\test_toc.docx")
OUT_DIR = Path(r"C:\Users\ChaoCheng成超\Desktop\benchmark_output\wdt_pipeline")
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_PDF = OUT_DIR / (SRC_DOCX.stem + ".pdf")
if OUT_PDF.exists():
    OUT_PDF.unlink()

class ResourceMonitor(threading.Thread):
    def __init__(self, target_process_names, interval=0.5):
        super().__init__(daemon=True)
        self.target_process_names = [name.lower() for name in target_process_names]
        self.interval = interval
        self.running = True
        self.peak_rss_mb = 0.0
        self.peak_vms_mb = 0.0
        self.peak_cpu_percent = 0.0

    def run(self):
        parent_proc = psutil.Process()
        while self.running:
            try:
                tracked_pids = set()
                try:
                    tracked_pids.add(parent_proc.pid)
                    for child in parent_proc.children(recursive=True):
                        tracked_pids.add(child.pid)
                except Exception:
                    pass

                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        pname = proc.info['name']
                        if pname and pname.lower() in self.target_process_names:
                            tracked_pids.add(proc.info['pid'])
                    except Exception:
                        pass

                total_rss = 0
                total_vms = 0
                total_cpu = 0.0
                for pid in tracked_pids:
                    try:
                        p = psutil.Process(pid)
                        mem = p.memory_info()
                        total_rss += mem.rss
                        total_vms += mem.vms
                        total_cpu += p.cpu_percent(interval=None)
                    except Exception:
                        pass

                rss_mb = total_rss / (1024 * 1024)
                vms_mb = total_vms / (1024 * 1024)
                if rss_mb > self.peak_rss_mb:
                    self.peak_rss_mb = rss_mb
                if vms_mb > self.peak_vms_mb:
                    self.peak_vms_mb = vms_mb
                if total_cpu > self.peak_cpu_percent:
                    self.peak_cpu_percent = total_cpu
            except Exception:
                pass
            time.sleep(self.interval)

    def stop(self):
        self.running = False

print("[WDT PDF Export] Starting DOCX -> PDF conversion via wdt...")
monitor = ResourceMonitor(target_process_names=["soffice.bin", "soffice.exe"])
monitor.start()

t0 = time.perf_counter()
exit_code = wdt_main([
    "convert",
    str(SRC_DOCX),
    "-o", str(OUT_DIR),
    "--to", "pdf",
    "-f"
])
t_total = time.perf_counter() - t0
monitor.stop()
monitor.join(timeout=2.0)

pdf_size_mb = round(OUT_PDF.stat().st_size / (1024 * 1024), 2) if OUT_PDF.exists() else 0.0
print(f"[WDT PDF Export] Finished in {t_total:.2f}s, exit code: {exit_code}, PDF: {pdf_size_mb} MB")

import json
res = {
    "tool": "wdt",
    "exit_code": exit_code,
    "duration_sec": round(t_total, 2),
    "peak_rss_mb": round(monitor.peak_rss_mb, 2),
    "peak_vms_mb": round(monitor.peak_vms_mb, 2),
    "pdf_size_mb": pdf_size_mb,
    "pdf_path": str(OUT_PDF),
}
print("WDT_PDF_RESULT:", json.dumps(res, ensure_ascii=False, indent=2))
