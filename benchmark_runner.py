import os
import sys
import time
import shutil
import zipfile
import threading
from pathlib import Path
import psutil
import pythoncom
import win32com.client
from wdt.main import main as wdt_main

INPUT_FILE = Path(r"C:\Users\ChaoCheng成超\Desktop\BLUpan-03-001[28OCT26]_Listing.DOC")
OUT_DIR = Path(r"C:\Users\ChaoCheng成超\Desktop\benchmark_output")

class ResourceMonitor(threading.Thread):
    def __init__(self, target_process_names, interval=0.5):
        super().__init__(daemon=True)
        self.target_process_names = [name.lower() for name in target_process_names]
        self.interval = interval
        self.running = True
        self.peak_rss_mb = 0.0
        self.peak_vms_mb = 0.0
        self.peak_cpu_percent = 0.0
        self.samples = []

    def run(self):
        parent_proc = psutil.Process()
        while self.running:
            try:
                # Find current process, its children, and any matching process names
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

                self.samples.append((time.perf_counter(), rss_mb, total_cpu))
            except Exception:
                pass
            time.sleep(self.interval)

    def stop(self):
        self.running = False

def inspect_docx(docx_path: Path):
    if not docx_path.exists():
        return {"exists": False, "error": "File not found"}
    size_mb = docx_path.stat().st_size / (1024 * 1024)
    info = {
        "exists": True,
        "size_mb": round(size_mb, 2),
        "uncompressed_xml_mb": 0.0,
        "xml_chars": 0,
        "table_count": 0,
        "para_count": 0,
        "is_valid_zip": False,
    }
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            info["is_valid_zip"] = True
            if "word/document.xml" in z.namelist():
                xml_info = z.getinfo("word/document.xml")
                info["uncompressed_xml_mb"] = round(xml_info.file_size / (1024 * 1024), 2)
                # Count elements quickly without loading entire DOM into memory if huge
                xml_bytes = z.read("word/document.xml")
                info["xml_chars"] = len(xml_bytes)
                info["table_count"] = xml_bytes.count(b"<w:tbl")
                info["para_count"] = xml_bytes.count(b"<w:p")
    except Exception as e:
        info["error"] = str(e)
    return info

def run_wdt_benchmark():
    wdt_out = OUT_DIR / "wdt"
    wdt_out.mkdir(parents=True, exist_ok=True)
    out_file = wdt_out / (INPUT_FILE.stem + ".docx")
    if out_file.exists():
        out_file.unlink()

    print("[WDT] Starting conversion...")
    monitor = ResourceMonitor(target_process_names=["soffice.bin", "soffice.exe"])
    monitor.start()
    t0 = time.perf_counter()

    exit_code = wdt_main([
        "convert",
        str(INPUT_FILE),
        "-o", str(wdt_out),
        "--to", "docx",
        "-f"
    ])
    t_total = time.perf_counter() - t0
    monitor.stop()
    monitor.join(timeout=2.0)

    print(f"[WDT] Finished in {t_total:.2f}s, exit code: {exit_code}")
    docx_info = inspect_docx(out_file)
    return {
        "tool": "wdt",
        "exit_code": exit_code,
        "duration_sec": round(t_total, 2),
        "peak_rss_mb": round(monitor.peak_rss_mb, 2),
        "peak_vms_mb": round(monitor.peak_vms_mb, 2),
        "peak_cpu_percent": round(monitor.peak_cpu_percent, 1),
        "docx_info": docx_info,
        "out_file": str(out_file),
    }

def run_wincom_benchmark():
    wincom_out = OUT_DIR / "wincom"
    wincom_out.mkdir(parents=True, exist_ok=True)
    out_file = wincom_out / (INPUT_FILE.stem + ".docx")
    if out_file.exists():
        out_file.unlink()

    print("[WINCOM] Starting Word COM conversion...")
    monitor = ResourceMonitor(target_process_names=["winword.exe"])
    monitor.start()

    t0 = time.perf_counter()
    word = None
    doc = None
    t_init = 0.0
    t_open = 0.0
    t_save = 0.0
    t_close = 0.0
    err_msg = ""
    exit_code = 0

    try:
        pythoncom.CoInitialize()
        t_init_start = time.perf_counter()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        word.ScreenUpdating = False
        word.FeatureInstall = 0  # msoFeatureInstallNone
        t_init = time.perf_counter() - t_init_start

        print(f"[WINCOM] Word COM initialized in {t_init:.2f}s. Opening document...")
        t_open_start = time.perf_counter()
        doc = word.Documents.Open(
            FileName=str(INPUT_FILE),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        t_open = time.perf_counter() - t_open_start

        print(f"[WINCOM] Document opened in {t_open:.2f}s. Saving as docx...")
        t_save_start = time.perf_counter()
        # 16 = wdFormatXMLDocument
        doc.SaveAs2(
            FileName=str(out_file),
            FileFormat=16,
            CompatibilityMode=15
        )
        t_save = time.perf_counter() - t_save_start

        print(f"[WINCOM] Document saved in {t_save:.2f}s. Closing...")
        t_close_start = time.perf_counter()
        doc.Close(SaveChanges=0)
        doc = None
        word.Quit()
        word = None
        t_close = time.perf_counter() - t_close_start
    except Exception as e:
        exit_code = 1
        err_msg = str(e)
        print(f"[WINCOM] Error during conversion: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    t_total = time.perf_counter() - t0
    monitor.stop()
    monitor.join(timeout=2.0)

    print(f"[WINCOM] Finished in {t_total:.2f}s, exit code: {exit_code}")
    docx_info = inspect_docx(out_file)
    return {
        "tool": "wincom",
        "exit_code": exit_code,
        "error": err_msg,
        "duration_sec": round(t_total, 2),
        "stage_durations": {
            "init_sec": round(t_init, 2),
            "open_sec": round(t_open, 2),
            "save_sec": round(t_save, 2),
            "close_sec": round(t_close, 2),
        },
        "peak_rss_mb": round(monitor.peak_rss_mb, 2),
        "peak_vms_mb": round(monitor.peak_vms_mb, 2),
        "peak_cpu_percent": round(monitor.peak_cpu_percent, 1),
        "docx_info": docx_info,
        "out_file": str(out_file),
    }

if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    import json
    if mode in ("wdt", "all"):
        res_wdt = run_wdt_benchmark()
        print("RESULT_WDT:", json.dumps(res_wdt, ensure_ascii=False, indent=2))
    if mode in ("wincom", "all"):
        res_wincom = run_wincom_benchmark()
        print("RESULT_WINCOM:", json.dumps(res_wincom, ensure_ascii=False, indent=2))
