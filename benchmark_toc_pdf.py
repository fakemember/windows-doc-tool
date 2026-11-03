import os
import sys
import time
import zipfile
import threading
from pathlib import Path
import psutil
import pythoncom
import win32com.client

INPUT_DOC = Path(r"C:\Users\ChaoCheng成超\Desktop\BLUpan-03-001[28OCT26]_Listing.DOC")
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

def run_wincom_pipeline():
    out_dir = OUT_DIR / "wincom_pipeline"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_docx = out_dir / (INPUT_DOC.stem + "_toc.docx")
    out_pdf = out_dir / (INPUT_DOC.stem + "_toc.pdf")
    for f in [out_docx, out_pdf]:
        if f.exists():
            f.unlink()

    print("[WINCOM Pipeline] Starting Word COM: Open DOC -> Insert TOC -> Save DOCX -> Export PDF...")
    monitor = ResourceMonitor(target_process_names=["winword.exe"])
    monitor.start()

    t0 = time.perf_counter()
    word = None
    doc = None
    t_open = 0.0
    t_toc = 0.0
    t_save_docx = 0.0
    t_export_pdf = 0.0
    t_close = 0.0
    err_msg = ""
    exit_code = 0

    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        word.ScreenUpdating = False
        word.FeatureInstall = 0

        # Step 1: Open DOC
        t1 = time.perf_counter()
        doc = word.Documents.Open(
            FileName=str(INPUT_DOC),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        t_open = time.perf_counter() - t1
        print(f"[WINCOM Pipeline] Document opened in {t_open:.2f}s")

        # Step 2: Insert TOC at beginning
        t2 = time.perf_counter()
        rng = doc.Range(0, 0)
        # Add a page break after TOC position
        rng.InsertParagraphBefore()
        rng.Collapse(0) # wdCollapseStart
        
        # Add TOC: Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel, ...
        toc = doc.TablesOfContents.Add(
            Range=rng,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=3,
            UseFields=True,
            TableID="C",
            RightAlignPageNumbers=True,
            IncludePageNumbers=True,
            UseHyperlinks=True,
            HidePageNumbersInWeb=True,
            UseOutlineLevels=True
        )
        # Insert page break right after TOC
        toc_range_end = doc.Range(toc.Range.End, toc.Range.End)
        toc_range_end.InsertBreak(7) # 7 = wdPageBreak
        
        # Update TOC fields
        doc.Fields.Update()
        t_toc = time.perf_counter() - t2
        print(f"[WINCOM Pipeline] TOC inserted and updated in {t_toc:.2f}s")

        # Step 3: Save as DOCX
        t3 = time.perf_counter()
        doc.SaveAs2(
            FileName=str(out_docx),
            FileFormat=16, # wdFormatXMLDocument
            CompatibilityMode=15
        )
        t_save_docx = time.perf_counter() - t3
        print(f"[WINCOM Pipeline] DOCX saved in {t_save_docx:.2f}s")

        # Step 4: Export to PDF
        t4 = time.perf_counter()
        # 17 = wdExportFormatPDF
        # 1 = wdExportCreateHeadingBookmarks
        doc.ExportAsFixedFormat(
            OutputFileName=str(out_pdf),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0, # wdExportOptimizeForPrint
            CreateBookmarks=1
        )
        t_export_pdf = time.perf_counter() - t4
        print(f"[WINCOM Pipeline] PDF exported in {t_export_pdf:.2f}s")

        t5 = time.perf_counter()
        doc.Close(SaveChanges=0)
        doc = None
        word.Quit()
        word = None
        t_close = time.perf_counter() - t5
    except Exception as e:
        exit_code = 1
        err_msg = str(e)
        print(f"[WINCOM Pipeline] Error: {e}")
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

    pdf_size_mb = round(out_pdf.stat().st_size / (1024 * 1024), 2) if out_pdf.exists() else 0.0
    docx_size_mb = round(out_docx.stat().st_size / (1024 * 1024), 2) if out_docx.exists() else 0.0

    print(f"[WINCOM Pipeline] Finished in {t_total:.2f}s, DOCX: {docx_size_mb} MB, PDF: {pdf_size_mb} MB")
    return {
        "tool": "wincom",
        "exit_code": exit_code,
        "error": err_msg,
        "duration_sec": round(t_total, 2),
        "stage_durations": {
            "open_sec": round(t_open, 2),
            "toc_sec": round(t_toc, 2),
            "save_docx_sec": round(t_save_docx, 2),
            "export_pdf_sec": round(t_export_pdf, 2),
            "close_sec": round(t_close, 2),
        },
        "peak_rss_mb": round(monitor.peak_rss_mb, 2),
        "peak_vms_mb": round(monitor.peak_vms_mb, 2),
        "docx_size_mb": docx_size_mb,
        "pdf_size_mb": pdf_size_mb,
        "docx_path": str(out_docx),
        "pdf_path": str(out_pdf),
    }

if __name__ == "__main__":
    import json
    res = run_wincom_pipeline()
    print("WINCOM_PIPELINE_RESULT:", json.dumps(res, ensure_ascii=False, indent=2))
