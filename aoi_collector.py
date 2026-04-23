"""
AOI Image Collector & Compressor
Thu thap anh ALL PASS / fail tu folder nguon,
nen JPEG (Pillow) roi dong goi thanh .zip.
Dependencies: PyQt5 (da co san) + Pillow (cai thu cong vao site-packages).
"""
import os, sys, zipfile, io, time
from datetime import datetime
from pathlib import Path

os.environ.setdefault("QT_LOGGING_RULES","qt.qpa.fonts.warning=false")

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QLineEdit, QFileDialog,
    QProgressBar, QTextEdit, QFrame, QSizePolicy, QSpinBox,
    QCheckBox, QComboBox, QMessageBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

try:
    from PIL import Image
    PILLOW_OK = True
except ImportError:
    PILLOW_OK = False

IMAGE_EXTS   = {".jpg",".jpeg",".png",".bmp",".tif",".tiff",".gif",".webp"}
PASS_KEYWORD = "all pass"
FAIL_KEYWORD = "fail"

BG="#0B0E14"; PANEL="#151A25"; DEEP="#0F1520"; TXT="#E2E8F0"
MUTED="#94A3B8"; DIM_C="#475569"; BORDER="#1E293B"
NEON="#00E5FF"; GREEN="#00E676"; RED="#FF3D00"; YELLOW="#F59E0B"
BLUE="#2979FF"

# ═══ WORKER ════════════════════════════════════════════════════════════════
class CollectorWorker(QThread):
    sig_total   = pyqtSignal(int)
    sig_value   = pyqtSignal(int)
    sig_ptext   = pyqtSignal(str)
    sig_log     = pyqtSignal(str, str)
    sig_done    = pyqtSignal(dict)
    sig_err     = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.src_folder = ""; self.dst_folder = ""; self.zip_name = ""
        self.filter_mode = "all"; self.quality = 80; self.max_width = 0
        self.recursive = True; self._cancel = False

    def cancel(self): self._cancel = True

    def run(self):
        self._cancel = False
        try:
            self._collect()
        except Exception as e:
            self.sig_err.emit(str(e))

    def _collect(self):
        src = Path(self.src_folder)
        dst = Path(self.dst_folder)
        if not src.is_dir():
            self.sig_err.emit(f"Source folder not found:\n{src}"); return
        dst.mkdir(parents=True, exist_ok=True)

        self.sig_ptext.emit("Scanning images...")
        self.sig_log.emit("Scanning source folder...", "info")
        all_images = []
        walk = src.rglob("*") if self.recursive else src.iterdir()
        for f in walk:
            if self._cancel: self.sig_err.emit("Cancelled."); return
            if not f.is_file(): continue
            if f.suffix.lower() not in IMAGE_EXTS: continue
            nl = f.name.lower()
            if self.filter_mode == "pass" and PASS_KEYWORD not in nl: continue
            if self.filter_mode == "fail" and FAIL_KEYWORD not in nl: continue
            all_images.append(f)

        total = len(all_images)
        if total == 0:
            self.sig_err.emit("No matching images found."); return
        self.sig_log.emit(f"Found {total} image(s).", "ok")
        self.sig_total.emit(total)

        zip_name = self.zip_name.strip() or f"AOI_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        if not zip_name.endswith(".zip"): zip_name += ".zip"
        zip_path = dst / zip_name
        self.sig_log.emit(f"Output: {zip_path}", "info")

        counts = {"pass":0,"fail":0,"other":0,"errors":0}
        done = 0; t0 = time.time()

        with zipfile.ZipFile(zip_path,"w",compression=zipfile.ZIP_DEFLATED,compresslevel=6) as zf:
            for img_path in all_images:
                if self._cancel:
                    zf.close(); zip_path.unlink(missing_ok=True)
                    self.sig_err.emit("Cancelled."); return
                nl = img_path.name.lower()
                try:
                    data = self._compress(img_path) if PILLOW_OK else img_path.read_bytes()
                    try: rel = img_path.relative_to(src)
                    except ValueError: rel = Path(img_path.name)
                    zf.writestr(str(rel), data)

                    if PASS_KEYWORD in nl:   counts["pass"]+=1;  tag="PASS"
                    elif FAIL_KEYWORD in nl: counts["fail"]+=1;  tag="FAIL"
                    else:                   counts["other"]+=1;  tag="----"

                    done += 1
                    self.sig_value.emit(done)
                    if done % 10 == 0 or done == total:
                        spd = done / max(time.time()-t0, 0.01)
                        self.sig_ptext.emit(f"[{done}/{total}]  {spd:.0f} img/s  |  {rel.name}")
                    self.sig_log.emit(f"[{tag}]  {rel}",
                                      "ok" if tag=="PASS" else "err" if tag=="FAIL" else "info")
                except Exception as e:
                    counts["errors"] += 1
                    self.sig_log.emit(f"Error: {img_path.name}: {e}", "warn")

        elapsed = time.time()-t0
        self.sig_done.emit({
            "total":total,"pass":counts["pass"],"fail":counts["fail"],
            "other":counts["other"],"errors":counts["errors"],
            "zip_path":str(zip_path),
            "zip_size":zip_path.stat().st_size/1024/1024,
            "elapsed":elapsed
        })

    def _compress(self, path):
        with Image.open(path) as img:
            if img.mode in ("RGBA","P","LA"):
                bg = Image.new("RGB", img.size, (255,255,255))
                src_img = img.convert("RGBA") if img.mode=="P" else img
                mask = src_img.split()[-1] if img.mode in ("RGBA","LA") else None
                bg.paste(src_img, mask=mask); img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")
            if self.max_width > 0 and img.width > self.max_width:
                ratio = self.max_width / img.width
                img = img.resize((self.max_width, int(img.height*ratio)), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=self.quality, optimize=True, progressive=True)
            return buf.getvalue()


# ═══ MAIN WINDOW ═══════════════════════════════════════════════════════════
def mk_panel():
    f=QFrame(); f.setStyleSheet(f"QFrame{{background:{PANEL};border-radius:8px;border:0.5px solid {BORDER};}}"); return f
def mk_lbl(t,sz=11,c=TXT,bold=False):
    l=QLabel(t); w="700" if bold else "400"
    l.setStyleSheet(f"color:{c};font-size:{sz}px;font-weight:{w};background:transparent;border:none;"); return l
def mk_btn(t,c,h,sz=12):
    b=QPushButton(t)
    b.setStyleSheet(f"QPushButton{{background:{c};color:white;padding:9px;border-radius:6px;font-weight:bold;font-size:{sz}px;border:none;}}QPushButton:hover{{background:{h};}}QPushButton:disabled{{background:#37474F;color:#78909C;}}")
    b.setCursor(Qt.PointingHandCursor); return b
def stag(t,bg,fg):
    l=QLabel(t); l.setStyleSheet(f"background:{bg};color:{fg};font-size:8px;font-weight:600;padding:2px 9px;border-radius:9px;border:none;letter-spacing:1px;"); l.setFixedHeight(18); return l
def input_sty():
    return f"QLineEdit{{background:{DEEP};color:{TXT};border:1px solid {BORDER};border-radius:5px;padding:5px 8px;font-size:11px;}}QLineEdit:focus{{border-color:{NEON};}}"
def combo_sty():
    return f"QComboBox{{background:{DEEP};color:{TXT};border:1px solid {BORDER};border-radius:5px;padding:5px 8px;font-size:11px;}}QComboBox::drop-down{{border:none;}}QComboBox QAbstractItemView{{background:{PANEL};color:{TXT};}}"
def spin_sty():
    return f"QSpinBox{{background:{DEEP};color:{TXT};border:1px solid {BORDER};border-radius:5px;padding:4px 6px;font-size:11px;}}QSpinBox::up-button,QSpinBox::down-button{{width:18px;}}"

class AOICollector(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AOI Image Collector & Compressor")
        self.setMinimumSize(860,640)
        self.setStyleSheet(f"QMainWindow{{background:{BG};}}QWidget{{background:{BG};color:{TXT};font-family:'Segoe UI',Arial;}}")
        self._output_folder = ""
        self.worker = CollectorWorker()
        self.worker.sig_total.connect(lambda n: (self.prog.setMaximum(n), self.prog.setFormat(f"%v / {n}  (%p%)")))
        self.worker.sig_value.connect(self.prog.setValue)
        self.worker.sig_ptext.connect(self.prog_lbl.setText if hasattr(self,"prog_lbl") else lambda x:None)
        self.worker.sig_log.connect(self._log)
        self.worker.sig_done.connect(self._on_done)
        self.worker.sig_err.connect(self._on_err)
        self._build()
        # Re-wire after build
        self.worker.sig_ptext.connect(self.prog_lbl.setText)

    def _build(self):
        root=QWidget(); self.setCentralWidget(root); root.setStyleSheet(f"background:{BG};")
        rl=QVBoxLayout(root); rl.setContentsMargins(16,12,16,12); rl.setSpacing(10)

        # Title
        top=QHBoxLayout()
        top.addWidget(mk_lbl("AOI IMAGE COLLECTOR & COMPRESSOR",14,NEON,True))
        top.addStretch()
        top.addWidget(mk_lbl("Pillow: ready" if PILLOW_OK else "Pillow not found — zip only",
                              9, GREEN if PILLOW_OK else YELLOW))
        rl.addLayout(top)

        # Folders
        fp=mk_panel(); fl=QGridLayout(fp); fl.setContentsMargins(12,10,12,10); fl.setSpacing(8)
        fl.addWidget(stag("FOLDER SETTINGS","#0B2235","#38BDF8"),0,0,1,3)

        fl.addWidget(mk_lbl("Source Folder",10,MUTED),1,0)
        self.src_edit=QLineEdit(); self.src_edit.setPlaceholderText("Folder containing AOI images...")
        self.src_edit.setStyleSheet(input_sty()); self.src_edit.setReadOnly(True)
        fl.addWidget(self.src_edit,1,1)
        bs=mk_btn("Browse","#0288D1","#0277BD",10); bs.setFixedWidth(80); bs.clicked.connect(self._pick_src)
        fl.addWidget(bs,1,2)

        fl.addWidget(mk_lbl("Destination Folder",10,MUTED),2,0)
        self.dst_edit=QLineEdit(); self.dst_edit.setPlaceholderText("Where to save the .zip file...")
        self.dst_edit.setStyleSheet(input_sty()); self.dst_edit.setReadOnly(True)
        fl.addWidget(self.dst_edit,2,1)
        bd=mk_btn("Browse","#0288D1","#0277BD",10); bd.setFixedWidth(80); bd.clicked.connect(self._pick_dst)
        fl.addWidget(bd,2,2)

        fl.addWidget(mk_lbl("ZIP File Name",10,MUTED),3,0)
        self.zip_edit=QLineEdit(); self.zip_edit.setPlaceholderText("Leave blank → auto: AOI_YYYYMMDD_HHMMSS.zip")
        self.zip_edit.setStyleSheet(input_sty())
        fl.addWidget(self.zip_edit,3,1,1,2)
        fl.setColumnStretch(1,1)
        rl.addWidget(fp)

        # Options
        op=mk_panel(); ol=QGridLayout(op); ol.setContentsMargins(12,10,12,10); ol.setSpacing(8)
        ol.addWidget(stag("OPTIONS","#0A1F10","#4ADE80"),0,0,1,4)

        ol.addWidget(mk_lbl("Collect",10,MUTED),1,0)
        self.filter_combo=QComboBox()
        self.filter_combo.addItems(["All images  (PASS + FAIL + others)",
                                    "PASS only   (filename contains 'ALL PASS')",
                                    "FAIL only   (filename contains 'fail')"])
        self.filter_combo.setStyleSheet(combo_sty()); ol.addWidget(self.filter_combo,1,1)

        self.chk_rec=QCheckBox("Include sub-folders (recursive)")
        self.chk_rec.setChecked(True)
        self.chk_rec.setStyleSheet(f"color:{TXT};font-size:10px;")
        ol.addWidget(self.chk_rec,1,2,1,2)

        if PILLOW_OK:
            ol.addWidget(mk_lbl("JPEG Quality",10,MUTED),2,0)
            self.spin_q=QSpinBox(); self.spin_q.setRange(30,95); self.spin_q.setValue(80)
            self.spin_q.setSuffix("   (30=smallest  95=best quality)")
            self.spin_q.setStyleSheet(spin_sty()); ol.addWidget(self.spin_q,2,1)
            ol.addWidget(mk_lbl("Max Width (px)",10,MUTED),2,2)
            self.spin_w=QSpinBox(); self.spin_w.setRange(0,9999); self.spin_w.setValue(0)
            self.spin_w.setSpecialValueText("0 = keep original size")
            self.spin_w.setStyleSheet(spin_sty()); ol.addWidget(self.spin_w,2,3)
        else:
            n=mk_lbl("Image compression disabled. Install Pillow to enable JPEG quality control and resizing.",9,YELLOW)
            n.setWordWrap(True); ol.addWidget(n,2,0,1,4)
        ol.setColumnStretch(1,1); ol.setColumnStretch(3,1)
        rl.addWidget(op)

        # Buttons
        act=QHBoxLayout(); act.setSpacing(10)
        self.btn_start=mk_btn("Start Collection","#00897B","#00695C",13)
        self.btn_cancel=mk_btn("Cancel","#B71C1C","#8B0000",13); self.btn_cancel.setEnabled(False)
        self.btn_open=mk_btn("Open Output Folder","#37474F","#546E7A",11); self.btn_open.setEnabled(False)
        self.btn_start.clicked.connect(self._start)
        self.btn_cancel.clicked.connect(self._cancel)
        self.btn_open.clicked.connect(self._open_out)
        act.addWidget(self.btn_start); act.addWidget(self.btn_cancel)
        act.addStretch(); act.addWidget(self.btn_open)
        rl.addLayout(act)

        # Progress
        pp=mk_panel(); pl=QVBoxLayout(pp); pl.setContentsMargins(12,8,12,8); pl.setSpacing(6)
        pl.addWidget(stag("PROGRESS","#1C0D2B","#C084FC"))
        self.prog=QProgressBar(); self.prog.setValue(0)
        self.prog.setStyleSheet(f"QProgressBar{{background:{DEEP};border-radius:4px;border:none;height:18px;color:{TXT};font-size:10px;}}QProgressBar::chunk{{background:{NEON};border-radius:4px;}}")
        pl.addWidget(self.prog)
        self.prog_lbl=mk_lbl("Ready — select folders then click Start.",9,DIM_C)
        pl.addWidget(self.prog_lbl)
        rl.addWidget(pp)

        # Log
        lp=mk_panel(); ll=QVBoxLayout(lp); ll.setContentsMargins(10,8,10,8); ll.setSpacing(4)
        lt=QHBoxLayout(); lt.addWidget(stag("LOG","#0B2235","#38BDF8")); lt.addStretch()
        bc=QPushButton("Clear"); bc.setFixedSize(52,18)
        bc.setStyleSheet(f"QPushButton{{background:{DEEP};color:{MUTED};border:1px solid {BORDER};border-radius:4px;font-size:9px;}}QPushButton:hover{{color:{TXT};}}")
        bc.clicked.connect(lambda: self.log_box.clear()); lt.addWidget(bc); ll.addLayout(lt)
        self.log_box=QTextEdit(); self.log_box.setReadOnly(True); self.log_box.setMaximumHeight(200)
        self.log_box.setStyleSheet(f"background:{DEEP};color:{MUTED};border:none;font-family:Consolas,'Courier New',monospace;font-size:10px;")
        ll.addWidget(self.log_box); rl.addWidget(lp)

    def _pick_src(self):
        d=QFileDialog.getExistingDirectory(self,"Select Source Folder (AOI images)")
        if d: self.src_edit.setText(d); self._log(f"Source: {d}","info")
    def _pick_dst(self):
        d=QFileDialog.getExistingDirectory(self,"Select Destination Folder (for ZIP)")
        if d: self.dst_edit.setText(d); self._log(f"Destination: {d}","info")

    def _start(self):
        src=self.src_edit.text().strip(); dst=self.dst_edit.text().strip()
        if not src: QMessageBox.warning(self,"Missing","Please select a source folder."); return
        if not dst: QMessageBox.warning(self,"Missing","Please select a destination folder."); return
        if self.worker.isRunning(): QMessageBox.warning(self,"Busy","Already running."); return

        self.worker.src_folder  = src
        self.worker.dst_folder  = dst
        self.worker.zip_name    = self.zip_edit.text().strip()
        self.worker.filter_mode = ["all","pass","fail"][self.filter_combo.currentIndex()]
        self.worker.recursive   = self.chk_rec.isChecked()
        if PILLOW_OK:
            self.worker.quality   = self.spin_q.value()
            self.worker.max_width = self.spin_w.value()
        self._output_folder = dst
        self.log_box.clear(); self.prog.setValue(0); self.prog.setMaximum(100)
        self._log("Starting...","info")
        self.btn_start.setEnabled(False); self.btn_cancel.setEnabled(True); self.btn_open.setEnabled(False)
        self.worker.start()

    def _cancel(self):
        if self.worker.isRunning():
            self.worker.cancel(); self._log("Cancel requested...","warn"); self.btn_cancel.setEnabled(False)

    def _open_out(self):
        if self._output_folder and os.path.isdir(self._output_folder):
            if sys.platform=="win32": os.startfile(self._output_folder)
            else: os.system(f'xdg-open "{self._output_folder}"')

    def _on_done(self,s):
        self.btn_start.setEnabled(True); self.btn_cancel.setEnabled(False); self.btn_open.setEnabled(True)
        self.prog.setValue(self.prog.maximum())
        for l in [f"{'='*52}",
                  f"  DONE  —  {s['total']} images  in  {s['elapsed']:.1f}s",
                  f"  PASS : {s['pass']}    FAIL : {s['fail']}    Other : {s['other']}    Errors : {s['errors']}",
                  f"  ZIP  : {s['zip_path']}",
                  f"  Size : {s['zip_size']:.2f} MB",
                  f"{'='*52}"]:
            self._log(l,"ok")
        self.prog_lbl.setText(f"Done — {s['total']} images  |  PASS: {s['pass']}  FAIL: {s['fail']}  |  {s['zip_size']:.2f} MB  |  {s['elapsed']:.1f}s")
        QMessageBox.information(self,"Collection Complete",
            f"Collected {s['total']} images\nPASS: {s['pass']}   FAIL: {s['fail']}   Other: {s['other']}\n\nOutput: {s['zip_path']}\nSize: {s['zip_size']:.2f} MB")

    def _on_err(self,msg):
        self.btn_start.setEnabled(True); self.btn_cancel.setEnabled(False)
        self._log(f"ERROR: {msg}","err"); self.prog_lbl.setText(f"Failed: {msg}")
        if "Cancelled" not in msg: QMessageBox.critical(self,"Error",msg)

    def _log(self,msg,level="info"):
        c={"ok":GREEN,"warn":YELLOW,"err":RED,"info":MUTED}.get(level,MUTED)
        ts=datetime.now().strftime("%H:%M:%S")
        self.log_box.append(f'<span style="color:{DIM_C};">[{ts}]</span> <span style="color:{c};">{msg}</span>')
        sb=self.log_box.verticalScrollBar(); sb.setValue(sb.maximum())

if __name__=="__main__":
    app=QApplication(sys.argv); app.setStyle("Fusion")
    w=AOICollector(); w.show()
    sys.exit(app.exec_())
