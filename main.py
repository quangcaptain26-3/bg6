import sys, os, re, io, math, time, warnings
from collections import defaultdict
os.environ.setdefault("QT_LOGGING_RULES","qt.qpa.fonts.warning=false")
warnings.filterwarnings("ignore")
import logging
logging.getLogger("xlrd").setLevel(logging.ERROR)
logging.getLogger("xlrd.biffh").setLevel(logging.ERROR)

import numpy as np
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QGridLayout, QPushButton, QFileDialog, QLabel, QFrame,
    QComboBox, QMessageBox, QSizePolicy)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QTimer, QPoint
from PyQt5.QtGui  import QFont, QColor, QPainter, QPen, QBrush

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.ticker as mticker
import matplotlib.font_manager as fm
import matplotlib as mpl

def _setup_cjk_font():
    """Tự động tìm font hỗ trợ CJK, bao gồm scan Windows font folder."""
    candidates = [
        "Microsoft YaHei",
        "Microsoft YaHei UI",
        "SimHei",
        "SimSun",
        "NSimSun",
        "FangSong",
        "KaiTi",
        "WenQuanYi Micro Hei",
        "Noto Sans CJK SC",
        "PingFang SC",
        "Hiragino Sans GB",
        "Arial Unicode MS",
    ]
    # Scan registered fonts
    available = {f.name for f in fm.fontManager.ttflist}
    for c in candidates:
        if c in available:
            return c
    # Fallback: scan Windows fonts folder trực tiếp
    import os
    win_font_dir = r"C:\Windows\Fonts"
    cjk_files = ["msyh.ttc","msyhbd.ttc","simhei.ttf","simsun.ttc"]
    if os.path.isdir(win_font_dir):
        for fname in cjk_files:
            fpath = os.path.join(win_font_dir, fname)
            if os.path.exists(fpath):
                try:
                    fm.fontManager.addfont(fpath)
                    # reload
                    prop = fm.FontProperties(fname=fpath)
                    return prop.get_name()
                except Exception:
                    pass
    return "DejaVu Sans"

_CJK_FONT = _setup_cjk_font()

def _apply_font():
    """Apply CJK-compatible font cho matplotlib."""
    mpl.rcParams["font.family"]        = "sans-serif"
    mpl.rcParams["font.sans-serif"]    = [_CJK_FONT, "Segoe UI", "Arial", "DejaVu Sans"]
    mpl.rcParams["axes.unicode_minus"] = False  # fix dấu trừ bị vỡ

_apply_font()

# ── logging ──────────────────────────────────────────────────────────────────
def log(level,tag,msg):
    c={"INFO":"\033[96m","OK":"\033[92m","WARN":"\033[93m","ERROR":"\033[91m"}.get(level.upper(),"")
    print(f"\033[2m[{time.strftime('%H:%M:%S')}]\033[0m \033[1m{c}[{level.upper():5s}]\033[0m {c}[{tag}]\033[0m {msg}",flush=True)

# ── theme ─────────────────────────────────────────────────────────────────────
BG="#0B0E14"; PANEL="#151A25"; DEEP="#0F1520"; TXT="#E2E8F0"; MUTED="#94A3B8"
DIM_C="#475569"; BORDER="#1E293B"; NEON="#00E5FF"; GREEN="#00E676"; RED="#FF3D00"
YELLOW="#F59E0B"; BLUE="#2979FF"; PURPLE="#7C3AED"; OEE_F4="#FF6D00"; OEE_F5="#F9A825"
GREY_BAR="#2D3748"
def qc(h): return QColor(h)

# ═══════════════════════════════════════════════════════════════════════════
# I18N — 3 ngôn ngữ: vi / en / zh
# ═══════════════════════════════════════════════════════════════════════════
I18N = {
    "vi": {
        "app_title":        "⚙  SMART FACTORY DASHBOARD",
        "ready":            "Sẵn sàng",
        "done":             "Hoàn thành ✓",
        "busy":             "Đang xử lý. Vui lòng chờ.",
        "sys_data":         "SYSTEM & DATA",
        "btn_agv":          "📜  Load AGV Logs",
        "btn_oee":          "📊  Load OEE File",
        "btn_aoi":          "🖼   Select AOI Folder",
        "btn_aoi_reset":    "↺  Reset AOI",
        "date_filter":      "📅  Date Filter (All)",
        "agv_task":         "AGV — TASK",
        "kpi_task_today":   "Task hôm nay",
        "kpi_done_start":   "Hoàn thành / Bắt đầu",
        "mes_server":       "MES SERVER",
        "mes_fail_sub":     "lần báo cáo thất bại",
        "mes_ok_sub":       "MES kết nối bình thường",
        "mes_fail_ins":     "MES server không phản hồi",
        "mes_ok_ins":       "✓ Báo cáo gửi thành công",
        "agv_health":       "AGV — SỨC KHOẺ XE",
        "chart_health":     "Offline count / xe",
        "ins_no_data":      "Chưa có dữ liệu",
        "ins_xe_bat":       "⚠ Xe {car}: {n} lần offline — kiểm tra WiFi",
        "ins_xe_warn":      "Xe {car}: {n} lần offline — theo dõi thêm",
        "ins_xe_ok":        "✓ Tất cả xe hoạt động ổn định",
        "oee_hieuqua":      "OEE — HIỆU SUẤT",
        "kpi_f4":           "F4 avg",
        "kpi_f5":           "F5 avg",
        "apq_label":        "A · P · Q",
        "agv_traffic":      "AGV — TRAFFIC THEO GIỜ",
        "chart_traffic":    "Hourly Traffic",
        "agv_top":          "AGV — TOP TRẠM",
        "chart_stations":   "Top Active Stations (AGV)",
        "oee_chart":        "OEE — THEO DÂY CHUYỀN",
        "chart_oee":        "OEE F4 & F5",
        "aoi_quality":      "AOI — QUALITY PER CHUYỀN",
        "aoi_details":      "AOI quality details",
        "aoi_summary":      "Pass rate summary",
        "no_data":          "— no data",
        "all_pass":         "✓ Tất cả chuyền đạt chuẩn ≥90%",
        "worst_fail":       "⚠ {name} fail nặng ({rate:.0f}%) — kiểm tra ngay",
        "worst_warn":       "⚠ {name} chưa đạt ({rate:.0f}%)",
        "oee_worst_low":    "⚠ F5 rất thấp ({f5:.0f}%) — kiểm tra bảo trì",
        "oee_worst_line":   "⚠ {wl} ({wv:.0f}%) kéo hiệu suất xuống",
        "oee_story":        "F4 avg: {f4:.0f}%  ·  F5 avg: {f5:.0f}%  ·  Worst: {wl} ({wv:.0f}%)",
        "peak_sub":         "Peak: {ph} ({pc} task)",
        "confirm_ok":       "Tất cả confirm",
        "confirm_gap":      "{gap} chưa confirm",
        "busiest_ins":      "Bận nhất: {bu}  |  Ca {ph} cao điểm ({pc} task/giờ)",
        "task_unit":        "task",
        "aoi_reset_done":   "AOI đã reset",
        "f4_sub":           "F4 lines avg",
        "f5_sub":           "F5 lines avg",
        "pass_n":           "{gp} pass / {gt} total",
        "loaded_lanes":     "{n} chuyền có data",
    },
    "en": {
        "app_title":        "⚙  SMART FACTORY DASHBOARD",
        "ready":            "Ready",
        "done":             "Done ✓",
        "busy":             "Processing. Please wait.",
        "sys_data":         "SYSTEM & DATA",
        "btn_agv":          "📜  Load AGV Logs",
        "btn_oee":          "📊  Load OEE File",
        "btn_aoi":          "🖼   Select AOI Folder",
        "btn_aoi_reset":    "↺  Reset AOI",
        "date_filter":      "📅  Date Filter (All)",
        "agv_task":         "AGV — TASKS",
        "kpi_task_today":   "Tasks today",
        "kpi_done_start":   "Done / Started",
        "mes_server":       "MES SERVER",
        "mes_fail_sub":     "report failures",
        "mes_ok_sub":       "MES connected normally",
        "mes_fail_ins":     "MES server not responding",
        "mes_ok_ins":       "✓ All reports sent successfully",
        "agv_health":       "AGV — VEHICLE HEALTH",
        "chart_health":     "Offline count / vehicle",
        "ins_no_data":      "No data yet",
        "ins_xe_bat":       "⚠ Vehicle {car}: {n} disconnects — check WiFi/hardware",
        "ins_xe_warn":      "Vehicle {car}: {n} disconnects — monitor closely",
        "ins_xe_ok":        "✓ All vehicles operating normally",
        "oee_hieuqua":      "OEE — PERFORMANCE",
        "kpi_f4":           "F4 avg",
        "kpi_f5":           "F5 avg",
        "apq_label":        "A · P · Q",
        "agv_traffic":      "AGV — HOURLY TRAFFIC",
        "chart_traffic":    "Hourly Traffic",
        "agv_top":          "AGV — TOP STATIONS",
        "chart_stations":   "Top Active Stations (AGV)",
        "oee_chart":        "OEE — BY LINE",
        "chart_oee":        "OEE F4 & F5",
        "aoi_quality":      "AOI — QUALITY PER LINE",
        "aoi_details":      "AOI quality details",
        "aoi_summary":      "Pass rate summary",
        "no_data":          "— no data",
        "all_pass":         "✓ All lines passed ≥90%",
        "worst_fail":       "⚠ {name} high fail rate ({rate:.0f}%) — check AOI now",
        "worst_warn":       "⚠ {name} below standard ({rate:.0f}%)",
        "oee_worst_low":    "⚠ F5 very low ({f5:.0f}%) — check maintenance",
        "oee_worst_line":   "⚠ {wl} ({wv:.0f}%) dragging performance down",
        "oee_story":        "F4 avg: {f4:.0f}%  ·  F5 avg: {f5:.0f}%  ·  Worst: {wl} ({wv:.0f}%)",
        "peak_sub":         "Peak: {ph} ({pc} tasks)",
        "confirm_ok":       "All confirmed",
        "confirm_gap":      "{gap} unconfirmed",
        "busiest_ins":      "Busiest: {bu}  |  {ph} peak hour ({pc} tasks/hr)",
        "task_unit":        "tasks",
        "aoi_reset_done":   "AOI reset",
        "f4_sub":           "F4 lines avg",
        "f5_sub":           "F5 lines avg",
        "pass_n":           "{gp} pass / {gt} total",
        "loaded_lanes":     "{n} lines with data",
    },
    "zh": {
        "app_title":        "⚙  智能工厂仪表板",
        "ready":            "就绪",
        "done":             "完成 ✓",
        "busy":             "处理中，请稍候。",
        "sys_data":         "系统与数据",
        "btn_agv":          "📜  加载 AGV 日志",
        "btn_oee":          "📊  加载 OEE 文件",
        "btn_aoi":          "🖼   选择 AOI 文件夹",
        "btn_aoi_reset":    "↺  重置 AOI",
        "date_filter":      "📅  日期筛选（全部）",
        "agv_task":         "AGV — 任务",
        "kpi_task_today":   "今日任务数",
        "kpi_done_start":   "完成 / 开始",
        "mes_server":       "MES 服务器",
        "mes_fail_sub":     "次报告失败",
        "mes_ok_sub":       "MES 连接正常",
        "mes_fail_ins":     "MES 服务器无响应",
        "mes_ok_ins":       "✓ 所有报告发送成功",
        "agv_health":       "AGV — 小车健康状态",
        "chart_health":     "离线次数 / 小车",
        "ins_no_data":      "暂无数据",
        "ins_xe_bat":       "⚠ 小车 {car}：{n} 次断线 — 检查 WiFi/硬件",
        "ins_xe_warn":      "小车 {car}：{n} 次断线 — 持续观察",
        "ins_xe_ok":        "✓ 所有小车运行正常",
        "oee_hieuqua":      "OEE — 设备效率",
        "kpi_f4":           "F4 平均",
        "kpi_f5":           "F5 平均",
        "apq_label":        "A · P · Q",
        "agv_traffic":      "AGV — 每小时流量",
        "chart_traffic":    "每小时流量",
        "agv_top":          "AGV — 热门站点",
        "chart_stations":   "最活跃站点（AGV）",
        "oee_chart":        "OEE — 按产线",
        "chart_oee":        "OEE F4 & F5",
        "aoi_quality":      "AOI — 各产线质量",
        "aoi_details":      "AOI 质量详情",
        "aoi_summary":      "合格率汇总",
        "no_data":          "— 无数据",
        "all_pass":         "✓ 所有产线合格率 ≥90%",
        "worst_fail":       "⚠ {name} 失败率过高（{rate:.0f}%）— 立即检查",
        "worst_warn":       "⚠ {name} 未达标（{rate:.0f}%）",
        "oee_worst_low":    "⚠ F5 效率极低（{f5:.0f}%）— 检查维护计划",
        "oee_worst_line":   "⚠ {wl}（{wv:.0f}%）拉低整体效率",
        "oee_story":        "F4 均值: {f4:.0f}%  ·  F5 均值: {f5:.0f}%  ·  最差: {wl}（{wv:.0f}%）",
        "peak_sub":         "峰值: {ph}（{pc} 任务）",
        "confirm_ok":       "全部已确认",
        "confirm_gap":      "{gap} 条未确认",
        "busiest_ins":      "最繁忙: {bu}  |  {ph} 高峰（{pc} 任务/小时）",
        "task_unit":        "任务",
        "aoi_reset_done":   "AOI 已重置",
        "f4_sub":           "F4 产线平均",
        "f5_sub":           "F5 产线平均",
        "pass_n":           "{gp} 合格 / {gt} 总计",
        "loaded_lanes":     "{n} 条产线有数据",
    },
}
LANG = "vi"   # default

def t(key, **kwargs):
    """Translate key with optional format args."""
    text = I18N.get(LANG, I18N["vi"]).get(key, key)
    return text.format(**kwargs) if kwargs else text


# ── stations ──────────────────────────────────────────────────────────────────
STATIONS=[
    {"name":"PATH",             "down":"1033,1030,1027","up":"263,1052"},
    {"name":"R650",             "down":"1205",          "up":"1200"},
    {"name":"R770",             "down":"1063",          "up":"1065"},
    {"name":"R370",             "down":"175",           "up":"1061"},
    {"name":"ICX-8100",         "down":"1039",          "up":"1041"},
    {"name":"UX7",              "down":"1037",          "up":"1067"},
    {"name":"4C",               "down":"991",           "up":"989"},
    {"name":"4D-Revlon1-3",     "down":"987",           "up":"798"},
    {"name":"4E-Revlon2-4",     "down":"805",           "up":"802"},
    {"name":"4J",               "down":"972",           "up":"971"},
    {"name":"4K",               "down":"970",           "up":"969"},
    {"name":"4L",               "down":"1263",          "up":"1260"},
    {"name":"4M",               "down":"1257",          "up":"1254"},
    {"name":"4N",               "down":"1251",          "up":"1248"},
    {"name":"4P",               "down":"1245",          "up":"1242"},
    {"name":"4A-LP48-NEW",      "down":"1751",          "up":"1749"},
    {"name":"4B-LP8-NEW",       "down":"1747",          "up":"1744"},
    {"name":"4C-NEW",           "down":"1723",          "up":"1724"},
    {"name":"4D-Revlon1-3-NEW", "down":"1725",          "up":"1726"},
    {"name":"4E-Revlon2-4-NEW", "down":"1727",          "up":"1728"},
]
POINT_MAP={}
for s in STATIONS:
    for p in s["down"].split(","): POINT_MAP[p.strip()]=s["name"]
    for p in s["up"].split(","): POINT_MAP[p.strip()]=s["name"]

AOI_LANES=[
    {"display":"4A",        "aliases":["4A"]},
    {"display":"4B",        "aliases":["4B"]},
    {"display":"4C",        "aliases":["4C"]},
    {"display":"Revlon 1&3","aliases":["Revlon1-3","Revlon13","Revlon1&3","revlon1-3","revlon13"]},
    {"display":"Revlon 2&4","aliases":["Revlon2-4","Revlon24","Revlon2&4","revlon2-4","revlon24"]},
    {"display":"ICX",       "aliases":["ICX","icx"]},
    {"display":"CCX",       "aliases":["CCX","ccx"]},
    {"display":"Giftbox",   "aliases":["Giftbox","giftbox","GiftBox"]},
]
IMAGE_EXTS={".jpg",".jpeg",".png",".bmp",".tif",".tiff"}
TIME_RE=re.compile(r"\d{4}-\d{2}-\d{2}\s+(\d{2}):\d{2}:\d{2}")
POINT_RE=re.compile(r'"point":\s*"?(\d+)"?,\s*"action":\s*"?([a-zA-Z]+)"?')
OFFLINE_RE=re.compile(r"(\d+)号AGV已经掉线")
EXEC_S="AGV执行开始上报"; EXEC_E="AGV执行结束上报"; MES_FAIL="taskStart失败"


# ═══════════════════════════════════════════════════════════════════════════
# FLOATING TOOLTIP
# ═══════════════════════════════════════════════════════════════════════════
class FloatingTooltip(QWidget):
    """Custom painted tooltip — dark card với border + icon."""
    PAD_X = 14; PAD_Y = 10; RADIUS = 8

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self._lines   = []     # list of (icon, text, is_value)
        self._accent  = NEON
        self._t = QTimer(self)
        self._t.setSingleShot(True)
        self._t.timeout.connect(self.hide)

    def show_at(self, gpos, text, accent=None):
        """text format: "Label\nValue" hoặc "Label\nValue1\nValue2"."""
        self._accent = accent or NEON
        raw = text.strip().split("\n")
        self._lines = []
        for i, line in enumerate(raw):
            self._lines.append((line.strip(), i == 0))  # (text, is_header)

        # calc size
        fm_h = QFont("Segoe UI", 11); fm_h.setBold(True)
        fm_b = QFont("Segoe UI", 12); fm_b.setBold(True)
        fm_s = QFont("Segoe UI", 10)
        from PyQt5.QtGui import QFontMetrics
        max_w = 0
        total_h = self.PAD_Y * 2
        for txt, is_hdr in self._lines:
            fm = QFontMetrics(fm_h if is_hdr else fm_b)
            max_w = max(max_w, fm.horizontalAdvance(txt))
            total_h += fm.height() + (4 if is_hdr else 2)
        w = max_w + self.PAD_X * 2 + 8
        h = total_h + 6
        self.resize(w, h)

        x = gpos.x() + 16; y = gpos.y() - h - 8
        sc = QApplication.primaryScreen().geometry()
        if x + w > sc.right():  x = gpos.x() - w - 16
        if y < sc.top():        y = gpos.y() + 20
        self.move(x, y); self.show(); self.raise_()
        self._t.start(4000)

    def hide_now(self): self._t.stop(); self.hide()

    def paintEvent(self, _):
        qp = QPainter(self)
        qp.setRenderHint(QPainter.Antialiasing)
        W, H = self.width(), self.height()

        # ── shadow (offset rect, semi-transparent) ──
        qp.setPen(Qt.NoPen)
        qp.setBrush(QColor(0, 0, 0, 80))
        qp.drawRoundedRect(4, 4, W-4, H-4, self.RADIUS, self.RADIUS)

        # ── background card ──
        qp.setBrush(QColor("#0F1825"))
        qp.drawRoundedRect(0, 0, W-4, H-4, self.RADIUS, self.RADIUS)

        # ── accent border left ──
        qp.setBrush(qc(self._accent))
        qp.drawRoundedRect(0, 0, 3, H-4, 2, 2)

        # ── top border line ──
        qp.setPen(QPen(qc(self._accent), 1))
        qp.drawLine(3, 0, W-4, 0)

        # ── text ──
        y = self.PAD_Y
        for txt, is_hdr in self._lines:
            if is_hdr:
                # header — muted small label
                qp.setFont(QFont("Segoe UI", 9))
                qp.setPen(qc(MUTED))
                qp.drawText(self.PAD_X, y, W, 18, Qt.AlignLeft, txt)
                y += 18
                # separator line
                qp.setPen(QPen(QColor(BORDER), 1))
                qp.drawLine(self.PAD_X, y+1, W-self.PAD_X-4, y+1)
                y += 6
            else:
                # value — big bright
                qp.setFont(QFont("Segoe UI", 12, QFont.Bold))
                qp.setPen(qc(TXT))
                qp.drawText(self.PAD_X, y, W, 22, Qt.AlignLeft, txt)
                y += 22
        qp.end()


# ═══════════════════════════════════════════════════════════════════════════
# MATPLOTLIB CHART CANVAS  — clean, stable, like the old version
# ═══════════════════════════════════════════════════════════════════════════
class ChartCanvas(FigureCanvas):
    """Matplotlib canvas với dark theme + floating tooltip."""
    def __init__(self, tip: FloatingTooltip, figsize=(5,3)):
        self.fig = Figure(figsize=figsize, dpi=88)
        self.fig.patch.set_facecolor(PANEL)
        super().__init__(self.fig)
        self.tip = tip
        self.ax = self.fig.add_subplot(111)
        self._style_ax(self.ax)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setStyleSheet(f"background:{PANEL};border-radius:8px;")
        # hover data
        self._hover_bars  = []  # list of (bar_artist, label, value)
        self._hover_points= []  # list of (x, y, label)
        self._accent      = NEON   # tooltip accent color
        self._is_pct      = False  # True → hiện % thay vì số nguyên
        self.mpl_connect("motion_notify_event", self._on_hover)
        self.mpl_connect("axes_leave_event",    lambda e: self.tip.hide_now())

    def _style_ax(self, ax):
        ax.set_facecolor(PANEL)
        ax.tick_params(colors=MUTED, labelsize=9, length=3)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        for sp in ["bottom","left"]:
            ax.spines[sp].set_edgecolor(BORDER)
        ax.yaxis.set_minor_locator(mticker.NullLocator())

    def _redraw(self):
        self.fig.tight_layout(pad=1.4)
        self.draw_idle()

    # ── Bar chart (vertical) ──────────────────────────────────────────────
    def plot_bar(self, labels, values, colors, title="", ylabel=""):
        self.ax.clear(); self._style_ax(self.ax)
        self._hover_bars=[]; self._hover_points=[]
        x = np.arange(len(labels))
        bars = self.ax.bar(x, values, color=colors, edgecolor=BORDER,
                           linewidth=0.6, width=0.65, zorder=2)
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(labels, rotation=35, ha="right",
                                color=TXT, fontsize=9)
        self.ax.yaxis.set_tick_params(labelcolor=MUTED)
        self.ax.set_xlim(-0.6, len(labels)-0.4)
        self.ax.set_ylim(0, max(values)*1.18 if values else 1)
        self.ax.grid(axis="y", color=BORDER, linewidth=0.5, alpha=0.5)
        if title: self.ax.set_title(title, color=NEON, fontsize=10,
                                    fontweight="bold", pad=10)
        self._accent=BLUE; self._is_pct=False
        for b,l,v in zip(bars,labels,values):
            self._hover_bars.append((b,l,v))
        self._redraw()

    # ── Horizontal bar chart ──────────────────────────────────────────────
    def plot_hbar(self, labels, values, colors, title=""):
        self.ax.clear(); self._style_ax(self.ax)
        self._hover_bars=[]; self._hover_points=[]
        y = np.arange(len(labels))
        bars = self.ax.barh(y, values, color=colors, edgecolor=BORDER,
                            linewidth=0.6, height=0.55, zorder=2)
        self.ax.set_yticks(y)
        self.ax.set_yticklabels(labels, color=TXT, fontsize=9)
        self.ax.xaxis.set_tick_params(labelcolor=MUTED)
        self.ax.set_xlim(0, max(values)*1.18 if values else 1)
        self.ax.set_ylim(-0.6, len(labels)-0.4)
        self.ax.grid(axis="x", color=BORDER, linewidth=0.5, alpha=0.5)
        # value labels
        for b,v in zip(bars,values):
            self.ax.text(v+max(values)*0.01, b.get_y()+b.get_height()/2,
                         str(int(v)), va="center", color=MUTED, fontsize=8)
        if title: self.ax.set_title(title, color=NEON, fontsize=10,
                                    fontweight="bold", pad=10)
        self._accent=NEON; self._is_pct=False
        for b,l,v in zip(bars,labels,values):
            self._hover_bars.append((b,l,v))
        self._redraw()

    # ── Line chart ────────────────────────────────────────────────────────
    def plot_line(self, labels, values, title="", peak_idx=None):
        self.ax.clear(); self._style_ax(self.ax)
        self._hover_bars=[]; self._hover_points=[]
        # x = giờ thực để spacing đều
        def to_h(l):
            try: return int(l.split(":")[0])
            except: return 0
        xs = [to_h(l) for l in labels]
        ys = list(values)
        self.ax.fill_between(xs, ys, alpha=0.12, color=PURPLE)
        self.ax.plot(xs, ys, color=PURPLE, linewidth=2,
                     marker="o", markersize=5,
                     markerfacecolor=PURPLE, markeredgecolor=PANEL,
                     markeredgewidth=1.2, zorder=3)
        if peak_idx is not None and 0<=peak_idx<len(xs):
            self.ax.plot(xs[peak_idx], ys[peak_idx], "o",
                         color=RED, markersize=9, zorder=5,
                         markeredgecolor=TXT, markeredgewidth=1.2)
        self.ax.set_xlim(xs[0]-0.3, xs[-1]+0.3)
        self.ax.set_ylim(0, max(ys)*1.25 if ys else 1)
        self.ax.set_xticks(xs)
        self.ax.set_xticklabels(labels, rotation=35, ha="right",
                                color=TXT, fontsize=9)
        self.ax.yaxis.set_tick_params(labelcolor=MUTED)
        self.ax.grid(axis="y", color=BORDER, linewidth=0.5, alpha=0.5)
        if title: self.ax.set_title(title, color=NEON, fontsize=10,
                                    fontweight="bold", pad=10)
        self._accent=PURPLE; self._is_pct=False
        for x,y,l in zip(xs,ys,labels):
            self._hover_points.append((x,y,l))
        self._redraw()

    # ── Grouped bar (OEE F4/F5) ───────────────────────────────────────────
    def plot_grouped(self, lines, f4v, f5v, title=""):
        self.ax.clear(); self._style_ax(self.ax)
        self._hover_bars=[]; self._hover_points=[]
        x = np.arange(len(lines)); w=0.35
        b4 = self.ax.bar(x-w/2, f4v, w, color=OEE_F4, edgecolor=BORDER,
                         linewidth=0.5, label="F4", zorder=2)
        b5 = self.ax.bar(x+w/2, f5v, w, color=OEE_F5, edgecolor=BORDER,
                         linewidth=0.5, label="F5", zorder=2)
        self.ax.set_xticks(x)
        self.ax.set_xticklabels(lines, rotation=30, ha="right",
                                color=TXT, fontsize=9)
        self.ax.yaxis.set_tick_params(labelcolor=MUTED)
        self.ax.set_ylim(0, 105)
        self.ax.set_xlim(-0.6, len(lines)-0.4)
        self.ax.grid(axis="y", color=BORDER, linewidth=0.5, alpha=0.5)
        leg = self.ax.legend(fontsize=8, facecolor=PANEL,
                             edgecolor=BORDER, labelcolor=MUTED)
        if title: self.ax.set_title(title, color=NEON, fontsize=10,
                                    fontweight="bold", pad=10)
        self._accent=OEE_F4; self._is_pct=True
        for b,l,v in zip(list(b4)+list(b5), lines*2, f4v+f5v):
            self._hover_bars.append((b,l,v))
        self._redraw()

    # ── Hover handler ─────────────────────────────────────────────────────
    def _on_hover(self, event):
        if event.inaxes != self.ax:
            self.tip.hide_now(); return
        widget_y = int(self.height() - event.y)
        gpos = self.mapToGlobal(QPoint(int(event.x), widget_y))

        # check bars
        for b,l,v in self._hover_bars:
            if b.contains(event)[0]:
                # format value nicely
                v_str = f"{v:.1f}%" if self._is_pct else f"{int(v)}"
                self.tip.show_at(gpos, f"{l}\n{v_str}", self._accent)
                return
        # check line points
        if self._hover_points:
            if event.xdata is None: self.tip.hide_now(); return
            dists=[(abs(event.xdata-x), y, l) for x,y,l in self._hover_points]
            dists.sort(key=lambda t:t[0])
            if dists[0][0] < 0.6:
                _,y,l = dists[0]
                self.tip.show_at(gpos, f"{l}\n{int(y)} task", self._accent)
                return
        self.tip.hide_now()


# ═══════════════════════════════════════════════════════════════════════════
# DATA WORKER
# ═══════════════════════════════════════════════════════════════════════════
class DataWorker(QThread):
    oee_ready=pyqtSignal(object); agv_ready=pyqtSignal(dict)
    aoi_ready=pyqtSignal(list);   progress=pyqtSignal(str)
    task_error=pyqtSignal(str,str)
    def __init__(self): super().__init__(); self.task=""; self.file_paths=[]
    def run(self):
        if   self.task=="OEE":  self._oee()
        elif self.task=="LOGS": self._agv()
        elif self.task=="AOI":  self._aoi()

    def _oee(self):
        t0=time.time(); dfs=[]
        for path in self.file_paths:
            fname=os.path.basename(path); ext=os.path.splitext(path)[1].lower()
            self.progress.emit(f"Đọc {fname}…")
            try:
                if ext==".csv": df=pd.read_csv(path)
                elif ext==".xls":
                    try:
                        # suppress xlrd "file size not multiple of sector size" warning
                        import logging
                        logging.getLogger("xlrd").setLevel(logging.ERROR)
                        old_err=sys.stderr; sys.stderr=io.StringIO()
                        old_out=sys.stdout; sys.stdout=io.StringIO()
                        try:
                            df=pd.read_excel(path,engine="xlrd")
                        finally:
                            sys.stderr=old_err; sys.stdout=old_out
                    except Exception:
                        try:
                            tbs=pd.read_html(path); df=tbs[0] if tbs else pd.DataFrame()
                        except Exception:
                            df=pd.DataFrame()
                else: df=pd.read_excel(path,engine="openpyxl")
                if not df.empty: dfs.append(df)
            except Exception as e: log("ERROR","OEE",f"{fname}: {e}")
        if not dfs: self.task_error.emit("OEE","Không đọc được dữ liệu."); return
        try:
            df=pd.concat(dfs,ignore_index=True)
            df.columns=[str(c).strip() for c in df.columns]
            for col in ["樓層","綫","日","A","P","Q"]:
                if col in df.columns: df[col]=df[col].astype(str).str.strip()
            if "OEE" in df.columns:
                df["OEE_Num"]=(df["OEE"].astype(str).str.replace(r"[%\s]","",regex=True)
                               .pipe(pd.to_numeric,errors="coerce"))
            for col in ["A","P","Q"]:
                if col in df.columns:
                    df[f"{col}_Num"]=(df[col].astype(str).str.replace(r"[%\s]","",regex=True)
                                     .pipe(pd.to_numeric,errors="coerce"))
            log("OK","OEE",f"{len(df)} rows {time.time()-t0:.2f}s")
            self.oee_ready.emit(df)
        except Exception as e: self.task_error.emit("OEE",str(e))

    def _agv(self):
        t0=time.time()
        sc=defaultdict(int); ac=defaultdict(int); tc=defaultdict(int); oc=defaultdict(int)
        total=es=ee=mf=0
        for path in self.file_paths:
            self.progress.emit(f"Quét {os.path.basename(path)}…")
            try:
                with open(path,"r",encoding="utf-8",errors="ignore") as f:
                    for line in f:
                        m=OFFLINE_RE.search(line)
                        if m: oc[m.group(1)]+=1
                        if MES_FAIL in line: mf+=1
                        if EXEC_S in line: es+=1
                        if EXEC_E in line: ee+=1
                        if '"points":' not in line: continue
                        tm=TIME_RE.search(line)
                        if not tm: continue
                        hk=f"{tm.group(1)}:00"
                        for pid,act in POINT_RE.findall(line):
                            name=POINT_MAP.get(pid,f"Unknown({pid})")
                            total+=1; sc[name]+=1; ac[act.upper()]+=1; tc[hk]+=1
            except Exception as e: log("ERROR","AGV",str(e))
        filt={k:v for k,v in sc.items() if k!="PATH"}
        bu=max(filt,key=filt.get) if filt else "—"
        ph=max(tc,key=tc.get) if tc else "—"; pc=tc.get(ph,0)
        log("OK","AGV",f"total={total} mes_fail={mf} {time.time()-t0:.2f}s")
        self.agv_ready.emit({"stations":dict(sc),"actions":dict(ac),
            "timeline":dict(sorted(tc.items())),"total":total,
            "exec_start":es,"exec_end":ee,"mes_fail":mf,"offline":dict(oc),
            "peak_hour":ph,"peak_count":pc,"busiest":bu,"p_time":time.time()-t0})

    def _aoi(self):
        results=[{"name":L["display"],"pass":0,"fail":0,"total":0,"rate":0.0,"matched":False}
                 for L in AOI_LANES]
        def _scan(folder_path):
            p=f=0
            for dp,_,files in os.walk(folder_path):
                if dp.replace(folder_path,"").count(os.sep)>3: continue
                for file in files:
                    if os.path.splitext(file)[1].lower() not in IMAGE_EXTS: continue
                    nl=file.lower()
                    if "all pass" in nl: p+=1
                    elif "fail" in nl: f+=1
            return p,f
        def _find(name):
            return next((i for i,L in enumerate(AOI_LANES) if name in L["aliases"]),None)
        for rf in self.file_paths:
            fname=os.path.basename(rf)
            self.progress.emit(f"Quét {fname}…")
            try:
                li=_find(fname)
                if li is not None:
                    p,f=_scan(rf); results[li]["matched"]=True
                    results[li].update({"pass":p,"fail":f,"total":p+f,
                                        "rate":(p/(p+f)*100 if (p+f)>0 else 0.0)})
                    log("OK","AOI",f"[direct] {fname}: p={p} f={f}"); continue
                found=False
                for entry in os.scandir(rf):
                    if not entry.is_dir(): continue
                    li=_find(entry.name)
                    if li is None: log("WARN","AOI",f"'{entry.name}' không khớp"); continue
                    found=True; p,f=_scan(entry.path); results[li]["matched"]=True
                    results[li].update({"pass":p,"fail":f,"total":p+f,
                                        "rate":(p/(p+f)*100 if (p+f)>0 else 0.0)})
                    log("OK","AOI",f"{entry.name}: p={p} f={f}")
                if not found: log("WARN","AOI",f"Không thấy lane nào trong '{fname}'")
            except Exception as e: log("ERROR","AOI",str(e))
        self.aoi_ready.emit(results)


# ═══════════════════════════════════════════════════════════════════════════
# WIDGET HELPERS
# ═══════════════════════════════════════════════════════════════════════════
def mk_panel():
    f=QFrame(); f.setStyleSheet(f"QFrame{{background:{PANEL};border-radius:8px;border:0.5px solid {BORDER};}}"); return f

def mk_lbl(text,size=11,color=TXT,bold=False):
    l=QLabel(text); w="700" if bold else "400"
    l.setStyleSheet(f"color:{color};font-size:{size}px;font-weight:{w};background:transparent;border:none;"); return l

def mk_btn(text,color,hover):
    b=QPushButton(text)
    b.setStyleSheet(f"QPushButton{{background:{color};color:white;padding:10px;border-radius:6px;"
                    f"font-weight:bold;font-size:12px;border:none;}}"
                    f"QPushButton:hover{{background:{hover};}}"
                    f"QPushButton:disabled{{background:#37474F;color:#78909C;}}")
    b.setCursor(Qt.PointingHandCursor); return b

class SectionTag(QLabel):
    def __init__(self,text,bg,fg,parent=None):
        super().__init__(text,parent)
        self.setStyleSheet(f"background:{bg};color:{fg};font-size:8px;font-weight:600;"
                           f"padding:2px 8px;border-radius:9px;border:none;letter-spacing:1px;")
        self.setFixedHeight(18)

class InsightBox(QFrame):
    def __init__(self,text="",accent=NEON):
        super().__init__(); self._a=accent
        self.setMaximumHeight(36)
        lay=QVBoxLayout(self); lay.setContentsMargins(6,3,6,3)
        self.lbl=QLabel(text); self.lbl.setWordWrap(False)
        self.lbl.setStyleSheet(f"color:{MUTED};font-size:9px;background:transparent;border:none;")
        lay.addWidget(self.lbl); self._sty()
    def _sty(self): self.setStyleSheet(f"background:{DEEP};border-left:2px solid {self._a};border-radius:0 4px 4px 0;")
    def set_text(self,text,accent=None):
        self.lbl.setText(text)
        if accent: self._a=accent; self._sty()

class KpiCard(QFrame):
    def __init__(self,label,vc=NEON):
        super().__init__(); self.setStyleSheet(f"background:{DEEP};border-radius:5px;border:none;")
        self.setMaximumHeight(52)
        lay=QVBoxLayout(self); lay.setContentsMargins(7,3,7,3); lay.setSpacing(0)
        self.lbl=mk_lbl(label,8,DIM_C); self.val=mk_lbl("—",14,vc,True)
        self.sub=mk_lbl("",8,DIM_C); self.sub.setWordWrap(True)
        lay.addWidget(self.lbl); lay.addWidget(self.val); lay.addWidget(self.sub)
    def update(self,val,sub="",vc=None):
        self.val.setText(str(val)); self.sub.setText(sub)
        if vc: self.val.setStyleSheet(f"color:{vc};font-size:14px;font-weight:700;background:transparent;border:none;")

class AoiBar(QWidget):
    NAME_W=82; STAT_W=112; BAR_H=16; ROW_H=28
    def __init__(self,lane_name):
        super().__init__(); self.lane_name=lane_name
        self._pass=self._fail=self._total=0; self._rate=0.0; self._matched=False
        self.setFixedHeight(self.ROW_H)
    def set_data(self,p,f,total,rate,matched):
        self._pass=p; self._fail=f; self._total=total; self._rate=rate; self._matched=matched; self.update()
    def paintEvent(self,_):
        qp=QPainter(self); qp.setRenderHint(QPainter.Antialiasing)
        W=self.width(); H=self.ROW_H; by=(H-self.BAR_H)//2
        qp.setPen(qc(MUTED)); qp.setFont(QFont("Segoe UI",9))
        qp.drawText(0,0,self.NAME_W,H,Qt.AlignRight|Qt.AlignVCenter,self.lane_name)
        bx=self.NAME_W+8; bw=W-bx-self.STAT_W-8
        if bw<4: qp.end(); return
        qp.setPen(Qt.NoPen); qp.setBrush(qc(GREY_BAR)); qp.drawRoundedRect(bx,by,bw,self.BAR_H,3,3)
        if self._matched and self._total>0:
            pw=max(1,int(bw*self._pass/self._total)); fw=bw-pw
            if pw>0:
                qp.setBrush(qc(GREEN)); qp.drawRoundedRect(bx,by,pw,self.BAR_H,3,3)
                if fw>0: qp.drawRect(bx+max(0,pw-3),by,3,self.BAR_H)
            if fw>0:
                qp.setBrush(qc(RED)); qp.drawRoundedRect(bx+pw,by,fw,self.BAR_H,3,3)
                qp.drawRect(bx+pw,by,3,self.BAR_H)
        sx=bx+bw+8
        if not self._matched or self._total==0:
            qp.setPen(qc(DIM_C)); qp.setFont(QFont("Segoe UI",8))
            qp.drawText(sx,0,self.STAT_W,H,Qt.AlignLeft|Qt.AlignVCenter,t("no_data"))
        else:
            c=GREEN if self._rate>=90 else RED
            qp.setPen(qc(c)); qp.setFont(QFont("Segoe UI",8,QFont.Bold))
            qp.drawText(sx,0,self.STAT_W,H,Qt.AlignLeft|Qt.AlignVCenter,
                        f"{self._rate:.0f}%  P:{self._pass} F:{self._fail}")
        qp.end()


# ═══════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════════════
class SmartFactoryDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Smart Factory Dashboard")
        self.setStyleSheet(f"QMainWindow{{background:{BG};}} QWidget{{background:{BG};color:{TXT};font-family:'Segoe UI',Arial;}}")
        self.oee_df=None; self._aoi_bars=[]
        self._aoi_accum=[{"name":L["display"],"pass":0,"fail":0,"total":0,
                          "rate":0.0,"matched":False} for L in AOI_LANES]
        self.tip=FloatingTooltip(self)
        self._section_tags=[]   # list of (widget, i18n_key)
        self.worker=DataWorker()
        self.worker.oee_ready.connect(self._on_oee)
        self.worker.agv_ready.connect(self._on_agv)
        self.worker.aoi_ready.connect(self._on_aoi)
        self.worker.progress.connect(self._on_progress)
        self.worker.task_error.connect(self._on_error)
        self._build()

    def _build(self):
        root=QWidget(); self.setCentralWidget(root)
        root.setStyleSheet(f"background:{BG};")
        rl=QVBoxLayout(root); rl.setContentsMargins(10,8,10,8); rl.setSpacing(8)
        top=QHBoxLayout()
        self.title_lbl=mk_lbl(t("app_title"),13,NEON,True)
        top.addWidget(self.title_lbl)
        top.addStretch()
        # language switcher
        for lang,label in [("vi","🇻🇳 VI"),("en","🇬🇧 EN"),("zh","🇨🇳 中文")]:
            btn=QPushButton(label)
            btn.setFixedSize(62,24)
            btn.setStyleSheet(f"QPushButton{{background:{DEEP};color:{MUTED};border:1px solid {BORDER};"
                              f"border-radius:4px;font-size:10px;font-weight:600;}}"
                              f"QPushButton:hover{{background:#1E293B;color:{TXT};}}"
                              f"QPushButton:checked{{background:{NEON}22;color:{NEON};border-color:{NEON};}}")
            btn.setCheckable(True)
            btn.setChecked(lang=="vi")
            btn.clicked.connect(lambda _,lg=lang: self._set_lang(lg))
            btn.setObjectName(f"lang_{lang}")
            top.addWidget(btn)
        top.addSpacing(12)
        self.status_lbl=mk_lbl(t("ready"),10,DIM_C)
        top.addWidget(self.status_lbl); rl.addLayout(top)
        g=QGridLayout(); g.setSpacing(8); rl.addLayout(g,stretch=1)
        def mk_tag(text_key,bg,fg):
            tag=SectionTag(t(text_key),bg,fg)
            self._section_tags.append((tag,text_key))
            return tag
        g.setColumnStretch(0,2); g.setColumnStretch(1,3); g.setColumnStretch(2,4); g.setColumnStretch(3,3)
        g.setRowStretch(0,2); g.setRowStretch(1,5); g.setRowStretch(2,4)

        # ── Controls ──────────────────────────────────────────────────────
        cp=mk_panel(); cl=QVBoxLayout(cp); cl.setContentsMargins(10,10,10,10); cl.setSpacing(6)
        cl.addWidget(mk_tag("sys_data","#0B2235","#38BDF8"))
        self.btn_agv=mk_btn("📜  Load AGV Logs","#0288D1","#0277BD")
        self.btn_oee=mk_btn("📊  Load OEE File","#7B1FA2","#6A1B9A")
        self.btn_aoi=mk_btn("🖼   Select AOI Folder","#F57C00","#EF6C00")
        for btn,task in [(self.btn_agv,"LOGS"),(self.btn_oee,"OEE"),(self.btn_aoi,"AOI")]:
            btn.clicked.connect(lambda _,t=task: self._start(t)); cl.addWidget(btn)
        self.btn_aoi_reset=QPushButton("↺  Reset AOI")
        self.btn_aoi_reset.setStyleSheet(
            f"QPushButton{{background:#1E293B;color:{MUTED};padding:5px;border-radius:5px;"
            f"font-size:10px;border:1px solid {BORDER};}}"
            f"QPushButton:hover{{background:#334155;color:{TXT};}}")
        self.btn_aoi_reset.setCursor(Qt.PointingHandCursor)
        self.btn_aoi_reset.clicked.connect(self._reset_aoi)
        cl.addWidget(self.btn_aoi_reset)
        self.combo=QComboBox(); self.combo.addItem("📅  Date Filter (All)")
        self.combo.setStyleSheet(f"QComboBox{{background:{DEEP};color:{TXT};border:1px solid {BORDER};"
            f"padding:6px;border-radius:5px;font-size:11px;}}QComboBox::drop-down{{border:none;}}"
            f"QComboBox QAbstractItemView{{background:{PANEL};color:{TXT};}}")
        self.combo.currentTextChanged.connect(self._filter_oee)
        cl.addWidget(self.combo)
        self.status_detail=mk_lbl("",9,YELLOW); self.status_detail.setWordWrap(True)
        cl.addWidget(self.status_detail); cl.addStretch()
        g.addWidget(cp,0,0,1,1)

        # ── AGV KPI + MES ─────────────────────────────────────────────────
        ak=mk_panel(); akl=QVBoxLayout(ak); akl.setContentsMargins(8,8,8,8); akl.setSpacing(4)
        akl.addWidget(mk_tag("agv_task","#0B2235","#38BDF8"))
        kr=QHBoxLayout(); kr.setSpacing(4)
        self.kpi_total=KpiCard("Task hôm nay",NEON); self.kpi_exec=KpiCard("Hoàn thành / Bắt đầu",GREEN)
        kr.addWidget(self.kpi_total); kr.addWidget(self.kpi_exec); akl.addLayout(kr)
        self.ins_agv=InsightBox(t("ins_no_data"),NEON); akl.addWidget(self.ins_agv)
        mf=QFrame(); mf.setStyleSheet(f"background:{DEEP};border-radius:6px;border:none;")
        mfl=QVBoxLayout(mf); mfl.setContentsMargins(8,6,8,6); mfl.setSpacing(1)
        mfl.addWidget(mk_lbl("MES SERVER",8,DIM_C))
        mr=QHBoxLayout(); mr.setSpacing(8)
        self.mes_val=mk_lbl("—",30,DIM_C,True)
        self.mes_sub=mk_lbl("—",9,DIM_C); self.mes_sub.setWordWrap(True)
        mr.addWidget(self.mes_val); mr.addWidget(self.mes_sub,stretch=1); mfl.addLayout(mr)
        self.mes_ins=InsightBox("—",RED); mfl.addWidget(self.mes_ins)
        akl.addWidget(mf,stretch=1); g.addWidget(ak,0,1)

        # ── AGV Health ────────────────────────────────────────────────────
        hp=mk_panel(); hl=QVBoxLayout(hp); hl.setContentsMargins(8,8,8,6); hl.setSpacing(3)
        hl.addWidget(mk_tag("agv_health","#0B2235","#38BDF8"))
        self.chart_health=ChartCanvas(self.tip)
        hl.addWidget(self.chart_health,stretch=1)
        self.ins_health=InsightBox(t("ins_no_data"),NEON); hl.addWidget(self.ins_health)
        g.addWidget(hp,0,2)

        # ── OEE KPI ───────────────────────────────────────────────────────
        op=mk_panel(); ol=QVBoxLayout(op); ol.setContentsMargins(8,8,8,8); ol.setSpacing(4)
        ol.addWidget(mk_tag("oee_hieuqua","#1C0D2B","#C084FC"))
        kr2=QHBoxLayout(); kr2.setSpacing(4)
        self.kpi_f4=KpiCard("F4 avg",OEE_F4); self.kpi_f5=KpiCard("F5 avg",OEE_F5)
        kr2.addWidget(self.kpi_f4); kr2.addWidget(self.kpi_f5); ol.addLayout(kr2)
        self.ins_oee=InsightBox(t("ins_no_data"),"#C084FC"); ol.addWidget(self.ins_oee)
        apq_f=QFrame(); apq_f.setStyleSheet(f"background:{DEEP};border-radius:5px;border:none;")
        apql=QVBoxLayout(apq_f); apql.setContentsMargins(6,4,6,4); apql.setSpacing(2)
        self._apq_lbl_w=mk_lbl(t("apq_label"),8,DIM_C); apql.addWidget(self._apq_lbl_w)
        self.chart_apq=ChartCanvas(self.tip)
        self.chart_apq.setMaximumHeight(90)
        apql.addWidget(self.chart_apq)
        ol.addWidget(apq_f); g.addWidget(op,0,3)

        # ── AGV Traffic (spans col 0+1) ───────────────────────────────────
        tp=mk_panel(); tl=QVBoxLayout(tp); tl.setContentsMargins(8,8,8,6); tl.setSpacing(3)
        tl.addWidget(mk_tag("agv_traffic","#0B2235","#38BDF8"))
        self.chart_tm=ChartCanvas(self.tip)
        tl.addWidget(self.chart_tm,stretch=1); g.addWidget(tp,1,0,1,2)

        # ── AGV Top Stations ──────────────────────────────────────────────
        sp=mk_panel(); sl=QVBoxLayout(sp); sl.setContentsMargins(8,8,8,6); sl.setSpacing(3)
        sl.addWidget(mk_tag("agv_top","#0B2235","#38BDF8"))
        self.chart_st=ChartCanvas(self.tip)
        sl.addWidget(self.chart_st,stretch=1); g.addWidget(sp,1,2)

        # ── OEE Chart ─────────────────────────────────────────────────────
        ocp=mk_panel(); ocl=QVBoxLayout(ocp); ocl.setContentsMargins(8,8,8,6); ocl.setSpacing(3)
        ocl.addWidget(mk_tag("oee_chart","#1C0D2B","#C084FC"))
        self.chart_oee=ChartCanvas(self.tip)
        ocl.addWidget(self.chart_oee,stretch=1)
        self.oee_story=mk_lbl("—",9,DIM_C); ocl.addWidget(self.oee_story)
        g.addWidget(ocp,1,3)

        # ── AOI full row ──────────────────────────────────────────────────
        ap=mk_panel(); ao=QHBoxLayout(ap); ao.setContentsMargins(12,10,12,10); ao.setSpacing(12)
        bww=QWidget(); bww.setStyleSheet("background:transparent;")
        bwl=QVBoxLayout(bww); bwl.setContentsMargins(0,0,0,0); bwl.setSpacing(0)
        bwl.addWidget(mk_tag("aoi_quality","#0A1F10","#4ADE80"))
        bwl.addSpacing(6)
        self._aoi_bars=[]
        for L in AOI_LANES:
            bar=AoiBar(L["display"]); self._aoi_bars.append(bar); bwl.addWidget(bar)
        bwl.addStretch()
        bot=QHBoxLayout(); bot.setSpacing(8)
        pf=QFrame(); pf.setStyleSheet(f"background:{DEEP};border-radius:6px;border:none;")
        pfl=QVBoxLayout(pf); pfl.setContentsMargins(6,6,6,6)
        self._aoi_det_lbl=mk_lbl(t("aoi_details"),9,DIM_C); pfl.addWidget(self._aoi_det_lbl)
        self.chart_pi=ChartCanvas(self.tip)
        self.chart_pi.setMaximumHeight(130); pfl.addWidget(self.chart_pi)
        sf=QFrame(); sf.setStyleSheet(f"background:{DEEP};border-radius:6px;border:none;")
        sfl=QVBoxLayout(sf); sfl.setContentsMargins(8,8,8,8); sfl.setSpacing(4)
        self._aoi_sum_lbl=mk_lbl(t("aoi_summary"),9,DIM_C); sfl.addWidget(self._aoi_sum_lbl)
        self.aoi_rate=mk_lbl("—",26,DIM_C,True)
        self.aoi_count=mk_lbl("—",9,DIM_C)
        self.aoi_worst=InsightBox("—",RED)
        sfl.addWidget(self.aoi_rate); sfl.addWidget(self.aoi_count)
        sfl.addWidget(self.aoi_worst); sfl.addStretch()
        bot.addWidget(pf,1); bot.addWidget(sf,1)
        bwl.addLayout(bot); ao.addWidget(bww,1); g.addWidget(ap,2,0,1,4)

    # ── task dispatch ──────────────────────────────────────────────────────
    def _start(self,task):
        if self.worker.isRunning(): QMessageBox.warning(self,"Bận","Đang xử lý."); return
        if task=="AOI":
            path=QFileDialog.getExistingDirectory(self,"Chọn folder AOI")
            paths=[path] if path else []
        elif task=="LOGS":
            paths,_=QFileDialog.getOpenFileNames(self,"AGV Logs","","Log (*.log *.txt)")
        else:
            paths,_=QFileDialog.getOpenFileNames(self,"OEE File","","Excel/CSV (*.xls *.xlsx *.csv)")
        if not paths: return
        for b in [self.btn_agv,self.btn_oee,self.btn_aoi,self.btn_aoi_reset]: b.setEnabled(False)
        self.worker.finished.connect(self._on_done)
        self.worker.task=task; self.worker.file_paths=paths; self.worker.start()

    def _on_progress(self,msg):
        self.status_lbl.setText(msg); self.status_detail.setText(msg)

    def _on_done(self):
        for b in [self.btn_agv,self.btn_oee,self.btn_aoi,self.btn_aoi_reset]: b.setEnabled(True)
        self.status_lbl.setText("Hoàn thành ✓"); self.status_detail.setText("✓ Xong")
        try: self.worker.finished.disconnect(self._on_done)
        except TypeError: pass

    def _on_error(self,task,msg): QMessageBox.critical(self,f"Lỗi {task}",msg); self._on_done()

    def _reset_aoi(self):
        self._aoi_accum=[{"name":L["display"],"pass":0,"fail":0,"total":0,
                          "rate":0.0,"matched":False} for L in AOI_LANES]
        self._render_aoi(self._aoi_accum)
        self.status_lbl.setText(t("aoi_reset_done"))

    def _set_lang(self, lang):
        global LANG; LANG=lang
        _apply_font()   # re-apply font (CJK cần khi switch sang zh)
        root=self.centralWidget()
        for lg in ["vi","en","zh"]:
            btn=root.findChild(QPushButton, f"lang_{lg}")
            if btn: btn.setChecked(lg==lang)
        self._retranslate_ui()

    def _retranslate_ui(self):
        """Re-apply all translated strings to existing widgets."""
        self.title_lbl.setText(t("app_title"))
        self.status_lbl.setText(t("ready"))
        self.btn_agv.setText(t("btn_agv"))
        self.btn_oee.setText(t("btn_oee"))
        self.btn_aoi.setText(t("btn_aoi"))
        self.btn_aoi_reset.setText(t("btn_aoi_reset"))
        # update combo placeholder
        cur=self.combo.currentText()
        self.combo.setItemText(0, t("date_filter"))
        # section tags — find by stored ref
        for tag,key in self._section_tags:
            tag.setText(t(key))
        # kpi cards
        self.kpi_total.lbl.setText(t("kpi_task_today"))
        self.kpi_exec.lbl.setText(t("kpi_done_start"))
        self.kpi_f4.lbl.setText(t("kpi_f4"))
        self.kpi_f5.lbl.setText(t("kpi_f5"))
        self.kpi_f4.sub.setText(t("f4_sub"))
        self.kpi_f5.sub.setText(t("f5_sub"))
        self._apq_lbl_w.setText(t("apq_label"))
        if hasattr(self,"_aoi_det_lbl"): self._aoi_det_lbl.setText(t("aoi_details"))
        if hasattr(self,"_aoi_sum_lbl"): self._aoi_sum_lbl.setText(t("aoi_summary"))
        # re-render charts with new language if data exists
        if self.oee_df is not None:
            self._draw_oee(self.oee_df if "All" in self.combo.currentText()
                           else self.oee_df[self.oee_df["日"]==self.combo.currentText()])

    # ── AGV ────────────────────────────────────────────────────────────────
    def _on_agv(self,d):
        total=d["total"]; es=d["exec_start"]; ee=d["exec_end"]
        mf=d["mes_fail"]; off=d["offline"]
        ph=d["peak_hour"]; pc=d["peak_count"]; bu=d["busiest"]

        self.kpi_total.update(total, t("peak_sub",ph=ph,pc=pc))
        gap=es-ee; rc=GREEN if (es==0 or ee/es>=0.9) else YELLOW
        self.kpi_exec.update(f"{ee}/{es}", t("confirm_gap",gap=gap) if gap>0 else t("confirm_ok"), rc)
        self.ins_agv.set_text(t("busiest_ins",bu=bu,ph=ph,pc=pc), NEON)

        if mf>0:
            self.mes_val.setText(str(mf))
            self.mes_val.setStyleSheet(f"color:{RED};font-size:30px;font-weight:700;background:transparent;border:none;")
            self.mes_sub.setText(t("mes_fail_sub"))
            self.mes_ins.set_text(t("mes_fail_ins"),RED)
        else:
            self.mes_val.setText("0")
            self.mes_val.setStyleSheet(f"color:{GREEN};font-size:30px;font-weight:700;background:transparent;border:none;")
            self.mes_sub.setText(t("mes_ok_sub"))
            self.mes_ins.set_text(t("mes_ok_ins"),GREEN)

        if off:
            cars=sorted(off.keys(),key=lambda k:off[k],reverse=True)
            cnts=[off[c] for c in cars]
            cols=[RED if v>100 else YELLOW if v>20 else GREEN for v in cnts]
            self.chart_health.plot_hbar(cars,cnts,cols,t("chart_health"))
            wc=cars[0]; wv=cnts[0]
            if wv>100: self.ins_health.set_text(t("ins_xe_bat",car=wc,n=wv),RED)
            elif wv>20: self.ins_health.set_text(t("ins_xe_warn",car=wc,n=wv),YELLOW)
            else: self.ins_health.set_text(t("ins_xe_ok"),GREEN)

        tl=d["timeline"]
        if tl:
            lbls=list(tl.keys()); vals=list(tl.values())
            pi=lbls.index(ph) if ph in lbls else None
            self.chart_tm.plot_line(lbls,vals,t("chart_traffic"),pi)

        filt={k:v for k,v in d["stations"].items() if k!="PATH"}
        if filt:
            top=dict(sorted(filt.items(),key=lambda x:x[1],reverse=True)[:10])
            self.chart_st.plot_bar(list(top.keys()),list(top.values()),[BLUE]*len(top),t("chart_stations"))

    # ── OEE ────────────────────────────────────────────────────────────────
    def _on_oee(self,df):
        self.oee_df=df; self.combo.blockSignals(True); self.combo.clear()
        self.combo.addItem("📅  Date Filter (All)")
        if "日" in df.columns:
            for d in sorted(df["日"].dropna().unique()): self.combo.addItem(str(d))
        self.combo.blockSignals(False); self._draw_oee(df)

    def _filter_oee(self,txt):
        if self.oee_df is None: return
        self._draw_oee(self.oee_df if "All" in txt else self.oee_df[self.oee_df["日"]==txt])

    def _draw_oee(self,df):
        if df is None or df.empty or "OEE_Num" not in df.columns: return
        avg=df.groupby(["樓層","綫"])["OEE_Num"].mean().reset_index()
        avg.columns=["floor","line","oee"]; avg=avg.dropna()
        lines=sorted(avg["line"].unique())
        f4d={r["line"]:r["oee"] for _,r in avg[avg["floor"]=="F4"].iterrows()}
        f5d={r["line"]:r["oee"] for _,r in avg[avg["floor"]=="F5"].iterrows()}
        f4v=[f4d.get(l,0) for l in lines]; f5v=[f5d.get(l,0) for l in lines]
        self.chart_oee.plot_grouped(lines,f4v,f5v,t("chart_oee"))
        f4nz=[v for v in f4v if v>0]; f5nz=[v for v in f5v if v>0]
        f4a=sum(f4nz)/len(f4nz) if f4nz else 0; f5a=sum(f5nz)/len(f5nz) if f5nz else 0
        all_p=[(l,v) for l,v in zip(lines,f4v) if v>0]+[(l,v) for l,v in zip(lines,f5v) if v>0]
        wl,wv=min(all_p,key=lambda x:x[1]) if all_p else ("—",100)
        self.kpi_f4.update(f"{f4a:.0f}%",t("f4_sub"))
        self.kpi_f5.update(f"{f5a:.0f}%",t("f5_sub"))
        self.oee_story.setText(t("oee_story",f4=f4a,f5=f5a,wl=wl,wv=wv))
        if f5a<20 and f5a>0: self.ins_oee.set_text(t("oee_worst_low",f5=f5a),RED)
        elif wv<50: self.ins_oee.set_text(t("oee_worst_line",wl=wl,wv=wv),YELLOW)
        else: self.ins_oee.set_text(t("oee_story",f4=f4a,f5=f5a,wl=wl,wv=wv),"#C084FC")
        if "A_Num" in df.columns:
            a=df["A_Num"].mean()
            pv=df["P_Num"].mean() if "P_Num" in df.columns else 0
            qv=df["Q_Num"].mean() if "Q_Num" in df.columns else 0
            def apq_col(v): return GREEN if v>=80 else YELLOW if v>=50 else RED
            self.chart_apq.plot_hbar(["A","P","Q"],[round(a),round(pv),round(qv)],
                                     [apq_col(a),apq_col(pv),apq_col(qv)])

    # ── AOI ────────────────────────────────────────────────────────────────
    def _on_aoi(self,results):
        for new_r in results:
            acc=next((a for a in self._aoi_accum if a["name"]==new_r["name"]),None)
            if acc is None: continue
            if new_r["matched"]:
                acc["matched"]=True
                acc["pass"]+=new_r["pass"]; acc["fail"]+=new_r["fail"]
                acc["total"]+=new_r["total"]
                acc["rate"]=(acc["pass"]/acc["total"]*100 if acc["total"]>0 else 0.0)
        QTimer.singleShot(50,lambda: self._render_aoi(self._aoi_accum))

    def _render_aoi(self,results):
        for bar,r in zip(self._aoi_bars,results):
            bar.set_data(r["pass"],r["fail"],r["total"],r["rate"],r["matched"])
        gp=sum(r["pass"] for r in results); gf=sum(r["fail"] for r in results)
        gt=gp+gf; rate=(gp/gt*100) if gt>0 else 0; rc=GREEN if rate>=90 else RED
        self.aoi_rate.setText(f"{rate:.1f}%")
        self.aoi_rate.setStyleSheet(f"color:{rc};font-size:26px;font-weight:700;background:transparent;border:none;")
        self.aoi_count.setText(t("pass_n",gp=gp,gt=gt))
        matched=[r for r in results if r["matched"] and r["total"]>0]
        if matched:
            worst=min(matched,key=lambda r:r["rate"])
            if worst["rate"]<50: self.aoi_worst.set_text(t("worst_fail",name=worst["name"],rate=worst["rate"]),RED)
            elif worst["rate"]<90: self.aoi_worst.set_text(t("worst_warn",name=worst["name"],rate=worst["rate"]),YELLOW)
            else: self.aoi_worst.set_text(t("all_pass"),GREEN)
        if gt>0:
            n_loaded=sum(1 for a in self._aoi_accum if a["matched"])
            self.chart_pi.plot_bar(
                ["PASS","FAIL"],[gp,gf],[GREEN,RED],
                f"PASS {gp} ({gp/gt*100:.0f}%)  |  FAIL {gf}  [{t('loaded_lanes',n=n_loaded)}]")


if __name__=="__main__":
    app=QApplication(sys.argv); app.setStyle("Fusion")
    w=SmartFactoryDashboard(); w.showMaximized()
    sys.exit(app.exec_())
