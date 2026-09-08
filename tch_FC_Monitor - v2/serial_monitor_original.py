# -*- coding: utf-8 -*-
"""
tch_FC_Monitor - 串口监控工具

界面功能：
  1. 串口地址下拉框（自动枚举可用串口，默认 COM9）
  2. 波特率下拉框（常用波特率，默认 9600）
  3. “打开串口”按钮（打开 / 关闭切换）
  4. 发送显示框：显示串口发送的消息，每条前缀 hh:mm:ss ms->
  5. 接收显示框：显示串口接收到的文本，每条前缀 hh:mm:ss ms<-
  6. “循环”按钮 + “周期”下拉框（1-10 秒，默认 1）：
     串口打开时按下“循环”，按周期循环发送 LOG 指令
"""

import os
import re
import sys
import threading
from collections import deque
from datetime import datetime
from queue import Queue, Empty

import serial
import serial.tools.list_ports

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except ImportError:
    print("未找到 tkinter，请安装 Tcl/Tk 支持。")
    sys.exit(1)

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


def _port_sort_key(name):
    """COM 端口自然排序（COM2 排在 COM10 之前）。"""
    return [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", name)]


def _to_float(raw):
    """从带单位的字符串中提取数值，如 '27.09V' -> 27.09，'281n' -> 281.0。"""
    if raw is None:
        return None
    m = re.match(r'\s*(-?\d+\.?\d*)', str(raw))
    if not m:
        return None
    try:
        return float(m.group(1))
    except ValueError:
        return None


class CurveCanvas:
    """基于 tkinter Canvas 的实时曲线组件。

    - 黑色背景
    - 多条曲线按各自量程归一化到 0-100% 后同图显示
    - 横坐标为时间（左旧右新）
    """

    def __init__(self, master, title, series, max_points=240,
                 width=440, height=260):
        """
        series: list of dict {
            "label": str, "field": str, "color": str,
            "vmin": float, "vmax": float, "unit": str
        }
        """
        self.title = title
        self.series = series
        self.max_points = max_points
        self.buffers = [deque(maxlen=max_points) for _ in series]
        self.last_time = None
        self._w = width
        self._h = height

        self.frame = ttk.LabelFrame(master, text=title)
        self.canvas = tk.Canvas(self.frame, width=width, height=height,
                                bg="black", highlightthickness=1,
                                highlightbackground="#555")
        self.canvas.pack(fill="both", expand=True, padx=2, pady=2)
        self.canvas.bind("<Configure>", lambda e: self.redraw())

    def add_data(self, data_dict):
        """data_dict: {field_key: raw_value_string, ...}"""
        for i, s in enumerate(self.series):
            self.buffers[i].append(_to_float(data_dict.get(s["field"])))
        self.last_time = datetime.now()
        self.redraw()

    @staticmethod
    def _norm(val, vmin, vmax):
        if val is None:
            return None
        if vmax == vmin:
            return 50.0
        p = (val - vmin) / (vmax - vmin) * 100.0
        return max(0.0, min(100.0, p))

    def redraw(self):
        c = self.canvas
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()
        if w < 10 or h < 10:
            w, h = self._w, self._h

        n_series = len(self.series)
        axis_w = 46          # 每条曲线纵轴占用宽度
        left_margin = 6
        bottom_margin = 22
        top_margin = 30      # 顶部留空给曲线名称

        left = left_margin + n_series * axis_w
        right = w - 8
        top = top_margin
        bottom = h - bottom_margin
        plot_w = right - left
        plot_h = bottom - top
        if plot_w <= 0 or plot_h <= 0:
            return

        # 内部百分比参考网格
        for pct in (0, 25, 50, 75, 100):
            y = bottom - pct / 100.0 * plot_h
            c.create_line(left, y, right, y, fill="#1a1a1a", dash=(2, 4))
        # 绘图区边框
        c.create_rectangle(left, top, right, bottom, outline="#555")

        # 为每条曲线绘制独立纵轴（名称 + 数值刻度 + 范围）
        for i, s in enumerate(self.series):
            ax_x = left_margin + i * axis_w + axis_w - 8
            vmin, vmax = s["vmin"], s["vmax"]
            color = s["color"]
            # 轴线
            c.create_line(ax_x, top, ax_x, bottom, fill=color)
            # 刻度 + 实际数值标签
            for pct in (0, 25, 50, 75, 100):
                y = bottom - pct / 100.0 * plot_h
                val = vmin + pct / 100.0 * (vmax - vmin)
                c.create_line(ax_x - 3, y, ax_x, y, fill=color)
                val_str = f"{int(val)}" if val == int(val) else f"{val:.1f}"
                c.create_text(ax_x - 5, y, text=val_str, anchor="e",
                              fill=color, font=("Consolas", 8))
            # 曲线名称（轴上方第一行）
            c.create_text(ax_x, top - 16, text=s["label"], anchor="s",
                          fill=color, font=("Microsoft YaHei", 8, "bold"))
            # 数值范围（轴上方第二行）
            c.create_text(ax_x, top - 3, text=f"{vmin}-{vmax}{s['unit']}", anchor="s",
                          fill=color, font=("Consolas", 7))

        n = self.max_points
        # 绘制每条曲线
        for i, s in enumerate(self.series):
            buf = self.buffers[i]
            m = len(buf)
            pts = []
            for idx, val in enumerate(buf):
                x = left + (n - m + idx) / max(1, n - 1) * plot_w
                p = self._norm(val, s["vmin"], s["vmax"])
                pts.append(None if p is None else (x, bottom - p / 100.0 * plot_h))
            # 分段画折线（跳过 None）
            seg = []
            for p in pts:
                if p is None:
                    if len(seg) >= 2:
                        c.create_line(seg, fill=s["color"], width=1)
                    seg = []
                else:
                    seg.append(p)
            if len(seg) >= 2:
                c.create_line(seg, fill=s["color"], width=1)
            # 最新点圆点
            if pts and pts[-1] is not None:
                x, y = pts[-1]
                c.create_oval(x - 2.5, y - 2.5, x + 2.5, y + 2.5,
                              fill=s["color"], outline=s["color"])

        # 图例 + 当前值（右上）
        legend_y = top + 6
        for i, s in enumerate(self.series):
            buf = self.buffers[i]
            cur = buf[-1] if buf else None
            cur_str = f"{cur:.2f}{s['unit']}" if cur is not None else "--"
            c.create_text(right - 4, legend_y, anchor="ne",
                          text=f"{s['label']} {cur_str}",
                          fill=s["color"], font=("Consolas", 9))
            legend_y += 13

        # 横轴时间标注
        time_str = self.last_time.strftime("%H:%M:%S") if self.last_time else "--:--:--"
        c.create_text(right, bottom + 8, anchor="e", text=time_str,
                      fill="#aaa", font=("Consolas", 8))
        c.create_text(left, bottom + 8, anchor="w", text="时间→",
                      fill="#aaa", font=("Consolas", 8))


class DataRecorder:
    """将解析数据记录到 xlsx 文件，整点自动切换新文件。"""

    COLUMNS = ["序号", "时间", "输出功率(W)", "输出电压(V)",
               "输出电流(A)", "环境温度(℃)", "电堆温度(℃)"]
    FIELDS = ["POut", "UOut", "IOut", "TAmb", "TStack"]
    SAVE_EVERY = 20  # 每写入若干行保存一次磁盘

    def __init__(self, data_dir):
        self.data_dir = data_dir
        os.makedirs(self.data_dir, exist_ok=True)
        self.wb = None
        self.ws = None
        self.filepath = None
        self.row_index = 0
        self.file_hour = None
        self._since_save = 0

    def start(self):
        self._new_file()

    def _new_file(self):
        now = datetime.now()
        filename = now.strftime("%Y_%m_%d_%H_%M_%S") + ".xlsx"
        self.filepath = os.path.join(self.data_dir, filename)
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "FC_data"
        self.ws.append(self.COLUMNS)
        self.row_index = 0
        self.file_hour = now.replace(minute=0, second=0, microsecond=0)
        self._since_save = 0
        self._save()

    def add_row(self, parsed):
        """parsed: {field_key: raw_value_string, ...}"""
        now = datetime.now()
        # 整点切换：当前整点与文件整点不同则保存旧文件、新建文件
        now_hour = now.replace(minute=0, second=0, microsecond=0)
        if self.file_hour is not None and now_hour != self.file_hour:
            self._save()
            self._new_file()

        self.row_index += 1
        time_str = now.strftime("%H:%M:%S")
        row = [self.row_index, time_str]
        for field in self.FIELDS:
            val = _to_float(parsed.get(field))
            row.append(val if val is not None else "")
        self.ws.append(row)
        self._since_save += 1
        if self._since_save >= self.SAVE_EVERY:
            self._save()
            self._since_save = 0

    def _save(self):
        if self.wb is not None and self.filepath is not None:
            try:
                self.wb.save(self.filepath)
            except Exception as e:
                print(f"[DataRecorder] 保存失败: {e}")

    def stop(self):
        self._save()
        self.wb = None
        self.ws = None
        self.filepath = None


class SerialMonitorApp:
    BAUD_RATES = ["1200", "2400", "4800", "9600", "19200", "38400",
                  "57600", "115200", "230400", "460800"]
    LOOP_CMD = "LOG"
    RECV_FRAME_START = "LOG"
    RECV_FRAME_END = "PFC"

    # 需要解析的字段定义（显示名 -> 协议 key）
    DISPLAY_FIELDS = [
        ("输出功率", "POut"),
        ("电池电压", "UBat"),
        ("输出电压", "UOut"),
        ("输出电流", "IOut"),
        ("环境温度", "TAmb"),
        ("电堆温度", "TStack"),
        ("电堆运行时间", "StackOpTime"),
        ("系统启动次数", "SystemStarts"),
        ("电堆启动次数", "StackStarts"),
        ("当前时间", "SystemTime"),
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("tch_FC_Monitor 串口监控")
        self.root.geometry("1320x820")
        self.root.minsize(1180, 720)

        # 串口状态
        self.ser = None
        self.serial_open = False
        self.read_thread = None
        self.read_running = False
        self.recv_buffer_str = ""
        self.recv_queue = Queue()

        # 循环发送状态
        self.loop_running = False
        self.loop_after_id = None

        # 数据解析显示变量
        self.display_vars = {}
        # 曲线组件
        self.curve1 = None
        self.curve2 = None
        # 数据记录
        self.data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
        self.recorder = None

        self._build_ui()
        self._refresh_ports()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(80, self._poll_recv_queue)

    # ---------------- 界面 ----------------
    def _build_ui(self):
        # 左右分栏：左控制+收发，右曲线（左:右 = 1:2）
        content = ttk.Frame(self.root)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)
        left_frame = ttk.Frame(content)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        right_frame = ttk.Frame(content)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        # 顶部：串口 / 波特率 / 打开
        top = ttk.Frame(left_frame)
        top.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(top, text="串口:").grid(row=0, column=0, sticky="w", padx=4, pady=6)
        self.port_var = tk.StringVar(value="COM9")
        self.port_combo = ttk.Combobox(top, textvariable=self.port_var, width=12)
        self.port_combo.grid(row=0, column=1, sticky="w", padx=4, pady=6)
        ttk.Button(top, text="刷新", width=6, command=self._refresh_ports).grid(
            row=0, column=2, padx=2, pady=6)

        ttk.Label(top, text="波特率:").grid(row=0, column=3, sticky="w", padx=4, pady=6)
        self.baud_var = tk.StringVar(value="9600")
        self.baud_combo = ttk.Combobox(top, textvariable=self.baud_var,
                                       values=self.BAUD_RATES, width=10, state="readonly")
        self.baud_combo.grid(row=0, column=4, sticky="w", padx=4, pady=6)

        self.open_btn = ttk.Button(top, text="打开串口", width=10, command=self._toggle_port)
        self.open_btn.grid(row=0, column=5, padx=10, pady=6)

        # 循环 / 周期
        loop_frame = ttk.Frame(left_frame)
        loop_frame.pack(fill="x", padx=8, pady=2)
        ttk.Label(loop_frame, text="周期(秒):").pack(side="left", padx=(6, 2))
        self.period_var = tk.StringVar(value="1")
        self.period_combo = ttk.Combobox(
            loop_frame, textvariable=self.period_var,
            values=[str(i) for i in range(1, 11)], width=4, state="readonly")
        self.period_combo.pack(side="left", padx=2)
        self.loop_btn = ttk.Button(loop_frame, text="循环", width=8, command=self._toggle_loop)
        self.loop_btn.pack(side="left", padx=8)

        self.status_var = tk.StringVar(value="串口未打开")
        ttk.Label(loop_frame, textvariable=self.status_var, foreground="#666").pack(
            side="left", padx=10)

        # 数据解析显示区
        data_frame = ttk.LabelFrame(left_frame, text="数据解析")
        data_frame.pack(fill="x", padx=8, pady=4)
        # 使用 5 列网格布局，每行两个字段
        for idx, (label, _) in enumerate(self.DISPLAY_FIELDS):
            row = idx // 2
            col = (idx % 2) * 3
            ttk.Label(data_frame, text=label + ":", font=("Microsoft YaHei", 10)).grid(
                row=row, column=col, padx=(8, 2), pady=4, sticky="e")
            var = tk.StringVar(value="--")
            self.display_vars[label] = var
            entry = ttk.Entry(data_frame, textvariable=var, width=22,
                              font=("Consolas", 10), state="readonly")
            entry.grid(row=row, column=col + 1, padx=(2, 4), pady=4, sticky="w")
        # 让各列宽度自适应
        for c in range(6):
            data_frame.columnconfigure(c, weight=0, minsize=110)

        # 发送区
        send_frame = ttk.LabelFrame(left_frame, text="发送")
        send_frame.pack(fill="both", expand=True, padx=8, pady=4)
        send_inner = ttk.Frame(send_frame)
        send_inner.pack(fill="both", expand=True, padx=4, pady=4)
        self.send_text = tk.Text(send_inner, height=8, wrap="word",
                                 font=("Consolas", 10))
        send_sb = ttk.Scrollbar(send_inner, orient="vertical", command=self.send_text.yview)
        self.send_text.configure(yscrollcommand=send_sb.set)
        send_sb.pack(side="right", fill="y")
        self.send_text.pack(side="left", fill="both", expand=True)
        self._make_log(self.send_text)

        send_input = ttk.Frame(send_frame)
        send_input.pack(fill="x", padx=4, pady=(0, 4))
        self.input_var = tk.StringVar()
        entry = ttk.Entry(send_input, textvariable=self.input_var)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        entry.bind("<Return>", lambda e: self._manual_send())
        ttk.Button(send_input, text="发送", width=8, command=self._manual_send).pack(side="left")

        # 接收区
        recv_frame = ttk.LabelFrame(left_frame, text="接收")
        recv_frame.pack(fill="both", expand=True, padx=8, pady=4)
        recv_inner = ttk.Frame(recv_frame)
        recv_inner.pack(fill="both", expand=True, padx=4, pady=4)
        self.recv_text = tk.Text(recv_inner, height=10, wrap="word",
                                 font=("Consolas", 10))
        recv_sb = ttk.Scrollbar(recv_inner, orient="vertical", command=self.recv_text.yview)
        self.recv_text.configure(yscrollcommand=recv_sb.set)
        recv_sb.pack(side="right", fill="y")
        self.recv_text.pack(side="left", fill="both", expand=True)
        self._make_log(self.recv_text)

        # 右侧：两个曲线框
        self.curve1 = CurveCanvas(
            right_frame, "功率/电压/电流",
            [
                {"label": "输出功率", "field": "POut",  "color": "#00FF00", "vmin": 0, "vmax": 200, "unit": "W"},
                {"label": "输出电压", "field": "UOut",  "color": "#FFFF00", "vmin": 0, "vmax": 35,  "unit": "V"},
                {"label": "输出电流", "field": "IOut",  "color": "#FF4040", "vmin": 0, "vmax": 10,  "unit": "A"},
            ])
        self.curve1.frame.pack(fill="both", expand=True, padx=8, pady=(8, 4))

        self.curve2 = CurveCanvas(
            right_frame, "温度",
            [
                {"label": "环境温度", "field": "TAmb",   "color": "#00FF7F", "vmin": 0, "vmax": 50,  "unit": "℃"},
                {"label": "电堆温度", "field": "TStack", "color": "#FF4040", "vmin": 0, "vmax": 120, "unit": "℃"},
            ])
        self.curve2.frame.pack(fill="both", expand=True, padx=8, pady=4)

    def _make_log(self, widget):
        """让 Text 只读但允许鼠标选择与 Ctrl+C 复制。"""
        def block(e):
            # 放行 Ctrl 组合键（复制/全选），阻止其它编辑
            if e.state & 0x4:
                return None
            return "break"
        widget.bind("<Key>", block)
        widget.bind("<<Paste>>", lambda e: "break")
        widget.bind("<<Clear>>", lambda e: "break")
        widget.bind("<<Cut>>", lambda e: "break")

    # ---------------- 串口枚举 ----------------
    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        ports.sort(key=_port_sort_key)
        self.port_combo["values"] = ports
        # 需求要求默认 COM9；若枚举到 COM9 则选中，否则保留 COM9 文本作为默认显示
        if "COM9" in ports:
            self.port_var.set("COM9")
        elif not self.port_var.get():
            self.port_var.set("COM9")

    # ---------------- 打开 / 关闭 ----------------
    def _toggle_port(self):
        if self.serial_open:
            self._close_port()
        else:
            self._open_port()

    def _open_port(self):
        port = self.port_var.get().strip()
        if not port:
            messagebox.showwarning("提示", "请选择或输入串口")
            return
        try:
            baud = int(self.baud_var.get())
        except ValueError:
            messagebox.showwarning("提示", "波特率无效")
            return
        try:
            self.ser = serial.Serial(port, baudrate=baud, timeout=0.1, write_timeout=1.0)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开 {port}:\n{e}")
            return
        self.serial_open = True
        self.recv_buffer_str = ""
        self.open_btn.config(text="关闭串口")
        self.status_var.set(f"已打开 {port} @ {baud}")
        self.read_running = True
        self.read_thread = threading.Thread(target=self._read_loop, daemon=True)
        self.read_thread.start()

    def _close_port(self):
        self._stop_loop()
        self.read_running = False
        if self.read_thread is not None:
            self.read_thread.join(timeout=1.0)
            self.read_thread = None
        if self.ser is not None:
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None
        self.serial_open = False
        self.open_btn.config(text="打开串口")
        self.status_var.set("串口已关闭")

    # ---------------- 读取线程 ----------------
    def _read_loop(self):
        while self.read_running and self.ser is not None:
            try:
                data = self.ser.read(4096)
            except Exception:
                continue
            if data:
                self.recv_queue.put(data)

    def _poll_recv_queue(self):
        try:
            while True:
                data = self.recv_queue.get_nowait()
                self._handle_recv(data)
        except Empty:
            pass
        self.root.after(80, self._poll_recv_queue)

    def _handle_recv(self, data):
        try:
            text = data.decode("utf-8", errors="replace")
        except Exception:
            text = data.decode("latin-1", errors="replace")
        self.recv_buffer_str += text

        while True:
            start_index = self.recv_buffer_str.find(self.RECV_FRAME_START)
            if start_index < 0:
                # 保留可能属于下一批数据的帧头前缀。
                keep = len(self.RECV_FRAME_START) - 1
                self.recv_buffer_str = self.recv_buffer_str[-keep:]
                return
            if start_index > 0:
                self.recv_buffer_str = self.recv_buffer_str[start_index:]

            end_index = self.recv_buffer_str.find(
                self.RECV_FRAME_END, len(self.RECV_FRAME_START)
            )
            if end_index < 0:
                return

            frame_end = end_index + len(self.RECV_FRAME_END)
            frame = self.recv_buffer_str[:frame_end]
            self.recv_buffer_str = self.recv_buffer_str[frame_end:]
            payload = frame[len(self.RECV_FRAME_START):end_index].strip()
            self._log_recv(frame)
            self._parse_recv_line(payload)

    def _parse_recv_line(self, line):
        """解析接收到的一行数据，提取字段并更新显示面板。"""
        if not line.strip():
            return
        # 建立协议 key -> 数值 的映射
        parsed = {}
        for match in re.finditer(r'(\w+)\s+(\S+)', line):
            key = match.group(1)
            value = match.group(2)
            parsed[key] = value

        for display_name, proto_key in self.DISPLAY_FIELDS:
            if proto_key in parsed:
                self.display_vars[display_name].set(parsed[proto_key])

        # 推送曲线数据
        if self.curve1 is not None:
            self.curve1.add_data(parsed)
        if self.curve2 is not None:
            self.curve2.add_data(parsed)

        # 记录数据到 xlsx
        if self.recorder is not None:
            self.recorder.add_row(parsed)

    # ---------------- 发送 ----------------
    def _manual_send(self):
        if not self.serial_open:
            messagebox.showinfo("提示", "请先打开串口")
            return
        msg = self.input_var.get()
        if not msg:
            return
        self._send(msg)
        self.input_var.set("")

    def _send(self, msg):
        if not self.serial_open or self.ser is None:
            return False
        try:
            self.ser.write((msg + "\r\n").encode("utf-8"))
        except Exception as e:
            self._log_send(f"[发送失败: {e}]")
            return False
        self._log_send(msg)
        return True

    # ---------------- 循环发送 ----------------
    def _toggle_loop(self):
        if self.loop_running:
            self._stop_loop()
        else:
            if not self.serial_open:
                messagebox.showinfo("提示", "请先打开串口再循环发送")
                return
            if not HAS_OPENPYXL:
                messagebox.showwarning("缺少依赖",
                                       "未安装 openpyxl，无法记录数据。\n请运行：pip install openpyxl")
            self.loop_running = True
            self.loop_btn.config(text="停止")
            # 启动数据记录
            if HAS_OPENPYXL:
                self.recorder = DataRecorder(self.data_dir)
                self.recorder.start()
            self._loop_send()

    def _loop_send(self):
        if not self.loop_running:
            return
        if self.serial_open and self.ser is not None:
            self._send(self.LOOP_CMD)
        try:
            period = int(self.period_var.get())
        except ValueError:
            period = 1
        period = max(1, min(10, period))
        self.loop_after_id = self.root.after(period * 1000, self._loop_send)

    def _stop_loop(self):
        self.loop_running = False
        if self.loop_after_id is not None:
            self.root.after_cancel(self.loop_after_id)
            self.loop_after_id = None
        if hasattr(self, "loop_btn"):
            self.loop_btn.config(text="循环")
        # 停止数据记录并保存
        if self.recorder is not None:
            self.recorder.stop()
            self.recorder = None

    # ---------------- 日志 ----------------
    def _timestamp(self):
        now = datetime.now()
        return f"{now.strftime('%H:%M:%S')} {now.microsecond // 1000:03d}"

    def _log_send(self, msg):
        self._append_text(self.send_text, f"{self._timestamp()}-> {msg}\n")

    def _log_recv(self, msg):
        self._append_text(self.recv_text, f"{self._timestamp()}<- {msg}\n")

    def _append_text(self, widget, text):
        widget.insert("end", text)
        widget.see("end")

    # ---------------- 关闭 ----------------
    def _on_close(self):
        self._stop_loop()
        self.read_running = False
        if self.ser is not None:
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None
        self.root.destroy()


def main():
    root = tk.Tk()
    SerialMonitorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

