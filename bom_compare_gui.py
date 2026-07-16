from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from bom_compare_core import (
    DEFAULT_OUTPUT_DIR_NAME,
    CompareError,
    ComparisonRequest,
    ComparisonResult,
    WorkbookInspection,
    inspect_workbook,
    compare_workbooks,
)


class BomCompareApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BOM 比对工具")
        self.geometry("980x760")
        self.minsize(900, 680)

        self.left_file_var = tk.StringVar()
        self.right_file_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(
            value=str(Path(__file__).resolve().parent / DEFAULT_OUTPUT_DIR_NAME)
        )

        self.left_inspection: WorkbookInspection | None = None
        self.right_inspection: WorkbookInspection | None = None
        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self._build_layout()

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        file_frame = ttk.LabelFrame(self, text="文件选择")
        file_frame.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 8))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="左侧 BOM").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.left_file_var).grid(
            row=0, column=1, padx=8, pady=8, sticky="ew"
        )
        ttk.Button(file_frame, text="选择文件", command=self.select_left_file).grid(
            row=0, column=2, padx=8, pady=8
        )

        ttk.Label(file_frame, text="右侧 BOM").grid(row=1, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.right_file_var).grid(
            row=1, column=1, padx=8, pady=8, sticky="ew"
        )
        ttk.Button(file_frame, text="选择文件", command=self.select_right_file).grid(
            row=1, column=2, padx=8, pady=8
        )

        ttk.Label(file_frame, text="输出目录").grid(row=2, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_dir_var).grid(
            row=2, column=1, padx=8, pady=8, sticky="ew"
        )
        ttk.Button(file_frame, text="选择目录", command=self.select_output_dir).grid(
            row=2, column=2, padx=8, pady=8
        )

        action_frame = ttk.Frame(self)
        action_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(4, 12))
        action_frame.columnconfigure(0, weight=1)
        action_frame.rowconfigure(1, weight=1)

        button_bar = ttk.Frame(action_frame)
        button_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        button_bar.columnconfigure(1, weight=1)

        self.compare_button = ttk.Button(button_bar, text="开始比对", command=self.start_compare)
        self.compare_button.grid(row=0, column=0, padx=(0, 8))

        self.status_label = ttk.Label(button_bar, text="先选择两个 xlsx 文件。")
        self.status_label.grid(row=0, column=1, sticky="w")

        self.log_text = ScrolledText(action_frame, wrap=tk.WORD, font=("Menlo", 12))
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

    def log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message.rstrip() + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def set_status(self, message: str) -> None:
        self.status_label.configure(text=message)

    def select_left_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择左侧 BOM 文件",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if file_path:
            self.left_file_var.set(file_path)
            self._refresh_inspections()

    def select_right_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择右侧 BOM 文件",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if file_path:
            self.right_file_var.set(file_path)
            self._refresh_inspections()

    def select_output_dir(self) -> None:
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir_var.set(directory)

    def _refresh_inspections(self) -> None:
        left_path = self.left_file_var.get().strip()
        right_path = self.right_file_var.get().strip()
        if not left_path or not right_path:
            return

        try:
            self.left_inspection = inspect_workbook(Path(left_path))
            self.right_inspection = inspect_workbook(Path(right_path))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("读取失败", str(exc))
            self.set_status("文件读取失败，请检查 xlsx 文件。")
            return

        self._log_inspection_summary()
        self.set_status("文件已加载，直接点“开始比对”就行。")

    def _log_inspection_summary(self) -> None:
        if not self.left_inspection or not self.right_inspection:
            return

        self.log("已加载工作簿：")
        for tag, inspection in (("左侧", self.left_inspection), ("右侧", self.right_inspection)):
            self.log(f"{tag}：{inspection.workbook_path}")
            for sheet in inspection.sheets:
                if sheet.is_candidate:
                    self.log(
                        f"  - {sheet.sheet_name}：自动识别表头行 {sheet.detected_header_row}，"
                        f"主键列 {sheet.detected_key_name}"
                    )
                else:
                    self.log(f"  - {sheet.sheet_name}：未自动识别，原因：{sheet.detection_error}")

    def _build_request(self) -> ComparisonRequest:
        left_path = Path(self.left_file_var.get().strip())
        right_path = Path(self.right_file_var.get().strip())
        output_dir = Path(self.output_dir_var.get().strip())

        if not left_path.name or not right_path.name:
            raise CompareError("请先选择左右两个 xlsx 文件。")

        return ComparisonRequest(
            left_path=left_path,
            right_path=right_path,
            output_dir=output_dir,
            compare_all_common_sheets=True,
            ignore_headers=("序号", "层次"),
        )

    def start_compare(self) -> None:
        try:
            request = self._build_request()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("参数错误", str(exc))
            return

        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("处理中", "当前已有任务在运行，稍等它跑完。")
            return

        self.compare_button.configure(state="disabled")
        self.set_status("正在比对，请稍候……")
        self.log("开始执行 BOM 比对……")

        self.worker_thread = threading.Thread(
            target=self._run_compare_worker,
            args=(request,),
            daemon=True,
        )
        self.worker_thread.start()
        self.after(200, self._poll_result_queue)

    def _run_compare_worker(self, request: ComparisonRequest) -> None:
        try:
            result = compare_workbooks(request)
            self.result_queue.put(("success", result))
        except Exception as exc:  # noqa: BLE001
            self.result_queue.put(("error", exc))

    def _poll_result_queue(self) -> None:
        try:
            status, payload = self.result_queue.get_nowait()
        except queue.Empty:
            if self.worker_thread and self.worker_thread.is_alive():
                self.after(200, self._poll_result_queue)
            return

        self.compare_button.configure(state="normal")
        if status == "success":
            result = payload
            assert isinstance(result, ComparisonResult)
            self.log(result.summary_text())
            self.set_status("比对完成，结果文件已经生成。")
            messagebox.showinfo(
                "完成",
                "BOM 比对完成。\n\n"
                f"左侧结果：{result.left_output_path}\n"
                f"右侧结果：{result.right_output_path}",
            )
            return

        error = payload
        assert isinstance(error, Exception)
        self.log(f"执行失败：{error}")
        self.set_status("比对失败，请检查参数或表格结构。")
        messagebox.showerror("执行失败", str(error))


def run_app() -> None:
    app = BomCompareApp()
    app.mainloop()
