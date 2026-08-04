#!/usr/bin/env python3
"""
g2b_goods_detail_crawler.py를 위한 간단한 GUI.

물품식별번호(goodsIdntfcNo) 목록을 화면의 텍스트 박스에 직접 붙여넣어서
쉽게 바꿔가며 조회할 수 있다. 한 줄에 하나씩(또는 쉼표로 구분) 넣으면 된다.
"조회 시작" 버튼을 누르면 진행 로그가 창에 표시되고, 끝나면 CSV/TXT/JSON
세 가지 파일로 저장된다.

실행:
    pip install requests beautifulsoup4
    python3 scripts/g2b_goods_detail_crawler_gui.py
"""

import queue
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

import g2b_goods_detail_crawler as core


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("나라장터 상품 상세정보 조회")
        self.geometry("760x640")
        self.log_queue = queue.Queue()
        self.worker = None

        self._build_form()
        self._build_log_area()
        self.after(100, self._drain_log_queue)

    def _build_form(self):
        form = ttk.Frame(self, padding=10)
        form.pack(fill="x")

        self.goods_clsfc_no_var = tk.StringVar(value=core.FIXED_GOODS_CLSFC_NO)
        self.delay_var = tk.StringVar(value=str(core.REQUEST_DELAY))

        ttk.Label(form, text="물품분류번호(goodsClsfcNo)").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(form, textvariable=self.goods_clsfc_no_var, width=20).grid(
            row=0, column=1, sticky="w", pady=3
        )

        ttk.Label(form, text="요청 간 딜레이(초)").grid(row=0, column=2, sticky="w", padx=(20, 0), pady=3)
        ttk.Entry(form, textvariable=self.delay_var, width=8).grid(row=0, column=3, sticky="w", pady=3)

        ttk.Label(
            form,
            text="물품식별번호 목록 (한 줄에 하나씩, 또는 쉼표로 구분해서 붙여넣기)",
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(10, 3))

        self.ids_text = scrolledtext.ScrolledText(form, height=14, wrap="none")
        self.ids_text.grid(row=2, column=0, columnspan=4, sticky="we")
        self.ids_text.insert("1.0", "\n".join(core.GOODS_IDNTFC_NO_LIST))

        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)

        btn_row = ttk.Frame(form)
        btn_row.grid(row=3, column=0, columnspan=4, pady=(10, 0), sticky="w")
        self.run_button = ttk.Button(btn_row, text="조회 시작", command=self._start_crawl)
        self.run_button.pack(side="left")
        self.count_var = tk.StringVar(value=f"{len(core.GOODS_IDNTFC_NO_LIST)}건 입력됨")
        ttk.Label(btn_row, textvariable=self.count_var).pack(side="left", padx=10)
        self.status_var = tk.StringVar(value="대기 중")
        ttk.Label(btn_row, textvariable=self.status_var).pack(side="left", padx=10)

        self.ids_text.bind("<KeyRelease>", self._update_count)

    def _update_count(self, _event=None):
        ids = core.parse_goods_idntfc_no_list(self.ids_text.get("1.0", "end"))
        self.count_var.set(f"{len(ids)}건 입력됨")

    def _build_log_area(self):
        frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="진행 로그").pack(anchor="w")
        self.log_widget = scrolledtext.ScrolledText(frame, state="disabled", wrap="word")
        self.log_widget.pack(fill="both", expand=True)

    def _log(self, message):
        # 백그라운드 스레드에서 호출되므로 큐에 넣고 메인 스레드에서 위젯에 반영한다.
        self.log_queue.put(str(message))

    def _drain_log_queue(self):
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_widget.configure(state="normal")
            self.log_widget.insert("end", message + "\n")
            self.log_widget.see("end")
            self.log_widget.configure(state="disabled")
        self.after(100, self._drain_log_queue)

    def _start_crawl(self):
        if self.worker and self.worker.is_alive():
            return

        goods_clsfc_no = self.goods_clsfc_no_var.get().strip()
        ids = core.parse_goods_idntfc_no_list(self.ids_text.get("1.0", "end"))

        if not goods_clsfc_no:
            messagebox.showerror("오류", "물품분류번호(goodsClsfcNo)를 입력해주세요.")
            return
        if not ids:
            messagebox.showerror("오류", "물품식별번호를 한 줄에 하나씩 입력해주세요.")
            return
        try:
            delay = float(self.delay_var.get().strip())
        except ValueError:
            messagebox.showerror("오류", "요청 간 딜레이는 숫자로 입력해주세요 (예: 0.5).")
            return

        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")
        self.status_var.set(f"조회 중... (0/{len(ids)})")
        self.run_button.state(["disabled"])

        self.worker = threading.Thread(
            target=self._run_in_background,
            args=(goods_clsfc_no, ids, delay),
            daemon=True,
        )
        self.worker.start()

    def _run_in_background(self, goods_clsfc_no, ids, delay):
        try:
            result = core.run_crawl(goods_clsfc_no, ids, request_delay=delay, log=self._log)
            paths = core.write_outputs(result["rows"], result["fail_list"], log=self._log)
            self.after(0, self._on_done, result, paths)
        except Exception as e:
            self.after(0, self._on_error, e)

    def _on_done(self, result, paths):
        self.run_button.state(["!disabled"])
        ok_count = len(result["rows"])
        fail_count = len(result["fail_list"])
        self.status_var.set(f"완료 - {ok_count}건 성공, {fail_count}건 실패")
        messagebox.showinfo(
            "완료",
            f"성공: {ok_count}건\n실패: {fail_count}건\n\n"
            f"저장 위치:\n{paths['table']}\n{paths['txt']}\n{paths['json']}",
        )

    def _on_error(self, error):
        self.run_button.state(["!disabled"])
        self.status_var.set("오류")
        self._log(f"오류 발생: {error}")
        messagebox.showerror("오류", str(error))


if __name__ == "__main__":
    App().mainloop()
