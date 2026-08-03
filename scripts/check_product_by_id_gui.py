#!/usr/bin/env python3
"""
check_product_by_id.py를 위한 간단한 GUI (진단용 - 물품식별번호 직접 조회).

물품식별번호/서비스키를 화면에서 입력하고 "조회 시작" 버튼을 누르면
결과가 창 안에 그대로 출력된다. 업체명/날짜 필터는 걸지 않는다.
파이썬 표준 라이브러리(tkinter)만 사용하므로 별도 설치가 필요 없다.

실행:
    python3 scripts/check_product_by_id_gui.py
"""

import datetime
import os
import queue
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

import check_product_by_id as core


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("물품식별번호 직접 조회 (진단용)")
        self.geometry("720x560")
        self.log_queue = queue.Queue()
        self.worker = None

        self._build_form()
        self._build_log_area()
        self.after(100, self._drain_log_queue)

    def _build_form(self):
        form = ttk.Frame(self, padding=10)
        form.pack(fill="x")

        self.product_id_var = tk.StringVar(value="")
        today = datetime.date.today()
        self.begin_date_var = tk.StringVar(
            value=(today - datetime.timedelta(days=3 * 365)).strftime("%Y%m%d")
        )
        self.end_date_var = tk.StringVar(value=today.strftime("%Y%m%d"))
        self.service_key_var = tk.StringVar(value=os.environ.get("DATA_GO_KR_SERVICE_KEY", ""))

        rows = [
            ("물품식별번호", self.product_id_var),
            ("조회 시작일(YYYYMMDD)", self.begin_date_var),
            ("조회 종료일(YYYYMMDD)", self.end_date_var),
        ]
        for i, (label, var) in enumerate(rows):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", pady=3)
            ttk.Entry(form, textvariable=var, width=40).grid(
                row=i, column=1, sticky="we", pady=3, columnspan=2
            )

        r = len(rows)
        ttk.Label(form, text="서비스키").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Entry(form, textvariable=self.service_key_var, width=40, show="*").grid(
            row=r, column=1, sticky="we", pady=3, columnspan=2
        )

        form.columnconfigure(1, weight=1)

        r += 1
        btn_row = ttk.Frame(form)
        btn_row.grid(row=r, column=0, columnspan=3, pady=(10, 0), sticky="w")
        self.run_button = ttk.Button(btn_row, text="조회 시작", command=self._start_query)
        self.run_button.pack(side="left")
        self.status_var = tk.StringVar(value="대기 중")
        ttk.Label(btn_row, textvariable=self.status_var).pack(side="left", padx=10)

    def _build_log_area(self):
        frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="조회 결과").pack(anchor="w")
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

    def _start_query(self):
        if self.worker and self.worker.is_alive():
            return

        product_id = self.product_id_var.get().strip()
        begin_date = self.begin_date_var.get().strip()
        end_date = self.end_date_var.get().strip()
        service_key = self.service_key_var.get().strip()

        if not service_key:
            messagebox.showerror("오류", "서비스키를 입력해주세요.")
            return
        if not product_id:
            messagebox.showerror("오류", "물품식별번호를 입력해주세요.")
            return

        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")
        self.status_var.set("조회 중...")
        self.run_button.state(["disabled"])

        self.worker = threading.Thread(
            target=self._run_in_background,
            args=(product_id, begin_date, end_date, service_key),
            daemon=True,
        )
        self.worker.start()

    def _run_in_background(self, product_id, begin_date, end_date, service_key):
        try:
            result = core.lookup_product(product_id, service_key, begin_date, end_date, log=self._log)
            self.after(0, self._on_done, result)
        except Exception as e:
            self.after(0, self._on_error, e)

    def _on_done(self, result):
        self.run_button.state(["!disabled"])
        if result["ok"] and result["items"]:
            self.status_var.set(f"완료 - {len(result['items'])}건 발견")
        elif result["ok"]:
            self.status_var.set("완료 - 0건 (해당 물품 없음)")
        else:
            self.status_var.set("실패")
            messagebox.showwarning("결과 없음", "조회에 실패했습니다. 로그를 확인해주세요.")

    def _on_error(self, error):
        self.run_button.state(["!disabled"])
        self.status_var.set("오류")
        self._log(f"오류 발생: {error}")
        messagebox.showerror("오류", str(error))


if __name__ == "__main__":
    App().mainloop()
