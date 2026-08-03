#!/usr/bin/env python3
"""
check_doowon_cameras_shoppingmall.py를 위한 간단한 GUI.

getShoppingMallPrdctInfoList(나라장터 쇼핑몰 품목 정보 목록 조회) 오퍼레이션으로
조회한다. 계약업체명 또는 제조사 중 하나라도 입력한 이름을 포함하면 매칭되는
물품을 찾는다. 업체명/품명/규격/조회기간/서비스키를 화면에서 입력하고
"조회 시작" 버튼을 누르면 결과와 진행 로그를 창 안에서 볼 수 있다.
파이썬 표준 라이브러리(tkinter)만 사용하므로 별도 설치가 필요 없다.

실행:
    python3 scripts/check_doowon_cameras_shoppingmall_gui.py
"""

import datetime
import os
import queue
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

import check_doowon_cameras_shoppingmall as core


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("두원전자통신 물품 조회 (나라장터 쇼핑몰 품목 조회)")
        self.geometry("720x560")
        self.log_queue = queue.Queue()
        self.worker = None

        self._build_form()
        self._build_log_area()
        self.after(100, self._drain_log_queue)

    def _build_form(self):
        form = ttk.Frame(self, padding=10)
        form.pack(fill="x")

        self.corp_name_var = tk.StringVar(value="두원전자통신")
        self.category_var = tk.StringVar(value="보안용카메라")
        self.spec_var = tk.StringVar(value="")
        today = datetime.date.today()
        self.begin_date_var = tk.StringVar(
            value=(today - datetime.timedelta(days=365)).strftime("%Y%m%d")
        )
        self.end_date_var = tk.StringVar(value=today.strftime("%Y%m%d"))
        self.service_key_var = tk.StringVar(value=os.environ.get("DATA_GO_KR_SERVICE_KEY", ""))

        rows = [
            ("업체명 (계약업체명 또는 제조사, 둘 중 하나만 맞아도 매칭)", self.corp_name_var, None),
            ("품명 (쉼표로 여러 개, 비우면 전국 조회-느림)", self.category_var, None),
            ("규격 (쉼표로 여러 개, 예: 200만화소,4배줌,블렛형)", self.spec_var, None),
            ("조회 시작일(YYYYMMDD)", self.begin_date_var, None),
            ("조회 종료일(YYYYMMDD)", self.end_date_var, None),
        ]
        for i, (label, var, _) in enumerate(rows):
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

    def _start_query(self):
        if self.worker and self.worker.is_alive():
            return

        corp_name = self.corp_name_var.get().strip()
        category = self.category_var.get().strip()
        spec = self.spec_var.get().strip()
        begin_date = self.begin_date_var.get().strip()
        end_date = self.end_date_var.get().strip()
        service_key = self.service_key_var.get().strip()

        if not service_key:
            messagebox.showerror("오류", "서비스키를 입력해주세요.")
            return
        if not corp_name:
            messagebox.showerror("오류", "업체명을 입력해주세요.")
            return
        if not category:
            proceed = messagebox.askyesno(
                "확인",
                "품명을 비워두면 전국 모든 회사의 데이터를 조회합니다(느릴 수 있음). 계속하시겠습니까?",
            )
            if not proceed:
                return

        self.log_widget.configure(state="normal")
        self.log_widget.delete("1.0", "end")
        self.log_widget.configure(state="disabled")
        self.status_var.set("조회 중...")
        self.run_button.state(["disabled"])

        self.worker = threading.Thread(
            target=self._run_in_background,
            args=(corp_name, category, spec, begin_date, end_date, service_key),
            daemon=True,
        )
        self.worker.start()

    def _run_in_background(self, corp_name, category, spec, begin_date, end_date, service_key):
        try:
            result = core.run_query(
                corp_name, category, spec, begin_date, end_date, service_key, log=self._log
            )
            saved_path = None
            if result["ok"]:
                output_path = core.default_output_filename(corp_name, category, spec)
                saved_path = core.write_output(
                    result["camera_items"], corp_name, category, spec, begin_date, end_date,
                    output_path, log=self._log,
                )
            self.after(0, self._on_done, result, saved_path)
        except Exception as e:
            self.after(0, self._on_error, e)

    def _on_done(self, result, saved_path):
        self.run_button.state(["!disabled"])
        if result["ok"]:
            self.status_var.set(f"완료 - {len(result['distinct_names'])}종류")
            messagebox.showinfo(
                "완료",
                f"매칭 물품 수: {len(result['camera_items'])}건\n"
                f"종류 수: {len(result['distinct_names'])}개\n"
                + (f"저장 위치: {saved_path}" if saved_path else ""),
            )
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
