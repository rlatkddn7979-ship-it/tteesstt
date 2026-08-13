#!/usr/bin/env python3
"""
두원전자통신 물품 조회 도구 모음 - 통합 런처

지금까지 만든 여러 GUI 도구를 하나의 창에서 골라서 실행할 수 있게 모아놓은
런처입니다. 목록에서 도구를 고르고 "실행"을 누르면 그 도구의 창이 별도
프로세스로 새로 뜹니다 (런처 창은 그대로 열려있어서 여러 도구를 동시에
띄울 수도 있습니다).

실행 (Python으로):
    python3 scripts/launcher.py

실행파일(.exe) 하나로 묶기 (Windows에서, 이 scripts 폴더 안에서 실행):
    pip install pyinstaller openpyxl requests beautifulsoup4
    pyinstaller --onefile --noconsole --name 두원전자통신물품조회 launcher.py
    -> dist\\두원전자통신물품조회.exe 가 생성됩니다. 이 exe 파일 하나만 배포하면
       Python이 안 깔린 PC에서도 실행할 수 있습니다 (curl은 Windows 10/11에
       기본 내장되어 있어서 별도로 챙길 필요 없습니다).

주의:
    - PyInstaller 빌드는 빌드하는 그 OS용으로만 만들어집니다. Windows용 exe를
      만들려면 Windows PC에서 pyinstaller를 실행해야 합니다 (다른 OS에서
      크로스 빌드는 지원하지 않습니다).
    - 저장 파일(엑셀/CSV 등)은 항상 이 실행파일이 있는 폴더를 기준으로 저장됩니다.
"""

import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk

# PyInstaller가 정적 분석으로 찾아서 실행파일 하나에 다 묶어갈 수 있도록 미리 import해둔다.
import check_doowon_cameras_gui
import check_doowon_cameras_spec_gui
import check_doowon_cameras_byspec_gui
import check_doowon_cameras_keyword_gui
import check_doowon_cameras_bymaker_gui
import check_doowon_cameras_corp_or_maker_gui
import check_doowon_cameras_shoppingmall_gui
import check_doowon_cameras_shoppingmall_byspec_gui
import check_doowon_cameras_by_contract_gui
import check_doowon_cameras_general_gui
import check_doowon_cameras_mas_gui
import check_product_by_id_gui
import g2b_goods_detail_crawler_gui

# (도구 이름, 설명, 모듈) 목록. 순서가 런처 목록에 그대로 보인다.
TOOLS = [
    ("check_doowon_cameras_gui", "기본 조회 (제3자단가계약, 계약업체명 기준)", check_doowon_cameras_gui),
    ("check_doowon_cameras_spec_gui", "규격 진단 (품명별+규격별 개수 같이 보기)", check_doowon_cameras_spec_gui),
    ("check_doowon_cameras_byspec_gui", "규격 기준 종류 집계 (돔형/블렛형 등 자동 구분)", check_doowon_cameras_byspec_gui),
    ("check_doowon_cameras_keyword_gui", "키워드 통합검색 (품명+규격 한 번에)", check_doowon_cameras_keyword_gui),
    ("check_doowon_cameras_bymaker_gui", "제조사 기준 검색 (계약업체명 아닌 제조사로 찾기)", check_doowon_cameras_bymaker_gui),
    ("check_doowon_cameras_corp_or_maker_gui", "계약업체명 OR 제조사 (둘 중 하나만 맞아도 매칭)", check_doowon_cameras_corp_or_maker_gui),
    ("check_doowon_cameras_shoppingmall_gui", "나라장터 쇼핑몰 품목 조회 (추가선택품목 포함)", check_doowon_cameras_shoppingmall_gui),
    ("check_doowon_cameras_shoppingmall_byspec_gui", "쇼핑몰 품목 조회 + 규격 기준 종류 집계", check_doowon_cameras_shoppingmall_byspec_gui),
    ("check_doowon_cameras_by_contract_gui", "계약번호 기반 조회 (전국 스캔 없이 특정 업체 전체)", check_doowon_cameras_by_contract_gui),
    ("check_doowon_cameras_general_gui", "일반단가계약 조회", check_doowon_cameras_general_gui),
    ("check_doowon_cameras_mas_gui", "다수공급자계약(MAS) 조회", check_doowon_cameras_mas_gui),
    ("check_product_by_id_gui", "물품식별번호 직접 조회 (진단용)", check_product_by_id_gui),
    ("g2b_goods_detail_crawler_gui", "나라장터 목록정보시스템 상품 상세 크롤러", g2b_goods_detail_crawler_gui),
]

TOOL_MODULES = {name: module for name, _desc, module in TOOLS}


class LauncherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("두원전자통신 물품 조회 도구 모음")
        self.geometry("580x440")

        ttk.Label(self, text="실행할 도구를 선택하세요", font=("", 11, "bold")).pack(pady=(15, 5))

        list_frame = ttk.Frame(self, padding=10)
        list_frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(list_frame, font=("", 10))
        for _name, desc, _module in TOOLS:
            self.listbox.insert("end", desc)
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.bind("<Double-Button-1>", lambda e: self._launch_selected())

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        btn_row = ttk.Frame(self, padding=10)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="실행", command=self._launch_selected).pack(side="left")
        ttk.Label(btn_row, text="더블클릭으로도 실행할 수 있습니다").pack(side="left", padx=10)

    def _launch_selected(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("안내", "실행할 도구를 목록에서 선택해주세요.")
            return
        name, _desc, _module = TOOLS[selection[0]]

        # 별도 프로세스로 띄운다 (Tkinter는 root(Tk)를 여러 개 동시에 띄우는 걸
        # 권장하지 않아서, 이 실행파일 자신을 --tool=이름 으로 다시 실행하는 방식을 쓴다.
        # exe로 묶여있어도(PyInstaller --onefile) 그대로 동작한다).
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, f"--tool={name}"]
        else:
            cmd = [sys.executable, os.path.abspath(__file__), f"--tool={name}"]

        try:
            subprocess.Popen(cmd)
        except Exception as e:
            messagebox.showerror("오류", f"실행 중 오류가 발생했습니다: {e}")


def main():
    if len(sys.argv) > 1 and sys.argv[1].startswith("--tool="):
        tool_name = sys.argv[1].split("=", 1)[1]
        module = TOOL_MODULES.get(tool_name)
        if module is None:
            print(f"알 수 없는 도구: {tool_name}", file=sys.stderr)
            sys.exit(1)
        module.App().mainloop()
        return

    LauncherApp().mainloop()


if __name__ == "__main__":
    main()
