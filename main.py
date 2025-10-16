import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re

class ExcelSumApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 합계 계산기")
        self.root.geometry("460x320")

        self.file_path = None
        self.last_log = ""  # 디버그 로그 저장

        # --- UI 구성 ---
        tk.Label(root, text="텍스트 입력 (등록상품명 검색)").pack(pady=(12,6))
        self.entry = tk.Entry(root, width=40)
        self.entry.pack()

        btns = tk.Frame(root)
        btns.pack(pady=10)
        tk.Button(btns, text="엑셀 파일 선택", command=self.load_excel).pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="합계 계산", command=self.calculate_sum).pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="디버그 로그 보기", command=self.show_debug).pack(side=tk.LEFT, padx=6)

        self.result_label = tk.Label(root, text="결과: -", font=("Arial", 12))
        self.result_label.pack(pady=10)

    # ---------- 유틸 ----------
    def _log(self, s=""):
        self.last_log += (s + ("\n" if not s.endswith("\n") else ""))

    def _norm(self, s: str) -> str:
        """공백/괄호 정규화 + 소문자"""
        if s is None:
            return ""
        s = str(s)
        s = s.replace("（", "(").replace("）", ")")
        s = re.sub(r"\s+", "", s)
        return s.strip().lower()

    def _to_number_series(self, ser: pd.Series) -> pd.Series:
        # 쉼표/공백/원화기호 제거 후 숫자화
        ser2 = (
            ser.astype(str)
              .str.replace(",", "", regex=False)
              .str.replace(" ", "", regex=False)
              .str.replace("₩", "", regex=False)
              .str.replace("원", "", regex=False)
        )
        return pd.to_numeric(ser2, errors="coerce")

    # ---------- 파일 선택 ----------
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        self.file_path = file_path
        messagebox.showinfo("파일 선택", f"선택된 파일:\n{file_path}")

    # ---------- 합계 계산(디버그 포함) ----------
    def calculate_sum(self):
        from tkinter import messagebox
        import pandas as pd
        import re

        if not self.file_path:
            messagebox.showwarning("경고", "엑셀 파일을 먼저 선택하세요.")
            return

        keyword = self.entry.get().strip()
        if not keyword:
            messagebox.showwarning("경고", "검색할 텍스트를 입력하세요.")
            return

        # ---- 유틸 ----
        def norm(s: str) -> str:
            if s is None:
                return ""
            s = str(s)
            s = s.replace("（", "(").replace("）", ")").replace("[", "").replace("]", "")
            s = re.sub(r"\s+", "", s)
            return s.lower()

        def to_number(ser: pd.Series) -> pd.Series:
            ser = (ser.astype(str)
                   .str.replace(",", "", regex=False)
                   .str.replace(" ", "", regex=False)
                   .str.replace("₩", "", regex=False)
                   .str.replace("원", "", regex=False))
            return pd.to_numeric(ser, errors="coerce")

        try:
            # 1) 헤더 없이 통째로 읽기
            df_raw = pd.read_excel(self.file_path, header=None, engine="openpyxl")

            # 2) 타깃 헤더 정의
            #   - 상품명: '등록상품명'을 최우선. '상품명' 허용. 'ID' 포함 셀은 제외.
            #   - 금액: '할인적용가(A-B)' 최우선. 변형(공백/괄호/하이픈) 허용.
            name_priority = ["등록상품명", "상품명"]
            cost_priority = ["할인적용가(a-b)", "할인적용가", "할인 적용가", "할인적용가a-b"]

            name_row = name_col = None
            cost_row = cost_col = None

            # 3) 위에서 아래로 훑되, 더 아래쪽(두 번째 헤더 줄)을 "우선"으로 채택
            for i, row in df_raw.iterrows():
                raw = ["" if pd.isna(x) else str(x) for x in row.tolist()]
                txt = [norm(x) for x in raw]

                # (a) 상품명 후보 탐색: '등록상품명' 정확 매칭을 우선, 없으면 '상품명'
                #     단, 'id' 포함 셀은 제외
                for j, t in enumerate(txt):
                    if "id" in t:
                        continue
                    if any(norm(p) in t for p in name_priority):
                        # 더 아래 행으로 갱신(멀티헤더 시 마지막 줄을 쓰기 위함)
                        name_row, name_col = i, j

                # (b) 할인적용가(A-B) 후보 탐색
                for j, t in enumerate(txt):
                    if any(norm(p) in t for p in cost_priority):
                        cost_row, cost_col = i, j

            if name_col is None or cost_col is None:
                preview = df_raw.head(12).astype(str).fillna("").to_string(index=True)
                messagebox.showerror(
                    "에러",
                    "엑셀에서 '등록상품명' 또는 '할인적용가(A-B)' 헤더를 찾지 못했습니다.\n"
                    "상단 12줄을 확인하세요.\n\n" + preview
                )
                return

            # 4) 데이터 시작 행: 더 아래쪽 헤더의 다음 줄
            data_start = max(name_row, cost_row) + 1
            df = df_raw.iloc[data_start:].copy()

            # 5) 시리즈 추출 / 전처리
            name_ser = df.iloc[:, name_col].astype(str)
            cost_ser = to_number(df.iloc[:, cost_col])

            # 6) 키워드 매칭 & 합계
            mask = name_ser.fillna("").str.contains(re.escape(keyword), case=False, na=False)
            matched_count = int(mask.sum())
            total = cost_ser[mask].sum(skipna=True)

            # 7) 결과 표시 (항목 개수 포함)
            self.result_label.config(
                text=f"결과: {total:,.0f} 원\n(매칭된 항목 수: {matched_count}개)"
            )

            # 7) 결과 표시
            self.result_label.config(text=f"결과: {total:,.0f} 원")

        except Exception as e:
            messagebox.showerror("에러", f"처리 중 오류 발생:\n{e}")

    # ---------- 디버그 로그 팝업 ----------
    def show_debug(self):
        if not self.last_log:
            messagebox.showinfo("디버그", "아직 로그가 없습니다. 먼저 '합계 계산'을 실행하세요.")
            return
        win = tk.Toplevel(self.root)
        win.title("디버그 로그")
        win.geometry("760x520")
        txt = tk.Text(win, wrap="none")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", self.last_log)
        txt.config(state="disabled")

# 실행
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSumApp(root)
    root.mainloop()
