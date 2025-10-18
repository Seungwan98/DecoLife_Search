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

            # 6) 매칭 & 합계 (목록 수집 포함)
            self.debug_all = []
            self.debug_selected = []

            name_lower = name_ser.fillna("").str.lower()
            kw_lower = keyword.strip().lower()

            def row_item(idx, reason: str):
                name_val = str(name_ser.iloc[idx])
                cost_raw = str(df.iloc[idx, cost_col])
                # 숫자화 값
                from math import isnan
                val = cost_ser.iloc[idx]
                cost_num = None if pd.isna(val) else float(val)
                return {"row": int(data_start + idx), "name": name_val, "cost_raw": cost_raw, "cost_num": cost_num,
                        "reason": reason}

            if kw_lower == "hdd":
                # 6-1) HDD 포함
                mask_hdd = name_lower.str.contains("hdd", na=False)

                # 6-2) HDD 미포함 중 모델코드 포함
                model_codes = [
                    "WD10EZEX", "WD20EZAZ", "WD20EZBX", "WD30EZAX", "WD40EZAX", "WD60EZAX", "WD80EAZZ", "WD80EAAZ",
                    "WD10PURZ", "WD23PURZ", "WD33PURZ", "WD43PURZ", "WD64PURZ", "WD84PURZ", "WD8001PURP", "WD101PURP",
                    "WD121PURP", "WD141PURP", "WD181PURP", "WD2003FZEX", "WD4005FZBX", "WD8002FZWX", "WD101FZBX",
                    "WD20EFPX", "WD40EFPX", "WD60EZPX", "WD80EFZZ", "WD101EFBX", "WD120EFBX", "WD2002FFSX",
                    "WD4003FFBX", "WD6003FFBX", "WD8003FFBX", "WD8005FFBX", "WD102KFBX", "WD121KFBX", "WD142KFGX",
                    "WD161KFGX", "WD181KFGX", "WD201KFGX", "WD221KFGX", "WD240KFGX", "WD10SPZX", "WD20SPZX",
                    "WD5000LPZX"
                ]
                model_codes = list(dict.fromkeys(model_codes))
                pattern = r"(" + "|".join(map(re.escape, model_codes)) + r")"
                mask_model = name_ser.str.contains(pattern, case=False, na=False)

                mask_extra = (~mask_hdd) & mask_model
                final_mask = mask_hdd | mask_extra

                # 합계/개수
                cnt_hdd = int(mask_hdd.sum())
                cnt_model_only = int(mask_extra.sum())
                matched_count = int(final_mask.sum())
                total = cost_ser[final_mask].sum(skipna=True)

                # 🔎 디버그 목록 수집 (상한 300건)
                LIMIT = 300

                def row_item(idx, reason: str):
                    name_val = str(name_ser.loc[idx])
                    cost_raw = str(df.loc[idx, df.columns[cost_col]])
                    val = cost_ser.loc[idx]
                    cost_num = None if pd.isna(val) else float(val)
                    return {"row": int(idx), "name": name_val, "cost_raw": cost_raw, "cost_num": cost_num,
                            "reason": reason}

                # 전체 후보 수집
                for idx in list(mask_hdd[mask_hdd].index)[:LIMIT]:
                    self.debug_all.append(row_item(idx, "HDD"))
                for idx in list(mask_model[mask_model].index)[:LIMIT]:
                    if not mask_hdd.loc[idx]:  # ✅ loc
                        self.debug_all.append(row_item(idx, "모델코드"))

                # 최종 선택 수집
                for idx in list(final_mask[final_mask].index)[:LIMIT]:
                    reason = "HDD" if mask_hdd.loc[idx] else "모델코드"  # ✅ loc
                    self.debug_selected.append(row_item(idx, reason))

                # 최종 선택: final_mask True
                for idx in list(final_mask[final_mask].index)[:LIMIT]:
                    # 두 경우 구분해 이유 표기
                    reason = "HDD" if mask_hdd.iloc[idx] else "모델코드"
                    self.debug_selected.append(row_item(idx, reason))

                # 7) 결과 표시
                self.result_label.config(
                    text=(
                        f"결과: {total:,.0f} 원\n"
                        f"(매칭된 항목 수: {matched_count}개 = HDD표기 {cnt_hdd}개 + 모델코드 {cnt_model_only}개)"
                    )
                )

            elif kw_lower == "ssd":
                # 6-1) SSD 포함
                mask_ssd = name_lower.str.contains("ssd", na=False)

                # 6-2) SSD 미포함 중 모델코드 포함
                ssd_models = [
                    "Green 3D", "Green SATA", "Green M.2", "SA510",
                    "SN350", "SN570", "SN580", "SN770", "SN770M",
                    "SN850X", "SN5000", "SN7100"
                ]
                ssd_models = list(dict.fromkeys(ssd_models))
                pattern = r"(" + "|".join(map(re.escape, ssd_models)) + r")"
                mask_model = name_ser.str.contains(pattern, case=False, na=False)

                mask_extra = (~mask_ssd) & mask_model
                final_mask = mask_ssd | mask_extra

                # 합계/개수
                cnt_ssd = int(mask_ssd.sum())
                cnt_model_only = int(mask_extra.sum())
                matched_count = int(final_mask.sum())
                total = cost_ser[final_mask].sum(skipna=True)

                # 🔎 디버그 목록 수집 (상한 300건)
                LIMIT = 300

                def row_item(idx, reason: str):
                    name_val = str(name_ser.loc[idx])
                    cost_raw = str(df.loc[idx, df.columns[cost_col]])
                    val = cost_ser.loc[idx]
                    cost_num = None if pd.isna(val) else float(val)
                    return {"row": int(idx), "name": name_val, "cost_raw": cost_raw, "cost_num": cost_num,
                            "reason": reason}

                for idx in list(mask_ssd[mask_ssd].index)[:LIMIT]:
                    self.debug_all.append(row_item(idx, "SSD"))

                for idx in list(mask_model[mask_model].index)[:LIMIT]:
                    if not mask_ssd.loc[idx]:  # ✅ iloc → loc
                        self.debug_all.append(row_item(idx, "모델코드"))

                for idx in list(final_mask[final_mask].index)[:LIMIT]:
                    reason = "SSD" if mask_ssd.loc[idx] else "모델코드"  # ✅ iloc → loc
                    self.debug_selected.append(row_item(idx, reason))

                # 7) 결과 표시
                self.result_label.config(
                    text=(
                        f"결과: {total:,.0f} 원\n"
                        f"(매칭된 항목 수: {matched_count}개 = SSD표기 {cnt_ssd}개 + 모델코드 {cnt_model_only}개)"
                    )
                )


            else:
                # 일반 키워드
                mask = name_lower.str.contains(re.escape(kw_lower), na=False)
                matched_count = int(mask.sum())
                total = cost_ser[mask].sum(skipna=True)

                LIMIT = 300
                # 전체 후보 = 키워드 포함 행
                for idx in list(mask[mask].index)[:LIMIT]:
                    self.debug_all.append(row_item(idx, f"키워드:{keyword}"))
                # 최종 선택 = 동일 (일반 검색은 보정 없음)
                for idx in list(mask[mask].index)[:LIMIT]:
                    self.debug_selected.append(row_item(idx, f"키워드:{keyword}"))

                self.result_label.config(
                    text=f"결과: {total:,.0f} 원\n(매칭된 항목 수: {matched_count}개)"
                )


        except Exception as e:
            messagebox.showerror("에러", f"처리 중 오류 발생:\n{e}")

    # ---------- 디버그 로그 팝업 ----------
    def show_debug(self):
        """'모델코드로만 잡힌' 항목들만 디버그 로그로 표시"""
        # 모델코드로만 잡힌 항목들만 필터링
        model_only_items = [it for it in self.debug_selected if it["reason"] == "모델코드"]

        # 표 형태 문자열 만들기
        def table_from(items, title):
            if not items:
                return f"[{title}]\n(모델코드로만 추가된 항목 없음)\n"
            lines = [f"[{title}] (총 {len(items)}건)\n"]
            lines.append(f"{'행':>6} | {'상품명':60} | {'금액(원본)':15} | {'금액(숫자)':>12}")
            lines.append("-" * 110)
            for it in items:
                cost_display = "-" if it['cost_num'] is None else f"{it['cost_num']:,.0f}"
                row = (
                    f"{it['row']:>6} | "
                    f"{it['name'][:60]:60} | "
                    f"{str(it['cost_raw'])[:15]:15} | "
                    f"{cost_display:>12}"
                )
                lines.append(row)
            return "\n".join(lines) + "\n\n"

        # 팝업창 생성
        win = tk.Toplevel(self.root)
        win.title("모델코드로만 추가된 항목들")
        win.geometry("900x600")

        txt = tk.Text(win, wrap="none", font=("Menlo", 11))
        txt.pack(fill="both", expand=True)

        txt.insert("1.0", table_from(model_only_items, "모델코드로만 추가된 항목들"))
        txt.config(state="disabled")


# 실행
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSumApp(root)
    root.mainloop()
