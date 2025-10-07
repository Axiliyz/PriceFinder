# pip install pandas openpyxl xlrd python-calamine customtkinter

import os
import re
import gc
import math
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd

# ---------- Excel reader with fallbacks ----------

def read_sheets_any(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        engines = ["openpyxl", "calamine"]
    elif ext == ".xls":
        engines = ["xlrd", "calamine"]
    else:
        engines = ["openpyxl", "xlrd", "calamine"]

    last_err = None
    for eng in engines:
        try:
            dfs = pd.read_excel(path, sheet_name=None, header=None, dtype=object, engine=eng)
            if isinstance(dfs, dict) and len(dfs) > 0:
                return dfs
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(
        f"Не удалось прочитать файл '{os.path.basename(path)}'. "
        f"Откройте его в Excel/LibreOffice и сохраните как .xlsx.\nПоследняя ошибка: {last_err}"
    )

# ---------- Text normalization & mojibake fix ----------

RE_LETTER_RU = re.compile(r"[a-zа-яё]", re.IGNORECASE)
RE_HAS_CYR = re.compile(r"[А-Яа-яЁё]")

MOJIBAKE_MARKERS = re.compile(r"[ÐÑÃÂÊËÄÅðñãâêëäåØøŸÝýÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÒÓÔÕÖ×ØÙÚÛÜÝÞß]")

def cyr_score(s: str) -> int:
    return len(RE_HAS_CYR.findall(s or ""))

def try_decode(s: str):
    # кандидаты преобразований
    candidates = [s]
    for fn in (
        lambda x: x.encode("latin1", errors="ignore").decode("utf-8", errors="ignore"),
        lambda x: x.encode("latin1", errors="ignore").decode("cp1251", errors="ignore"),
        lambda x: x.encode("cp1251", errors="ignore").decode("utf-8", errors="ignore"),
        lambda x: x.encode("utf-8", errors="ignore").decode("cp1251", errors="ignore"),
    ):
        try:
            candidates.append(fn(s))
        except Exception:
            pass
    # выбрать с максимальным числом кириллических символов; при равенстве — ближайший к исходному
    best = max(candidates, key=lambda t: (cyr_score(t), -abs(len(t) - len(s))))
    return best

def fix_mojibake(s):
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return s
    s = str(s)
    # если явные маркеры «крякозябры» — пробуем декодировать
    if MOJIBAKE_MARKERS.search(s) and cyr_score(s) < 0.5 * len(s):
        fixed = try_decode(s)
        # принимаем, если явное улучшение
        if cyr_score(fixed) > cyr_score(s):
            return fixed
    return s

def normalize_text(s: str) -> str:
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ""
    # сначала попытка починки, потом нормализация
    s = fix_mojibake(s)
    s = str(s).lower()
    s = s.replace("ё", "е")
    s = s.replace("\n", " ").replace("\r", " ").replace("_", " ")
    s = s.strip(" \t\"'`“”«»")
    s = re.sub(r"\s+", " ", s).strip()
    return s

# ---------- Price parsing & numeric heuristics ----------

def parse_number_price(val):
    """Цена: >0, без любых букв."""
    if val is None:
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        x = float(val)
        return x if x > 0 else None

    s = str(val).strip()
    if re.search(r"[A-Za-zА-Яа-яЁё]", s):
        return None

    s = s.replace("\u00A0", "").replace(" ", "")
    if re.match(r"^\d{1,3}(,\d{3})+(\.\d+)?$", s):
        s = s.replace(",", "")
    elif re.match(r"^\d+,\d+$", s):
        s = s.replace(",", ".")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        if s in ("", ".", "-", "-."):
            return None
        x = float(s)
        return x if x > 0 else None
    except Exception:
        return None

def density_scores_numeric(series: pd.Series, max_rows: int = 200):
    vals = series.iloc[:max_rows]
    raw_vals = [v for v in vals if str(v).strip() != ""]
    n = len(raw_vals)
    if n == 0:
        return {
            "n":0, "numeric_ratio":0.0, "int_ratio":0.0, "dec_ratio":0.0,
            "mean":0.0, "median":0.0, "std":0.0, "zero_ratio":0.0,
            "dominant_ratio":0.0, "long_digit_ratio":0.0, "unique_ratio":0.0
        }

    nums = []
    ints = 0; decs = 0; zeros = 0
    dominant_count = 0
    freq = {}
    long_digit = 0  # >=6 подряд цифр

    for v in raw_vals:
        s = str(v).strip()
        if re.fullmatch(r"\d{6,}", s):
            long_digit += 1

        x = parse_number_price(v)
        if x is not None:
            nums.append(x)
            if float(x).is_integer():
                ints += 1
                if x == 0:
                    zeros += 1
            else:
                decs += 1

        freq[s] = freq.get(s, 0) + 1
        dominant_count = max(dominant_count, freq[s])

    k = len(nums)
    if k:
        s_nums = pd.Series(nums)
        mean = float(s_nums.mean())
        median = float(s_nums.median())
        std = float(s_nums.std(ddof=0))
        int_ratio = ints / k
        dec_ratio = decs / k
        zero_ratio = zeros / k
    else:
        mean = median = std = 0.0
        int_ratio = dec_ratio = zero_ratio = 0.0

    unique_ratio = len(set(map(str, raw_vals))) / n if n else 0.0

    return {
        "n": n,
        "numeric_ratio": k / n,
        "int_ratio": int_ratio,
        "dec_ratio": dec_ratio,
        "mean": mean,
        "median": median,
        "std": std,
        "zero_ratio": zero_ratio,
        "dominant_ratio": dominant_count / n if n else 0.0,
        "long_digit_ratio": long_digit / n,
        "unique_ratio": unique_ratio
    }

# ---------- Column selection & header detection ----------

def select_name_price_columns_from_data(data: pd.DataFrame):
    """Определяем (name_col, price_col). Цена — в приоритете сосед справа от name."""
    n_cols = data.shape[1]
    n_check = min(n_cols, 80)

    # NAME
    name_scores = []
    for j in range(n_check):
        col = data.iloc[:, j]
        vals = col.iloc[:200]
        nonempty = [str(v) for v in vals if str(v).strip() != ""]
        if not nonempty:
            name_scores.append((-1e9, j)); continue
        letters = 0; digits = 0; total_len = 0
        for v in nonempty:
            s = normalize_text(v)
            if RE_LETTER_RU.search(s): letters += 1
            if re.search(r"\d", s): digits += 1
            total_len += len(s)
        alpha_ratio = letters / len(nonempty)
        digit_ratio = digits / len(nonempty)
        avg_len = total_len / len(nonempty)
        score = alpha_ratio*3 - digit_ratio + (avg_len/20)
        if len(nonempty) < 3: score -= 2
        name_scores.append((score, j))
    name_col = max(name_scores)[1] if name_scores else None

    # PRICE candidates
    candidates = []
    for j in range(n_check):
        if j == 0:
            candidates.append((-1e9, j, {"n":0})); continue
        dens = density_scores_numeric(data.iloc[:, j])
        if dens["numeric_ratio"] < 0.6:
            base_score = -1e9
        else:
            base_score = (
                dens["dec_ratio"]*3 +
                min(dens["median"]/300.0, 2.0) +
                min(dens["std"]/200.0, 2.0)
            )
            if dens["dominant_ratio"] >= 0.8: base_score -= 5.0
            if dens["zero_ratio"] >= 0.7:     base_score -= 4.0
            if dens["std"] == 0.0:            base_score -= 3.0
            if dens["int_ratio"] > 0.9 and dens["dec_ratio"] < 0.05 and dens["median"] <= 1000: base_score -= 3.0
            if dens["long_digit_ratio"] > 0.3 and dens["dec_ratio"] < 0.05:                      base_score -= 3.0

        if name_col is not None and j > name_col:
            distance = j - name_col - 1
            base_score -= 0.6 * distance
        else:
            base_score -= 5.0
        candidates.append((base_score, j, dens))

    # 1) жёстко сосед справа
    if name_col is not None and (name_col + 1) < n_check:
        j1 = name_col + 1
        d1 = density_scores_numeric(data.iloc[:, j1])
        def good_enough_as_price(d):
            if d.get("numeric_ratio", 0.0) < 0.5: return False
            if d.get("dominant_ratio", 0.0) >= 0.85: return False
            if d.get("std", 0.0) == 0.0: return False
            if d.get("long_digit_ratio", 0.0) > 0.4 and d.get("dec_ratio", 0.0) < 0.05: return False
            return True
        if good_enough_as_price(d1):
            return name_col, j1

    # 2) иначе — лучший справа
    candidates.sort(reverse=True, key=lambda x: x[0])

    def is_good_price(d, j, name_idx):
        if j == 0: return False
        if name_idx is not None and j <= name_idx: return False
        if d.get("numeric_ratio", 0.0) < 0.6: return False
        if d.get("dominant_ratio", 0.0) >= 0.8: return False
        if d.get("zero_ratio", 0.0) >= 0.7: return False
        if d.get("std", 0.0) == 0.0: return False
        if (d.get("dec_ratio", 0.0) >= 0.08) or (d.get("median", 0.0) >= 100):
            return True
        return False

    for score, j, d in candidates[:12]:
        if is_good_price(d, j, name_col):
            return name_col, j

    for score, j, d in candidates:
        basic_ok = (
            j != 0 and
            (name_col is None or j > name_col) and
            d.get("numeric_ratio", 0.0) >= 0.6 and
            d.get("std", 0.0) > 0.0 and
            d.get("dominant_ratio", 0.0) < 0.85
        )
        if basic_ok:
            return name_col, j

    return name_col, None

def choose_header_earliest_near_best(df: pd.DataFrame, max_try: int = 150, near_ratio: float = 0.95):
    scores = []
    limit = min(max_try, len(df)-1 if len(df) > 1 else 0)
    best_score = -1
    for h in range(limit):
        data = df.iloc[h+1:].reset_index(drop=True)
        if data.empty:
            scores.append((h, 0)); continue
        name_col, price_col = select_name_price_columns_from_data(data)
        if name_col is None or price_col is None:
            scores.append((h, 0)); continue

        good = 0
        rng = min(200, len(data))
        for i in range(rng):
            name_val = data.iat[i, name_col] if name_col < data.shape[1] else None
            price_raw = data.iat[i, price_col] if price_col < data.shape[1] else None
            if normalize_text(name_val) == "":
                continue
            if parse_number_price(price_raw) is None:
                continue
            good += 1

        scores.append((h, good))
        if good > best_score:
            best_score = good

    if best_score <= 0:
        return 0

    threshold = best_score * near_ratio
    for h, sc in scores:
        if sc >= threshold:
            return h
    return max(scores, key=lambda t: t[1])[0]

# ---------- App ----------

class ExcelNamePriceApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.title("Price Finder")
        self.geometry("1280x780")

        self.df_all = pd.DataFrame()
        self.df_view = pd.DataFrame()
        self.loaded_file_count = 0
        self.loaded_sheet_count = 0

        top = ctk.CTkFrame(self); top.pack(fill="x", padx=12, pady=(12,6))
        self.load_btn = ctk.CTkButton(top, text="Загрузить Excel-файлы", command=self.load_excels); self.load_btn.pack(side="left", padx=8, pady=8)
        self.query_entry = ctk.CTkEntry(top, width=420, placeholder_text="Поиск по названию"); self.query_entry.pack(side="left", padx=(0,8), pady=8)
        self.query_entry.bind("<Return>", lambda e: self.search())
        self.search_btn = ctk.CTkButton(top, text="Искать", command=self.search, state="disabled"); self.search_btn.pack(side="left", padx=(0,8), pady=8)
        self.reset_btn = ctk.CTkButton(top, text="Сбросить фильтр", command=self.reset_view, state="disabled"); self.reset_btn.pack(side="left", padx=(0,8), pady=8)
        self.save_btn = ctk.CTkButton(top, text="Сохранить результат (CSV)", command=self.save_csv, state="disabled"); self.save_btn.pack(side="left", padx=(0,8), pady=8)
        self.purge_btn = ctk.CTkButton(top, text="Очистить файлы из памяти", command=self.purge_memory, state="disabled"); self.purge_btn.pack(side="left", padx=(0,8), pady=8)
        self.info_label = ctk.CTkLabel(top, text="Файлы не загружены", anchor="w"); self.info_label.pack(side="left", padx=12, pady=8)

        table_frame = ctk.CTkFrame(self); table_frame.pack(fill="both", expand=True, padx=12, pady=(6,12))
        self._native = tk.Frame(table_frame); self._native.pack(fill="both", expand=True)

        import tkinter.ttk as ttk
        self.ttk = ttk
        self.tree = ttk.Treeview(self._native, show="headings"); self.tree.pack(side="left", fill="both", expand=True)
        self.vsb = tk.Scrollbar(self._native, orient="vertical", command=self.tree.yview)
        self.hsb = tk.Scrollbar(self._native, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.vsb.pack(side="right", fill="y"); self.hsb.pack(side="bottom", fill="x")

        style = ttk.Style(); style.configure("Treeview", rowheight=24)

    def load_excels(self):
        filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        paths = filedialog.askopenfilenames(title="Выберите Excel-файлы", filetypes=filetypes)
        if not paths:
            return

        records = []
        total_sheets = 0
        try:
            for path in paths:
                try:
                    sheets_dict = read_sheets_any(path)
                except Exception as e:
                    messagebox.showerror("Ошибка чтения", str(e))
                    continue

                for sheet_name, df in sheets_dict.items():
                    if df is None or df.empty:
                        continue

                    header_row = choose_header_earliest_near_best(df, max_try=150, near_ratio=0.95)

                    # для детектора колонок — блок ниже шапки
                    data_for_detection = df.iloc[header_row+1:].reset_index(drop=True)
                    if data_for_detection is None or data_for_detection.empty:
                        continue
                    name_col, price_col = select_name_price_columns_from_data(data_for_detection)
                    if name_col is None or price_col is None:
                        continue

                    # для извлечения строк — включая шапку (чтобы не потерять первую позицию)
                    data_for_rows = df.iloc[header_row:].reset_index(drop=True)

                    for idx, row in data_for_rows.iterrows():
                        name_val  = row[name_col]  if name_col  < len(row) else None
                        price_raw = row[price_col] if price_col < len(row) else None

                        # пропускаем только полностью пустые
                        if (normalize_text(name_val) == "") and (str(price_raw).strip() == ""):
                            continue

                        price_val = parse_number_price(price_raw)
                        row_index = header_row + idx + 1  # 1-based

                        records.append({
                            "name": normalize_text(name_val),   # уже с автопочинкой mojibake
                            "price": price_val,
                            "source_file": os.path.basename(path),
                            "sheet": str(sheet_name),
                            "row_index": int(row_index),
                        })

                    total_sheets += 1

            if not records:
                messagebox.showwarning("Пусто", "Не удалось выделить Название/Цену из выбранных файлов.")
                return

            self.df_all = pd.DataFrame.from_records(records)
            self.df_all["_search_text"] = self.df_all["name"].astype(str)
            self.df_view = self.df_all.drop(columns=["_search_text"])
            self.loaded_file_count = len(paths)
            self.loaded_sheet_count = total_sheets

            del records; gc.collect()

            self.populate_tree(self.df_view)
            self.info_label.configure(
                text=f"Загружено файлов: {self.loaded_file_count} | листов: {self.loaded_sheet_count} | строк: {len(self.df_all)}"
            )
            self.search_btn.configure(state="normal")
            self.reset_btn.configure(state="normal")
            self.save_btn.configure(state="normal")
            self.purge_btn.configure(state="normal")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файлы:\n{e}")

    def populate_tree(self, df: pd.DataFrame):
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())

        if df is None or df.empty:
            self.tree["columns"] = []
            return

        preferred = ["name", "price", "source_file", "sheet", "row_index"]
        cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col)
            width = 120
            if col in ("name",):
                width = 420
            elif col in ("source_file", "sheet"):
                width = 160
            elif col in ("price", "row_index"):
                width = 120
            self.tree.column(col, width=width, anchor="w", stretch=True)

        def fmt(val):
            if pd.isna(val) or val is None:
                return ""
            if isinstance(val, float):
                return f"{val:.2f}"
            s = str(val)
            return s if len(s) <= 500 else s[:497] + "..."

        batch = []
        for _, row in df.iterrows():
            values = [fmt(row.get(c, "")) for c in cols]
            batch.append(values)
            if len(batch) >= 1000:
                for v in batch:
                    self.tree.insert("", "end", values=v)
                batch.clear()
        for v in batch:
            self.tree.insert("", "end", values=v)

    def search(self):
        if self.df_all is None or self.df_all.empty:
            messagebox.showinfo("Нет данных", "Сначала загрузите файлы.")
            return
        q = normalize_text(self.query_entry.get())
        if not q:
            messagebox.showinfo("Поиск", "Введите текст для поиска.")
            return

        terms = [t for t in q.split() if t]
        mask = pd.Series(True, index=self.df_all.index)
        for t in terms:
            mask &= self.df_all["_search_text"].str.contains(re.escape(t), na=False)

        result = self.df_all[mask].drop(columns=["_search_text"])
        self.df_view = result
        self.populate_tree(self.df_view)
        self.info_label.configure(text=f"Найдено строк: {len(self.df_view)} по запросу: '{q}'")

    def reset_view(self):
        if self.df_all is None or self.df_all.empty:
            return
        self.df_view = self.df_all.drop(columns=["_search_text"])
        self.populate_tree(self.df_view)
        self.info_label.configure(text=f"Всего строк: {len(self.df_view)} (сброшен фильтр)")
        self.query_entry.delete(0, "end")

    def save_csv(self):
        if self.df_view is None or self.df_view.empty:
            messagebox.showinfo("Нет данных", "Нечего сохранять.")
            return
        path = filedialog.asksaveasfilename(
            title="Сохранить результат",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not path:
            return
        try:
            self.df_view.to_csv(path, index=False)
            messagebox.showinfo("Готово", f"Сохранено: {path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    def purge_memory(self):
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        if isinstance(self.df_all, pd.DataFrame) and "_search_text" in self.df_all.columns:
            try:
                self.df_all.drop(columns=["_search_text"], inplace=True, errors="ignore")
            except Exception:
                pass
        self.df_all = pd.DataFrame(); self.df_view = pd.DataFrame()
        self.loaded_file_count = 0; self.loaded_sheet_count = 0
        self.query_entry.delete(0, "end")
        self.search_btn.configure(state="disabled"); self.reset_btn.configure(state="disabled")
        self.save_btn.configure(state="disabled"); self.purge_btn.configure(state="disabled")
        gc.collect()
        self.info_label.configure(text="Память очищена. Файлы выгружены из программы.")

if __name__ == "__main__":
    app = ExcelNamePriceApp()
    app.mainloop()
