import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import openpyxl

class KaveKishCostApp:
    def __init__(self, root):
        self.root = root
        self.root.title("سیستم آنالیز و برآورد قیمت گیربکس و قطعات (مطابق اکسل كاوه كيش)")
        self.root.geometry("1100x750")
        self.root.configure(bg="#1e1e1e")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background="#1e1e1e", foreground="#ffffff", fieldbackground="#2d2d2d")
        style.configure("TLabel", font=("Tahoma", 9), background="#1e1e1e", foreground="#ffffff")
        style.configure("TButton", font=("Tahoma", 9, "bold"), background="#0e639c", foreground="#ffffff")

        # Header
        header = tk.Label(root, text="⚙️ سیستم آنالیز متریال، ساخت و قیمت‌گذاری گیربکس", font=("Tahoma", 14, "bold"), bg="#252526", fg="#007acc", pady=10)
        header.pack(fill="x")

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Inputs Frame
        inputs_frame = ttk.LabelFrame(main_frame, text=" افزودن قطعه جدید ", padding="10")
        inputs_frame.pack(fill="x", pady=5)

        # Row 1
        ttk.Label(inputs_frame, text="نام قطعه:").grid(row=0, column=0, sticky="w", pady=2)
        self.name_entry = ttk.Entry(inputs_frame, width=15)
        self.name_entry.insert(0, "GEAR")
        self.name_entry.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(inputs_frame, text="جنس/متریال:").grid(row=0, column=2, sticky="w", pady=2)
        self.mat_combo = ttk.Combobox(inputs_frame, values=["CK45", "ST52", "1.6582", "CuSn12Ni2", "St52-3"], width=12)
        self.mat_combo.current(0)
        self.mat_combo.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(inputs_frame, text="قطر (mm):").grid(row=0, column=4, sticky="w", pady=2)
        self.dia_entry = ttk.Entry(inputs_frame, width=10)
        self.dia_entry.insert(0, "200")
        self.dia_entry.grid(row=0, column=5, padx=5, pady=2)

        ttk.Label(inputs_frame, text="طول (mm):").grid(row=0, column=6, sticky="w", pady=2)
        self.len_entry = ttk.Entry(inputs_frame, width=10)
        self.len_entry.insert(0, "100")
        self.len_entry.grid(row=0, column=7, padx=5, pady=2)

        # Row 2 Rates
        ttk.Label(inputs_frame, text="قیمت متریال (ریال/کیلو):").grid(row=1, column=0, sticky="w", pady=2)
        self.mat_rate_entry = ttk.Entry(inputs_frame, width=15)
        self.mat_rate_entry.insert(0, "3000000")
        self.mat_rate_entry.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(inputs_frame, text="هزینه تراش (ریال):").grid(row=1, column=2, sticky="w", pady=2)
        self.turning_entry = ttk.Entry(inputs_frame, width=12)
        self.turning_entry.insert(0, "4000000")
        self.turning_entry.grid(row=1, column=3, padx=5, pady=2)

        ttk.Label(inputs_frame, text="هزینه دنده (ریال):").grid(row=1, column=4, sticky="w", pady=2)
        self.gear_cut_entry = ttk.Entry(inputs_frame, width=10)
        self.gear_cut_entry.insert(0, "2000000")
        self.gear_cut_entry.grid(row=1, column=5, padx=5, pady=2)

        ttk.Label(inputs_frame, text="تعداد:").grid(row=1, column=6, sticky="w", pady=2)
        self.qty_entry = ttk.Entry(inputs_frame, width=10)
        self.qty_entry.insert(0, "1")
        self.qty_entry.grid(row=1, column=7, padx=5, pady=2)

        # Action Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=5)

        add_btn = ttk.Button(btn_frame, text="➕ افزودن قطعه به جدول", command=self.add_item)
        add_btn.pack(side="left", padx=5)

        clear_btn = ttk.Button(btn_frame, text="🗑️ پاک کردن همه", command=self.clear_all)
        clear_btn.pack(side="left", padx=5)

        excel_btn = ttk.Button(btn_frame, text="📥 خروجی اکسل دقیق (.xlsx)", command=self.export_excel)
        excel_btn.pack(side="right", padx=5)

        # Table Display
        cols = ("name", "mat", "dia", "len", "qty", "weight", "mat_cost", "machining_cost", "total")
        self.tree = ttk.Treeview(main_frame, columns=cols, show="headings", height=10)
        
        self.tree.heading("name", text="نام قطعه")
        self.tree.heading("mat", text="جنس")
        self.tree.heading("dia", text="قطر (mm)")
        self.tree.heading("len", text="طول (mm)")
        self.tree.heading("qty", text="تعداد")
        self.tree.heading("weight", text="وزن (kg)")
        self.tree.heading("mat_cost", text="قیمت متریال (ریال)")
        self.tree.heading("machining_cost", text="اجرت تراش و دنده (ریال)")
        self.tree.heading("total", text="قیمت ساخت کل")

        for c in cols:
            self.tree.column(c, width=110, anchor="center")

        self.tree.pack(fill="both", expand=True, pady=5)

        # Summary Frame
        summary_frame = ttk.LabelFrame(main_frame, text=" خلاصه آنالیز گیربکس (جمع کل + ۳۰٪ سود و بالاسری) ", padding="10")
        summary_frame.pack(fill="x", pady=5)

        self.total_weight_lbl = tk.Label(summary_frame, text="وزن کل: ۰ کیلوگرم", font=("Tahoma", 10, "bold"), bg="#1e1e1e", fg="#ffffff")
        self.total_weight_lbl.grid(row=0, column=0, padx=20)

        self.total_cost_lbl = tk.Label(summary_frame, text="قیمت ساخت کل: ۰ ریال", font=("Tahoma", 10, "bold"), bg="#1e1e1e", fg="#ffffff")
        self.total_cost_lbl.grid(row=0, column=1, padx=20)

        self.final_price_lbl = tk.Label(summary_frame, text="قیمت اعلامی با ۳۰٪ سود: ۰ ریال", font=("Tahoma", 11, "bold"), bg="#1e1e1e", fg="#4ec9b0")
        self.final_price_lbl.grid(row=0, column=2, padx=20)

        self.items_data = []

    def add_item(self):
        try:
            name = self.name_entry.get()
            mat = self.mat_combo.get()
            dia = float(self.dia_entry.get()) if self.dia_entry.get() else 0
            length = float(self.len_entry.get()) if self.len_entry.get() else 0
            qty = int(self.qty_entry.get())
            mat_rate = float(self.mat_rate_entry.get())
            turning = float(self.turning_entry.get())
            gear_cut = float(self.gear_cut_entry.get())

            # محاسبه وزن استوانه‌ای (فولاد density ~ 7.85 g/cm3)
            weight = (math.pi * ((dia / 2) ** 2) * length * 7.85 / 1000000) * qty if dia and length else 0
            mat_cost = weight * mat_rate
            machining_cost = (turning + gear_cut) * qty
            total_make = mat_cost + machining_cost

            item = {
                "name": name, "mat": mat, "dia": dia, "len": length, "qty": qty,
                "weight": weight, "mat_cost": mat_cost, "machining_cost": machining_cost,
                "total": total_make
            }
            self.items_data.append(item)

            self.tree.insert("", "end", values=(
                name, mat, dia if dia else "-", length if length else "-", qty,
                f"{weight:.2f}", f"{mat_cost:,.0f}", f"{machining_cost:,.0f}", f"{total_make:,.0f}"
            ))

            self.update_totals()

        except Exception as e:
            messagebox.showerror("خطا", f"لطفاً ورودی‌ها را بررسی کنید:\n{e}")

    def update_totals(self):
        tot_weight = sum(i["weight"] for i in self.items_data)
        tot_make = sum(i["total"] for i in self.items_data)
        final_with_profit = tot_make * 1.30  # ۳۰ درصد سود مطابق اکسل

        self.total_weight_lbl.config(text=f"وزن کل: {tot_weight:.2f} کیلوگرم")
        self.total_cost_lbl.config(text=f"قیمت ساخت کل: {tot_make:,.0f} ریال")
        self.final_price_lbl.config(text=f"قیمت اعلامی با ۳۰٪ سود: {final_with_profit:,.0f} ریال")

    def clear_all(self):
        self.items_data.clear()
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.update_totals()

    def export_excel(self):
        if not self.items_data:
            messagebox.showwarning("هشدار", "هیچ قطعه‌ای در جدول وجود ندارد.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "آنالیز قیمت کاوه کیش"
            ws.views.sheetView[0].rightToLeft = True

            headers = ["ردیف", "نام قطعه", "جنس", "قطر (mm)", "طول (mm)", "تعداد", "وزن (kg)", "قیمت متریال", "اجرت ساخت", "قیمت کل ساخت", "قیمت با سود ۳۰٪"]
            ws.append(headers)

            for idx, item in enumerate(self.items_data, 1):
                ws.append([
                    idx, item["name"], item["mat"], item["dia"], item["len"], item["qty"],
                    round(item["weight"], 2), item["mat_cost"], item["machining_cost"],
                    item["total"], item["total"] * 1.30
                ])

            tot_make = sum(i["total"] for i in self.items_data)
            ws.append([])
            ws.append(["جمع کل به همراه ۳۰٪ سود و بالاسری", "", "", "", "", "", "", "", "", tot_make, tot_make * 1.30])

            wb.save(file_path)
            messagebox.showinfo("موفقیت", "فایل اکسل با موفقیت ذخیره شد.")

if __name__ == "__main__":
    root = tk.Tk()
    app = KaveKishCostApp(root)
    root.mainloop()
