import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import openpyxl

class AdvancedGearboxCostApp:
    def __init__(self, root):
        self.root = root
        self.root.title("سیستم پیشرفته آنالیز متریال، خدمات و برآورد قیمت گیربکس و قطعات")
        self.root.geometry("1280x800")
        self.root.configure(bg="#1e1e1e")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background="#1e1e1e", foreground="#ffffff", fieldbackground="#2d2d2d")
        style.configure("TLabel", font=("Tahoma", 8), background="#1e1e1e", foreground="#ffffff")
        style.configure("TButton", font=("Tahoma", 9, "bold"), background="#0e639c", foreground="#ffffff")
        style.configure("Treeview", font=("Tahoma", 8), rowheight=22)
        style.configure("Treeview.Heading", font=("Tahoma", 8, "bold"), background="#252526", foreground="#007acc")

        # Header
        header = tk.Label(root, text="⚙️ آنالیز جامع قیمت قطعات و گیربکس (امکان ورود مستقیم وزن و کلیه هزینه‌ها)", font=("Tahoma", 13, "bold"), bg="#252526", fg="#007acc", pady=8)
        header.pack(fill="x")

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Inputs Section
        inputs_frame = ttk.LabelFrame(main_frame, text=" ورودی‌های مشخصات قطعه و هزینه‌ها ", padding="10")
        inputs_frame.pack(fill="x", pady=5)

        # Row 0: Basic Specs
        ttk.Label(inputs_frame, text="نام/مشخصات قطعه:").grid(row=0, column=0, sticky="w", pady=3)
        self.name_entry = ttk.Entry(inputs_frame, width=18)
        self.name_entry.insert(0, "GEAR")
        self.name_entry.grid(row=0, column=1, padx=5, pady=3)

        ttk.Label(inputs_frame, text="جنس/متریال (قابل تایپ):").grid(row=0, column=2, sticky="w", pady=3)
        self.mat_combo = ttk.Combobox(inputs_frame, values=["CK45", "ST52", "1.6582", "CuSn12Ni2", "St52-3", "MO40", "VCN150"], width=15)
        self.mat_combo.set("CK45")
        self.mat_combo.grid(row=0, column=3, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قطر (mm):").grid(row=0, column=4, sticky="w", pady=3)
        self.dia_entry = ttk.Entry(inputs_frame, width=10)
        self.dia_entry.insert(0, "0")
        self.dia_entry.grid(row=0, column=5, padx=5, pady=3)

        ttk.Label(inputs_frame, text="طول (mm):").grid(row=0, column=6, sticky="w", pady=3)
        self.len_entry = ttk.Entry(inputs_frame, width=10)
        self.len_entry.insert(0, "0")
        self.len_entry.grid(row=0, column=7, padx=5, pady=3)

        ttk.Label(inputs_frame, text="وزن دستی (kg):").grid(row=0, column=8, sticky="w", pady=3)
        self.manual_weight_entry = ttk.Entry(inputs_frame, width=10)
        self.manual_weight_entry.insert(0, "50.0")
        self.manual_weight_entry.grid(row=0, column=9, padx=5, pady=3)

        ttk.Label(inputs_frame, text="تعداد:").grid(row=0, column=10, sticky="w", pady=3)
        self.qty_entry = ttk.Entry(inputs_frame, width=8)
        self.qty_entry.insert(0, "1")
        self.qty_entry.grid(row=0, column=11, padx=5, pady=3)

        # Row 1: Costs (Part 1)
        ttk.Label(inputs_frame, text="قیمت متریال (ریال/کیلو):").grid(row=1, column=0, sticky="w", pady=3)
        self.mat_rate_entry = ttk.Entry(inputs_frame, width=18)
        self.mat_rate_entry.insert(0, "3000000")
        self.mat_rate_entry.grid(row=1, column=1, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت مدلسازی:").grid(row=1, column=2, sticky="w", pady=3)
        self.model_entry = ttk.Entry(inputs_frame, width=15)
        self.model_entry.insert(0, "0")
        self.model_entry.grid(row=1, column=3, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت تراشکاری:").grid(row=1, column=4, sticky="w", pady=3)
        self.turning_entry = ttk.Entry(inputs_frame, width=10)
        self.turning_entry.insert(0, "5000000")
        self.turning_entry.grid(row=1, column=5, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت دنده‌زنی:").grid(row=1, column=6, sticky="w", pady=3)
        self.gear_cut_entry = ttk.Entry(inputs_frame, width=10)
        self.gear_cut_entry.insert(0, "3000000")
        self.gear_cut_entry.grid(row=1, column=7, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت سنگ‌زنی:").grid(row=1, column=8, sticky="w", pady=3)
        self.grind_entry = ttk.Entry(inputs_frame, width=10)
        self.grind_entry.insert(0, "0")
        self.grind_entry.grid(row=1, column=9, padx=5, pady=3)

        # Row 2: Costs (Part 2)
        ttk.Label(inputs_frame, text="عملیات حرارتی:").grid(row=2, column=0, sticky="w", pady=3)
        self.heat_entry = ttk.Entry(inputs_frame, width=18)
        self.heat_entry.insert(0, "0")
        self.heat_entry.grid(row=2, column=1, padx=5, pady=3)

        ttk.Label(inputs_frame, text="تست NDT / بازرسی:").grid(row=2, column=2, sticky="w", pady=3)
        self.ndt_entry = ttk.Entry(inputs_frame, width=15)
        self.ndt_entry.insert(0, "500000")
        self.ndt_entry.grid(row=2, column=3, padx=5, pady=3)

        ttk.Label(inputs_frame, text="بسته‌بندی و بارگیری:").grid(row=2, column=4, sticky="w", pady=3)
        self.pack_entry = ttk.Entry(inputs_frame, width=10)
        self.pack_entry.insert(0, "200000")
        self.pack_entry.grid(row=2, column=5, padx=5, pady=3)

        ttk.Label(inputs_frame, text="هزینه حمل و ارسال:").grid(row=2, column=6, sticky="w", pady=3)
        self.ship_entry = ttk.Entry(inputs_frame, width=10)
        self.ship_entry.insert(0, "300000")
        self.ship_entry.grid(row=2, column=7, padx=5, pady=3)

        # Action Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=5)

        add_btn = ttk.Button(btn_frame, text="➕ محاسبه و افزودن به جدول", command=self.add_item)
        add_btn.pack(side="left", padx=5)

        del_btn = ttk.Button(btn_frame, text="❌ حذف سطر انتخاب‌شده", command=self.delete_selected)
        del_btn.pack(side="left", padx=5)

        clear_btn = ttk.Button(btn_frame, text="🗑️ پاک کردن همه", command=self.clear_all)
        clear_btn.pack(side="left", padx=5)

        excel_btn = ttk.Button(btn_frame, text="📥 دریافت خروجی اکسل کامل (.xlsx)", command=self.export_excel)
        excel_btn.pack(side="right", padx=5)

        # Extended Table Display
        cols = (
            "name", "mat", "dia", "len", "qty", "weight", "mat_cost", 
            "model_cost", "turn_cost", "gear_cost", "grind_cost", 
            "heat_cost", "ndt_cost", "pack_cost", "ship_cost", "total_make", "final_price"
        )
        self.tree = ttk.Treeview(main_frame, columns=cols, show="headings", height=12)

        headings = {
            "name": "نام قطعه", "mat": "جنس", "dia": "قطر", "len": "طول", "qty": "تعداد",
            "weight": "وزن (kg)", "mat_cost": "متریال", "model_cost": "مدلسازی",
            "turn_cost": "تراش", "gear_cost": "دنده", "grind_cost": "سنگ",
            "heat_cost": "عملیات حرارتی", "ndt_cost": "تست/بازرسی",
            "pack_cost": "بسته‌بندی", "ship_cost": "ارسال", "total_make": "ساخت کل", "final_price": "اعلامی (+۳۰٪)"
        }

        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=75, anchor="center")

        self.tree.column("name", width=110, anchor="w")
        self.tree.pack(fill="both", expand=True, pady=5)

        # Summary Frame
        summary_frame = ttk.LabelFrame(main_frame, text=" جمع‌بندی نهایی پروژه / گیربکس ", padding="10")
        summary_frame.pack(fill="x", pady=5)

        self.total_weight_lbl = tk.Label(summary_frame, text="مجموع وزن: ۰ کیلوگرم", font=("Tahoma", 10, "bold"), bg="#1e1e1e", fg="#ffffff")
        self.total_weight_lbl.grid(row=0, column=0, padx=25)

        self.total_cost_lbl = tk.Label(summary_frame, text="مجموع قیمت ساخت: ۰ ریال", font=("Tahoma", 10, "bold"), bg="#1e1e1e", fg="#ffffff")
        self.total_cost_lbl.grid(row=0, column=1, padx=25)

        self.final_price_lbl = tk.Label(summary_frame, text="قیمت کل با ۳۰٪ سود و بالاسری: ۰ ریال", font=("Tahoma", 11, "bold"), bg="#1e1e1e", fg="#4ec9b0")
        self.final_price_lbl.grid(row=0, column=2, padx=25)

        self.items_data = []

    def add_item(self):
        try:
            name = self.name_entry.get()
            mat = self.mat_combo.get()
            dia = float(self.dia_entry.get()) if self.dia_entry.get() else 0
            length = float(self.len_entry.get()) if self.len_entry.get() else 0
            manual_weight = float(self.manual_weight_entry.get()) if self.manual_weight_entry.get() else 0
            qty = int(self.qty_entry.get())

            # اولویت وزن: اگر قطر و طول صفر باشند از وزن دستی استفاده می‌شود
            if dia > 0 and length > 0:
                weight = (math.pi * ((dia / 2) ** 2) * length * 7.85 / 1000000) * qty
            else:
                weight = manual_weight * qty

            mat_rate = float(self.mat_rate_entry.get())
            mat_cost = weight * mat_rate

            model_cost = float(self.model_entry.get()) * qty
            turn_cost = float(self.turning_entry.get()) * qty
            gear_cost = float(self.gear_cut_entry.get()) * qty
            grind_cost = float(self.grind_entry.get()) * qty
            heat_cost = float(self.heat_entry.get()) * qty
            ndt_cost = float(self.ndt_entry.get()) * qty
            pack_cost = float(self.pack_entry.get()) * qty
            ship_cost = float(self.ship_entry.get()) * qty

            total_make = mat_cost + model_cost + turn_cost + gear_cost + grind_cost + heat_cost + ndt_cost + pack_cost + ship_cost
            final_price = total_make * 1.30  # ۳0 درصد سود و بالاسری

            item = {
                "name": name, "mat": mat, "dia": dia, "len": length, "qty": qty,
                "weight": weight, "mat_cost": mat_cost, "model_cost": model_cost,
                "turn_cost": turn_cost, "gear_cost": gear_cost, "grind_cost": grind_cost,
                "heat_cost": heat_cost, "ndt_cost": ndt_cost, "pack_cost": pack_cost,
                "ship_cost": ship_cost, "total_make": total_make, "final_price": final_price
            }
            self.items_data.append(item)

            self.tree.insert("", "end", values=(
                name, mat, dia if dia else "-", length if length else "-", qty,
                f"{weight:.2f}", f"{mat_cost:,.0f}", f"{model_cost:,.0f}",
                f"{turn_cost:,.0f}", f"{gear_cost:,.0f}", f"{grind_cost:,.0f}",
                f"{heat_cost:,.0f}", f"{ndt_cost:,.0f}", f"{pack_cost:,.0f}",
                f"{ship_cost:,.0f}", f"{total_make:,.0f}", f"{final_price:,.0f}"
            ))

            self.update_totals()

        except Exception as e:
            messagebox.showerror("خطا", f"لطفاً مقادیر ورودی را بررسی نمایید:\n{e}")

    def delete_selected(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("هشدار", "لطفاً سطر مورد نظر را انتخاب کنید.")
            return
        for item in selected_item:
            idx = self.tree.index(item)
            self.tree.delete(item)
            del self.items_data[idx]
        self.update_totals()

    def update_totals(self):
        tot_weight = sum(i["weight"] for i in self.items_data)
        tot_make = sum(i["total_make"] for i in self.items_data)
        tot_final = sum(i["final_price"] for i in self.items_data)

        self.total_weight_lbl.config(text=f"مجموع وزن: {tot_weight:.2f} کیلوگرم")
        self.total_cost_lbl.config(text=f"مجموع قیمت ساخت: {tot_make:,.0f} ریال")
        self.final_price_lbl.config(text=f"قیمت کل با ۳۰٪ سود و بالاسری: {tot_final:,.0f} ریال")

    def clear_all(self):
        self.items_data.clear()
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.update_totals()

    def export_excel(self):
        if not self.items_data:
            messagebox.showwarning("هشدار", "هیچ داده‌ای برای خروجی وجود ندارد.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "آنالیز قیمت کاوه کیش"
            ws.views.sheetView[0].rightToLeft = True

            headers = [
                "ردیف", "مشخصات و نوع کالا/خدمات", "جنس", "قطر", "طول", "تعداد", "وزن (kg)",
                "قیمت متریال", "قیمت مدلسازی", "قیمت تراش", "قیمت دنده", "قیمت سنگ",
                "عملیات حرارتی", "تست/بازرسی", "بسته‌بندی", "ارسال", "قیمت ساخت کل", "قیمت اعلامی با ۳۰٪ سود"
            ]
            ws.append(headers)

            for idx, item in enumerate(self.items_data, 1):
                ws.append([
                    idx, item["name"], item["mat"], item["dia"], item["len"], item["qty"],
                    round(item["weight"], 2), item["mat_cost"], item["model_cost"],
                    item["turn_cost"], item["gear_cost"], item["grind_cost"],
                    item["heat_cost"], item["ndt_cost"], item["pack_cost"], item["ship_cost"],
                    item["total_make"], item["final_price"]
                ])

            tot_weight = sum(i["weight"] for i in self.items_data)
            tot_make = sum(i["total_make"] for i in self.items_data)
            tot_final = sum(i["final_price"] for i in self.items_data)

            ws.append([])
            ws.append(["جمع کل پروژه", "", "", "", "", "", round(tot_weight, 2), "", "", "", "", "", "", "", "", "", tot_make, tot_final])

            wb.save(file_path)
            messagebox.showinfo("موفقیت", "فایل اکسل جامع با موفقیت ذخیره گردید.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedGearboxCostApp(root)
    root.mainloop()
