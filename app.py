import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import openpyxl

class TookaTarhCostApp:
    def __init__(self, root):
        self.root = root
        self.root.title("آنالیز متریال دنده و گیربکس جهت شرکت توکا طرح صفاهان")
        self.root.geometry("1300x850")
        self.root.configure(bg="#f4f6f9")

        # Style Configuration
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure(".", background="#f4f6f9", foreground="#2c3e50", fieldbackground="#ffffff")
        style.configure("TLabel", font=("Tahoma", 8, "bold"), background="#f4f6f9", foreground="#2c3e50")
        style.configure("TLabelframe", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", font=("Tahoma", 9, "bold"), foreground="#16a085", background="#ffffff")
        style.configure("TButton", font=("Tahoma", 9, "bold"), background="#2980b9", foreground="#ffffff")
        style.map("TButton", background=[("active", "#3498db")])

        style.configure("TNotebook", background="#f4f6f9", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Tahoma", 10, "bold"), padding=[15, 6], background="#e2e8f0", foreground="#2c3e50")
        style.map("TNotebook.Tab", background=[("selected", "#2980b9")], foreground=[("selected", "#ffffff")])

        style.configure("Treeview", font=("Tahoma", 8), rowheight=24, background="#ffffff", fieldbackground="#ffffff", foreground="#2c3e50")
        style.configure("Treeview.Heading", font=("Tahoma", 8, "bold"), background="#34495e", foreground="#ffffff")

        # Header Title
        header_frame = tk.Frame(root, bg="#1a365d", pady=12)
        header_frame.pack(fill="x")
        header = tk.Label(
            header_frame, 
            text="⚙️ آنالیز متریال دنده و گیربکس جهت شرکت توکا طرح صفاهان", 
            font=("Tahoma", 14, "bold"), 
            bg="#1a365d", 
            fg="#ffffff"
        )
        header.pack()

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Tab Navigation
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)

        self.tab_parts = ttk.Frame(self.notebook, padding="10")
        self.tab_gearbox = ttk.Frame(self.notebook, padding="10")

        self.notebook.add(self.tab_parts, text=" ⚙️ ۱. آنالیز قطعات و متعلقات ")
        self.notebook.add(self.tab_gearbox, text=" 📦 ۲. آنالیز قیمت گیربکس (بر اساس فایل نمونه) ")

        self.build_parts_tab()
        self.build_gearbox_tab()

    # =========================================================
    # TAB 1: آنالیز قطعات
    # =========================================================
    def build_parts_tab(self):
        inputs_frame = ttk.LabelFrame(self.tab_parts, text=" مشخصات و هزینه‌های قطعه ", padding="10")
        inputs_frame.pack(fill="x", pady=5)

        ttk.Label(inputs_frame, text="مشخصات/نام قطعه:").grid(row=0, column=0, sticky="w", pady=3)
        self.p_name_entry = ttk.Entry(inputs_frame, width=18)
        self.p_name_entry.insert(0, "GEAR")
        self.p_name_entry.grid(row=0, column=1, padx=5, pady=3)

        ttk.Label(inputs_frame, text="جنس متریال:").grid(row=0, column=2, sticky="w", pady=3)
        self.p_mat_combo = ttk.Combobox(inputs_frame, values=["CK45", "ST52", "1.6582", "CuSn12Ni2", "St52-3", "MO40", "VCN150"], width=14)
        self.p_mat_combo.set("CK45")
        self.p_mat_combo.grid(row=0, column=3, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قطر (mm):").grid(row=0, column=4, sticky="w", pady=3)
        self.p_dia_entry = ttk.Entry(inputs_frame, width=8)
        self.p_dia_entry.insert(0, "0")
        self.p_dia_entry.grid(row=0, column=5, padx=5, pady=3)

        ttk.Label(inputs_frame, text="طول (mm):").grid(row=0, column=6, sticky="w", pady=3)
        self.p_len_entry = ttk.Entry(inputs_frame, width=8)
        self.p_len_entry.insert(0, "0")
        self.p_len_entry.grid(row=0, column=7, padx=5, pady=3)

        ttk.Label(inputs_frame, text="وزن دستی (kg):").grid(row=0, column=8, sticky="w", pady=3)
        self.p_weight_entry = ttk.Entry(inputs_frame, width=10)
        self.p_weight_entry.insert(0, "50.0")
        self.p_weight_entry.grid(row=0, column=9, padx=5, pady=3)

        ttk.Label(inputs_frame, text="تعداد:").grid(row=0, column=10, sticky="w", pady=3)
        self.p_qty_entry = ttk.Entry(inputs_frame, width=6)
        self.p_qty_entry.insert(0, "1")
        self.p_qty_entry.grid(row=0, column=11, padx=5, pady=3)

        # Costs
        ttk.Label(inputs_frame, text="نرخ متریال (ریال/کیلو):").grid(row=1, column=0, sticky="w", pady=3)
        self.p_mat_rate_entry = ttk.Entry(inputs_frame, width=18)
        self.p_mat_rate_entry.insert(0, "3000000")
        self.p_mat_rate_entry.grid(row=1, column=1, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت مدلسازی:").grid(row=1, column=2, sticky="w", pady=3)
        self.p_model_entry = ttk.Entry(inputs_frame, width=14)
        self.p_model_entry.insert(0, "0")
        self.p_model_entry.grid(row=1, column=3, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت تراش:").grid(row=1, column=4, sticky="w", pady=3)
        self.p_turn_entry = ttk.Entry(inputs_frame, width=8)
        self.p_turn_entry.insert(0, "5000000")
        self.p_turn_entry.grid(row=1, column=5, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت دنده:").grid(row=1, column=6, sticky="w", pady=3)
        self.p_gear_entry = ttk.Entry(inputs_frame, width=8)
        self.p_gear_entry.insert(0, "3000000")
        self.p_gear_entry.grid(row=1, column=7, padx=5, pady=3)

        ttk.Label(inputs_frame, text="قیمت سنگ:").grid(row=1, column=8, sticky="w", pady=3)
        self.p_grind_entry = ttk.Entry(inputs_frame, width=10)
        self.p_grind_entry.insert(0, "0")
        self.p_grind_entry.grid(row=1, column=9, padx=5, pady=3)

        ttk.Label(inputs_frame, text="عملیات حرارتی:").grid(row=2, column=0, sticky="w", pady=3)
        self.p_heat_entry = ttk.Entry(inputs_frame, width=18)
        self.p_heat_entry.insert(0, "0")
        self.p_heat_entry.grid(row=2, column=1, padx=5, pady=3)

        ttk.Label(inputs_frame, text="تست/بازرسی NDT:").grid(row=2, column=2, sticky="w", pady=3)
        self.p_ndt_entry = ttk.Entry(inputs_frame, width=14)
        self.p_ndt_entry.insert(0, "500000")
        self.p_ndt_entry.grid(row=2, column=3, padx=5, pady=3)

        ttk.Label(inputs_frame, text="بسته‌بندی:").grid(row=2, column=4, sticky="w", pady=3)
        self.p_pack_entry = ttk.Entry(inputs_frame, width=8)
        self.p_pack_entry.insert(0, "200000")
        self.p_pack_entry.grid(row=2, column=5, padx=5, pady=3)

        ttk.Label(inputs_frame, text="هزینه ارسال:").grid(row=2, column=6, sticky="w", pady=3)
        self.p_ship_entry = ttk.Entry(inputs_frame, width=8)
        self.p_ship_entry.insert(0, "300000")
        self.p_ship_entry.grid(row=2, column=7, padx=5, pady=3)

        btn_frame = ttk.Frame(self.tab_parts)
        btn_frame.pack(fill="x", pady=5)

        ttk.Button(btn_frame, text="➕ افزودن قطعه", command=self.add_part_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="❌ حذف سطر", command=self.delete_part_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="📥 خروجی اکسل قطعات (.xlsx)", command=self.export_parts_excel).pack(side="right", padx=5)

        cols = ("name", "mat", "dia", "len", "qty", "weight", "mat_cost", "turn_cost", "gear_cost", "heat_cost", "total_make", "final_price")
        self.p_tree = ttk.Treeview(self.tab_parts, columns=cols, show="headings", height=12)

        headings = {
            "name": "نام قطعه", "mat": "جنس", "dia": "قطر", "len": "طول", "qty": "تعداد",
            "weight": "وزن (kg)", "mat_cost": "قیمت متریال", "turn_cost": "تراش", "gear_cost": "دنده",
            "heat_cost": "عملیات حرارتی", "total_make": "قیمت ساخت کل", "final_price": "اعلامی (+۳۰٪)"
        }
        for c in cols:
            self.p_tree.heading(c, text=headings[c])
            self.p_tree.column(c, width=95, anchor="center")

        self.p_tree.column("name", width=120, anchor="w")
        self.p_tree.pack(fill="both", expand=True, pady=5)

        sum_frame = ttk.LabelFrame(self.tab_parts, text=" جمع‌بندی قطعات ", padding="10")
        sum_frame.pack(fill="x", pady=5)

        self.p_total_weight_lbl = tk.Label(sum_frame, text="وزن کل: ۰ کیلوگرم", font=("Tahoma", 10, "bold"), bg="#ffffff", fg="#2c3e50")
        self.p_total_weight_lbl.grid(row=0, column=0, padx=30)

        self.p_total_cost_lbl = tk.Label(sum_frame, text="ساخت کل: ۰ ریال", font=("Tahoma", 10, "bold"), bg="#ffffff", fg="#2c3e50")
        self.p_total_cost_lbl.grid(row=0, column=1, padx=30)

        self.p_final_price_lbl = tk.Label(sum_frame, text="قیمت کل اعلامی (+۳۰٪ سود): ۰ ریال", font=("Tahoma", 11, "bold"), bg="#ffffff", fg="#16a085")
        self.p_final_price_lbl.grid(row=0, column=2, padx=30)

        self.parts_data = []

    def add_part_item(self):
        try:
            name = self.p_name_entry.get()
            mat = self.p_mat_combo.get()
            dia = float(self.p_dia_entry.get()) if self.p_dia_entry.get() else 0
            length = float(self.p_len_entry.get()) if self.p_len_entry.get() else 0
            manual_weight = float(self.p_weight_entry.get()) if self.p_weight_entry.get() else 0
            qty = int(self.p_qty_entry.get())

            if dia > 0 and length > 0:
                weight = (math.pi * ((dia / 2) ** 2) * length * 7.85 / 1000000) * qty
            else:
                weight = manual_weight * qty

            mat_cost = weight * float(self.p_mat_rate_entry.get())
            model_cost = float(self.p_model_entry.get()) * qty
            turn_cost = float(self.p_turn_entry.get()) * qty
            gear_cost = float(self.p_gear_entry.get()) * qty
            grind_cost = float(self.p_grind_entry.get()) * qty
            heat_cost = float(self.p_heat_entry.get()) * qty
            ndt_cost = float(self.p_ndt_entry.get()) * qty
            pack_cost = float(self.p_pack_entry.get()) * qty
            ship_cost = float(self.p_ship_entry.get()) * qty

            total_make = mat_cost + model_cost + turn_cost + gear_cost + grind_cost + heat_cost + ndt_cost + pack_cost + ship_cost
            final_price = total_make * 1.30

            item = {
                "name": name, "mat": mat, "dia": dia, "len": length, "qty": qty,
                "weight": weight, "mat_cost": mat_cost, "turn_cost": turn_cost,
                "gear_cost": gear_cost, "heat_cost": heat_cost, "total_make": total_make,
                "final_price": final_price
            }
            self.parts_data.append(item)

            self.p_tree.insert("", "end", values=(
                name, mat, dia if dia else "-", length if length else "-", qty,
                f"{weight:.2f}", f"{mat_cost:,.0f}", f"{turn_cost:,.0f}", f"{gear_cost:,.0f}",
                f"{heat_cost:,.0f}", f"{total_make:,.0f}", f"{final_price:,.0f}"
            ))

            tot_w = sum(i["weight"] for i in self.parts_data)
            tot_m = sum(i["total_make"] for i in self.parts_data)
            tot_f = sum(i["final_price"] for i in self.parts_data)

            self.p_total_weight_lbl.config(text=f"وزن کل: {tot_w:.2f} کیلوگرم")
            self.p_total_cost_lbl.config(text=f"ساخت کل: {tot_m:,.0f} ریال")
            self.p_final_price_lbl.config(text=f"قیمت کل اعلامی (+۳۰٪ سود): {tot_f:,.0f} ریال")

        except Exception as e:
            messagebox.showerror("خطا", f"لطفاً ورودی‌ها را بررسی نمایید:\n{e}")

    def delete_part_item(self):
        selected = self.p_tree.selection()
        if not selected:
            return
        for item in selected:
            idx = self.p_tree.index(item)
            self.p_tree.delete(item)
            del self.parts_data[idx]

    def export_parts_excel(self):
        if not self.parts_data:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "آنالیز قطعات"
            ws.views.sheetView[0].rightToLeft = True

            headers = ["ردیف", "نام قطعه", "جنس", "قطر", "طول", "تعداد", "وزن (kg)", "قیمت متریال", "تراش", "دنده", "عملیات حرارتی", "ساخت کل", "اعلامی (+۳۰٪)"]
            ws.append(headers)

            for idx, item in enumerate(self.parts_data, 1):
                ws.append([
                    idx, item["name"], item["mat"], item["dia"], item["len"], item["qty"],
                    round(item["weight"], 2), item["mat_cost"], item["turn_cost"], item["gear_cost"],
                    item["heat_cost"], item["total_make"], item["final_price"]
                ])

            wb.save(file_path)
            messagebox.showinfo("موفقیت", "فایل اکسل قطعات ذخیره شد.")

    # =========================================================
    # TAB 2: آنالیز گیربکس (بر اساس دقیقاً فرمول‌های فایل نمومه.xlsx)
    # =========================================================
    def build_gearbox_tab(self):
        gb_frame = ttk.LabelFrame(self.tab_gearbox, text=" ورودی‌های قیمت گیربکس (مطابق فایل نمونه) ", padding="15")
        gb_frame.pack(fill="x", pady=5)

        ttk.Label(gb_frame, text="وزن کل (kg):").grid(row=0, column=0, sticky="w", pady=6)
        self.gb_weight_entry = ttk.Entry(gb_frame, width=15)
        self.gb_weight_entry.insert(0, "250")
        self.gb_weight_entry.grid(row=0, column=1, padx=8, pady=6)

        ttk.Label(gb_frame, text="قیمت متریال هر کیلوگرم (ریال):").grid(row=0, column=2, sticky="w", pady=6)
        self.gb_mat_per_kg_entry = ttk.Entry(gb_frame, width=20)
        self.gb_mat_per_kg_entry.insert(0, "7500000")
        self.gb_mat_per_kg_entry.grid(row=0, column=3, padx=8, pady=6)

        ttk.Label(gb_frame, text="قیمت ساخت (ریال):").grid(row=1, column=0, sticky="w", pady=6)
        self.gb_make_cost_entry = ttk.Entry(gb_frame, width=15)
        self.gb_make_cost_entry.insert(0, "1000000000")
        self.gb_make_cost_entry.grid(row=1, column=1, padx=8, pady=6)

        ttk.Label(gb_frame, text="قیمت استاندارد (ریال):").grid(row=1, column=2, sticky="w", pady=6)
        self.gb_std_cost_entry = ttk.Entry(gb_frame, width=20)
        self.gb_std_cost_entry.insert(0, "500000000")
        self.gb_std_cost_entry.grid(row=1, column=3, padx=8, pady=6)

        ttk.Label(gb_frame, text="قیمت مونتاژ (ریال):").grid(row=2, column=0, sticky="w", pady=6)
        self.gb_assy_cost_entry = ttk.Entry(gb_frame, width=15)
        self.gb_assy_cost_entry.insert(0, "500000000")
        self.gb_assy_cost_entry.grid(row=2, column=1, padx=8, pady=6)

        ttk.Label(gb_frame, text="قیمت بسته‌بندی (ریال):").grid(row=2, column=2, sticky="w", pady=6)
        self.gb_pack_cost_entry = ttk.Entry(gb_frame, width=20)
        self.gb_pack_cost_entry.insert(0, "250000000")
        self.gb_pack_cost_entry.grid(row=2, column=3, padx=8, pady=6)

        ttk.Label(gb_frame, text="قیمت ارسال (ریال):").grid(row=3, column=0, sticky="w", pady=6)
        self.gb_ship_cost_entry = ttk.Entry(gb_frame, width=15)
        self.gb_ship_cost_entry.insert(0, "375000000")
        self.gb_ship_cost_entry.grid(row=3, column=1, padx=8, pady=6)

        btn_calc_gb = ttk.Button(gb_frame, text="🧮 محاسبه آنالیز قیمت گیربکس", command=self.calculate_sample_gearbox)
        btn_calc_gb.grid(row=3, column=2, columnspan=2, sticky="ew", padx=8, pady=6)

        # Output Summary
        res_frame = ttk.LabelFrame(self.tab_gearbox, text=" خروجی محاسبات (دقیقاً مطابق اکسل نمونه) ", padding="15")
        res_frame.pack(fill="x", pady=10)

        self.gb_res_mat_total = tk.Label(res_frame, text="قیمت متریال کل: ۰ ریال", font=("Tahoma", 10, "bold"), bg="#ffffff", fg="#2c3e50")
        self.gb_res_mat_total.grid(row=0, column=0, padx=20, pady=6, sticky="w")

        self.gb_res_total_cost = tk.Label(res_frame, text="قیمت کل (بدون سود): ۰ ریال", font=("Tahoma", 10, "bold"), bg="#ffffff", fg="#2c3e50")
        self.gb_res_total_cost.grid(row=0, column=1, padx=20, pady=6, sticky="w")

        self.gb_res_final_profit = tk.Label(res_frame, text="قیمت کل + سود (۳۰٪): ۰ ریال", font=("Tahoma", 11, "bold"), bg="#ffffff", fg="#16a085")
        self.gb_res_final_profit.grid(row=1, column=0, padx=20, pady=6, sticky="w")

        self.gb_res_per_kg = tk.Label(res_frame, text="قیمت هر کیلوگرم: ۰ ریال/کیلوگرم", font=("Tahoma", 11, "bold"), bg="#ffffff", fg="#2980b9")
        self.gb_res_per_kg.grid(row=1, column=1, padx=20, pady=6, sticky="w")

        ttk.Button(self.tab_gearbox, text="📥 دریافت خروجی اکسل گیربکس (.xlsx)", command=self.export_sample_gb_excel).pack(anchor="e", pady=5)

    def calculate_sample_gearbox(self):
        try:
            w = float(self.gb_weight_entry.get())
            mat_per_kg = float(self.gb_mat_per_kg_entry.get())
            
            # متریال کل = وزن * قیمت متریال هر کیلوگرم
            mat_total = w * mat_per_kg
            
            make_cost = float(self.gb_make_cost_entry.get())
            std_cost = float(self.gb_std_cost_entry.get())
            assy_cost = float(self.gb_assy_cost_entry.get())
            pack_cost = float(self.gb_pack_cost_entry.get())
            ship_cost = float(self.gb_ship_cost_entry.get())

            # قیمت کل = متریال کل + بقیه هزینه‌ها
            total_cost = mat_total + make_cost + std_cost + assy_cost + pack_cost + ship_cost
            
            # قیمت کل + سود (۳۰٪)
            total_with_profit = total_cost * 1.30
            
            # قیمت هر کیلوگرم = (قیمت کل + سود) / وزن کل
            price_per_kg = total_with_profit / w if w > 0 else 0

            self.gb_res_mat_total.config(text=f"قیمت متریال کل: {mat_total:,.0f} ریال")
            self.gb_res_total_cost.config(text=f"قیمت کل (بدون سود): {total_cost:,.0f} ریال")
            self.gb_res_final_profit.config(text=f"قیمت کل + سود (۳۰٪): {total_with_profit:,.0f} ریال")
            self.gb_res_per_kg.config(text=f"قیمت هر کیلوگرم: {price_per_kg:,.0f} ریال/کیلوگرم")

            self.last_gb_data = {
                "w": w, "mat_per_kg": mat_per_kg, "mat_total": mat_total,
                "make_cost": make_cost, "std_cost": std_cost, "assy_cost": assy_cost,
                "pack_cost": pack_cost, "ship_cost": ship_cost, "total_cost": total_cost,
                "total_with_profit": total_with_profit, "price_per_kg": price_per_kg
            }

        except Exception as e:
            messagebox.showerror("خطا", f"لطفاً ورودی‌های گیربکس را بررسی کنید:\n{e}")

    def export_sample_gb_excel(self):
        if not hasattr(self, 'last_gb_data'):
            messagebox.showwarning("هشدار", "ابتدا دکمه محاسبه را بزنید.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "آنالیز گیربکس"
            ws.views.sheetView[0].rightToLeft = True

            headers = [
                "وزن کل (kg)", "قیمت متریال هر کیلوگرم", "قیمت متریال کل", "قیمت ساخت", 
                "قیمت استاندارد", "قیمت مونتاژ", "قیمت بسته‌بندی", "قیمت ارسال", 
                "قیمت کل", "قیمت کل + سود (۳۰٪)", "قیمت هر کیلوگرم"
            ]
            ws.append(headers)

            d = self.last_gb_data
            ws.append([
                d["w"], d["mat_per_kg"], d["mat_total"], d["make_cost"],
                d["std_cost"], d["assy_cost"], d["pack_cost"], d["ship_cost"],
                d["total_cost"], d["total_with_profit"], d["price_per_kg"]
            ])

            wb.save(file_path)
            messagebox.showinfo("موفقیت", "فایل اکسل آنالیز گیربکس ذخیره گردید.")

if __name__ == "__main__":
    root = tk.Tk()
    app = TookaTarhCostApp(root)
    root.mainloop()
