import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import openpyxl

class TookaTarhCostApp:
    def __init__(self, root):
        self.root = root
        self.root.title("آنالیز متریال دنده و گیربکس جهت شرکت توکا طرح سپاهان")
        self.root.geometry("1300x850")
        self.root.configure(bg="#f4f6f9")

        # Style Configuration (Light Professional Theme)
        style = ttk.Style()
        style.theme_use("clam")
        
        # General Colors
        style.configure(".", background="#f4f6f9", foreground="#2c3e50", fieldbackground="#ffffff")
        style.configure("TLabel", font=("Tahoma", 8, "bold"), background="#f4f6f9", foreground="#2c3e50")
        style.configure("TLabelframe", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", font=("Tahoma", 9, "bold"), foreground="#16a085", background="#ffffff")
        style.configure("TButton", font=("Tahoma", 9, "bold"), background="#2980b9", foreground="#ffffff")
        style.map("TButton", background=[("active", "#3498db")])

        # Notebook / Tab Style
        style.configure("TNotebook", background="#f4f6f9", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Tahoma", 10, "bold"), padding=[15, 6], background="#e2e8f0", foreground="#2c3e50")
        style.map("TNotebook.Tab", background=[("selected", "#2980b9")], foreground=[("selected", "#ffffff")])

        # Treeview Style
        style.configure("Treeview", font=("Tahoma", 8), rowheight=24, background="#ffffff", fieldbackground="#ffffff", foreground="#2c3e50")
        style.configure("Treeview.Heading", font=("Tahoma", 8, "bold"), background="#34495e", foreground="#ffffff")

        # Header Title
        header_frame = tk.Frame(root, bg="#1a365d", py=12)
        header_frame.pack(fill="x")
        header = tk.Label(
            header_frame, 
            text="⚙️ آنالیز متریال دنده و گیربکس جهت شرکت توکا طرح سپاهان", 
            font=("Tahoma", 14, "bold"), 
            bg="#1a365d", 
            fg="#ffffff"
        )
        header.pack()

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Tab Navigation (قسمت تفکیک قطعات و گیربکس)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)

        # Tabs
        self.tab_parts = ttk.Frame(self.notebook, padding="10")
        self.tab_gearbox = ttk.Frame(self.notebook, padding="10")

        self.notebook.add(self.tab_parts, text=" ⚙️ ۱. آنالیز قطعات و متعلقات (Component Level) ")
        self.notebook.add(self.tab_gearbox, text=" 📦 ۲. آنالیز گیربکس کامل (Gearbox Assembly) ")

        # Build Content for Both Tabs
        self.build_parts_tab()
        self.build_gearbox_tab()

    # =========================================================
    # TAB 1: آنالیز قطعات
    # =========================================================
    def build_parts_tab(self):
        inputs_frame = ttk.LabelFrame(self.tab_parts, text=" مشخصات و هزینه‌های قطعه ", padding="10")
        inputs_frame.pack(fill="x", pady=5)

        # Row 0
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

        # Row 1 Costs
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

        # Row 2 Costs
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

        # Action Buttons
        btn_frame = ttk.Frame(self.tab_parts)
        btn_frame.pack(fill="x", pady=5)

        ttk.Button(btn_frame, text="➕ افزودن قطعه", command=self.add_part_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="❌ حذف سطر", command=self.delete_part_item).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="📥 دریافت خروجی اکسل قطعات (.xlsx)", command=self.export_parts_excel).pack(side="right", padx=5)

        # Treeview Parts
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

        # Summary Parts
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
    # TAB 2: آنالیز گیربکس کامل (مطابق جدول گیربکس اکسل کاوه کیش)
    # =========================================================
    def build_gearbox_tab(self):
        gb_frame = ttk.LabelFrame(self.tab_gearbox, text=" ورودی‌های آنالیز تجمیعی گیربکس ", padding="10")
        gb_frame.pack(fill="x", pady=5)

        # Grid inputs
        ttk.Label(gb_frame, text="عنوان گیربکس / پروژه:").grid(row=0, column=0, sticky="w", pady=4)
        self.gb_title_entry = ttk.Entry(gb_frame, width=20)
        self.gb_title_entry.insert(0, "گیربکس کاوه کیش")
        self.gb_title_entry.grid(row=0, column=1, padx=5, pady=4)

        ttk.Label(gb_frame, text="تعداد گیربکس:").grid(row=0, column=2, sticky="w", pady=4)
        self.gb_qty_entry = ttk.Entry(gb_frame, width=10)
        self.gb_qty_entry.insert(0, "18")
        self.gb_qty_entry.grid(row=0, column=3, padx=5, pady=4)

        ttk.Label(gb_frame, text="وزن هر گیربکس (kg):").grid(row=0, column=4, sticky="w", pady=4)
        self.gb_unit_w_entry = ttk.Entry(gb_frame, width=10)
        self.gb_unit_w_entry.insert(0, "250")
        self.gb_unit_w_entry.grid(row=0, column=5, padx=5, pady=4)

        # Cost Breakdown
        ttk.Label(gb_frame, text="قیمت متریال کل:").grid(row=1, column=0, sticky="w", pady=4)
        self.gb_mat_cost_entry = ttk.Entry(gb_frame, width=20)
        self.gb_mat_cost_entry.insert(0, "3376964387")
        self.gb_mat_cost_entry.grid(row=1, column=1, padx=5, pady=4)

        ttk.Label(gb_frame, text="قیمت تراشکاری کل:").grid(row=1, column=2, sticky="w", pady=4)
        self.gb_turn_cost_entry = ttk.Entry(gb_frame, width=10)
        self.gb_turn_cost_entry.insert(0, "3765494357")
        self.gb_turn_cost_entry.grid(row=1, column=3, padx=5, pady=4)

        ttk.Label(gb_frame, text="عملیات حرارتی کل:").grid(row=1, column=4, sticky="w", pady=4)
        self.gb_heat_cost_entry = ttk.Entry(gb_frame, width=10)
        self.gb_heat_cost_entry.insert(0, "55903932")
        self.gb_heat_cost_entry.grid(row=1, column=5, padx=5, pady=4)

        ttk.Label(gb_frame, text="قیمت قطعات استاندارد:").grid(row=2, column=0, sticky="w", pady=4)
        self.gb_std_cost_entry = ttk.Entry(gb_frame, width=20)
        self.gb_std_cost_entry.insert(0, "1200000000")
        self.gb_std_cost_entry.grid(row=2, column=1, padx=5, pady=4)

        ttk.Label(gb_frame, text="جوش و مونتاژ:").grid(row=2, column=2, sticky="w", pady=4)
        self.gb_weld_cost_entry = ttk.Entry(gb_frame, width=10)
        self.gb_weld_cost_entry.insert(0, "800000000")
        self.gb_weld_cost_entry.grid(row=2, column=3, padx=5, pady=4)

        ttk.Label(gb_frame, text="رنگ‌آمیزی:").grid(row=2, column=4, sticky="w", pady=4)
        self.gb_paint_cost_entry = ttk.Entry(gb_frame, width=10)
        self.gb_paint_cost_entry.insert(0, "30000000")
        self.gb_paint_cost_entry.grid(row=2, column=5, padx=5, pady=4)

        ttk.Label(gb_frame, text="بسته‌بندی:").grid(row=3, column=0, sticky="w", pady=4)
        self.gb_pack_cost_entry = ttk.Entry(gb_frame, width=20)
        self.gb_pack_cost_entry.insert(0, "40000000")
        self.gb_pack_cost_entry.grid(row=3, column=1, padx=5, pady=4)

        ttk.Label(gb_frame, text="ارسال و حمل:").grid(row=3, column=2, sticky="w", pady=4)
        self.gb_ship_cost_entry = ttk.Entry(gb_frame, width=10)
        self.gb_ship_cost_entry.insert(0, "90000000")
        self.gb_ship_cost_entry.grid(row=3, column=3, padx=5, pady=4)

        # Calculate Button
        btn_calc_gb = ttk.Button(gb_frame, text="🧮 محاسبه آنالیز قیمت گیربکس", command=self.calculate_gearbox)
        btn_calc_gb.grid(row=3, column=4, columnspan=2, sticky="ew", padx=5, pady=4)

        # Output Summary Card
        res_frame = ttk.LabelFrame(self.tab_gearbox, text=" نتایج محاسباتی گیربکس (بر اساس وزن و سود) ", padding="15")
        res_frame.pack(fill="x", pady=10)

        self.gb_res_total_weight = tk.Label(res_frame, text="وزن کل قطعات: ۰ کیلوگرم", font=("Tahoma", 10, "bold"), bg="#ffffff", fg="#2c3e50")
        self.gb_res_total_weight.grid(row=0, column=0, padx=20, pady=5, sticky="w")

        self.gb_res_make_cost = tk.Label(res_frame, text="قیمت ساخت کل: ۰ ریال", font=("Tahoma", 10, "bold"), bg="#ffffff", fg="#2c3e50")
        self.gb_res_make_cost.grid(row=0, column=1, padx=20, pady=5, sticky="w")

        self.gb_res_final_set = tk.Label(res_frame, text="قیمت هر ست با ۳۰٪ سود: ۰ ریال", font=("Tahoma", 11, "bold"), bg="#ffffff", fg="#16a085")
        self.gb_res_final_set.grid(row=1, column=0, padx=20, pady=5, sticky="w")

        self.gb_res_per_kg = tk.Label(res_frame, text="قیمت هر کیلوگرم گیربکس: ۰ ریال", font=("Tahoma", 11, "bold"), bg="#ffffff", fg="#2980b9")
        self.gb_res_per_kg.grid(row=1, column=1, padx=20, pady=5, sticky="w")

    def calculate_gearbox(self):
        try:
            qty = int(self.gb_qty_entry.get())
            unit_w = float(self.gb_unit_w_entry.get())
            total_w = qty * unit_w

            mat_c = float(self.gb_mat_cost_entry.get())
            turn_c = float(self.gb_turn_cost_entry.get())
            heat_c = float(self.gb_heat_cost_entry.get())
            std_c = float(self.gb_std_cost_entry.get())
            weld_c = float(self.gb_weld_cost_entry.get())
            paint_c = float(self.gb_paint_cost_entry.get())
            pack_c = float(self.gb_pack_cost_entry.get())
            ship_c = float(self.gb_ship_cost_entry.get())

            total_make = mat_c + turn_c + heat_c + std_c + weld_c + paint_c + pack_c + ship_c
            final_with_profit = total_make * 1.30
            price_per_kg = final_with_profit / total_w if total_w > 0 else 0

            self.gb_res_total_weight.config(text=f"وزن کل قطعات: {total_w:,.1f} کیلوگرم")
            self.gb_res_make_cost.config(text=f"قیمت ساخت کل: {total_make:,.0f} ریال")
            self.gb_res_final_set.config(text=f"قیمت کل پروژه با ۳۰٪ سود: {final_with_profit:,.0f} ریال")
            self.gb_res_per_kg.config(text=f"قیمت هر کیلوگرم گیربکس: {price_per_kg:,.0f} ریال/کیلوگرم")

        except Exception as e:
            messagebox.showerror("خطا", f"لطفاً ورودی‌های گیربکس را بررسی کنید:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TookaTarhCostApp(root)
    root.mainloop()
