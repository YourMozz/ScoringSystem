import tkinter as tk
from tkinter import ttk, messagebox
import re
import pandas as pd
from datetime import datetime
import os

EXCEL_FILE = "credit_applications.xlsx"
JOBS_FILE = "jobs_salary.xlsx"
EXPERIENCE_FILE = "work_experience.xlsx"
PLACE_FILE = "place_of_work.xlsx"
COLUMNS = [
    "Номер заявки", "Телефон", "Возраст", "Пол", "Семейное положение",
    "Возраст супруга", "Пол супруга", "Дети", "Профессия", "Стаж", "Округ",
    "Зарплата", "Зарплата супруга", "Сумма кредита", "Срок",
    "Процентная ставка", "Ежемесячный платеж", "Коэффициент",
    "Наличие судимостей", "Статус", "Причина отказа", "Выбранный тип кредита"
]

if not os.path.exists(EXCEL_FILE):
    pd.DataFrame(columns=COLUMNS).to_excel(EXCEL_FILE, index=False)

def load_job_salaries():
    if os.path.exists(JOBS_FILE):
        try:
            df = pd.read_excel(JOBS_FILE)
            df.iloc[:, 0] = df.iloc[:, 0].str.strip().str.lower()
            return df
        except:
            return pd.DataFrame()
    return pd.DataFrame()

def load_experience_coefficients():
    if os.path.exists(EXPERIENCE_FILE):
        try:
            df = pd.read_excel(EXPERIENCE_FILE)
            return dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
        except:
            return {}
    return {}

def load_place_coefficients():
    if os.path.exists(PLACE_FILE):
        try:
            df = pd.read_excel(PLACE_FILE)
            df.iloc[:, 0] = df.iloc[:, 0].str.strip().str.lower()
            return dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
        except:
            return {}
    return {}

class CreditOptionsWindow(tk.Toplevel):
    def __init__(self, parent, variants, app_data):
        super().__init__(parent)
        self.app_data = app_data.copy()
        self.variants = variants
        self.geometry("800x400")
        
        main_frame = ttk.Frame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Label(main_frame, text="Выберите вариант кредита:").pack(pady=10)
        
        columns = ("Тип", "Сумма", "Срок", "Ставка", "Платеж")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=5)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        
        for v in self.variants:
            self.tree.insert("", "end", values=v[:5])
        
        self.tree.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        ttk.Button(main_frame, text="Подтвердить", command=self.save_selection).pack(pady=10)

    def save_selection(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Ошибка", "Выберите вариант!")
            return

        item = self.tree.item(selected_item)
        selected_type = item['values'][0]

        for variant in self.variants:
            if variant[0] == selected_type:
                self.app_data.update({
                    "Сумма кредита": float(variant[1].split()[0]),
                    "Срок": int(variant[2]),
                    "Процентная ставка": float(variant[3].replace('%', '')),
                    "Ежемесячный платеж": float(variant[4].split()[0]),
                    "Коэффициент": float(variant[5]),
                    "Выбранный тип кредита": selected_type,
                    "Статус": "Одобрено"
                })
                break

        try:
            df = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame(columns=COLUMNS)
            full_data = {col: self.app_data.get(col, None) for col in COLUMNS}
            new_df = pd.DataFrame([full_data])
            df = pd.concat([df, new_df], ignore_index=True)
            df[COLUMNS].to_excel(EXCEL_FILE, index=False)
            messagebox.showinfo("Успех", "Данные сохранены!")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

class ThirdWindow(tk.Toplevel):
    def __init__(self, parent, fields, app_data):
        super().__init__(parent)
        self.app_data = app_data
        self.geometry("400x300")
        
        self.vars = {}
        ttk.Label(self, text="Подтвердите пункты:").pack(pady=10)
        
        for field in fields:
            frame = ttk.Frame(self)
            frame.pack(fill='x', padx=5, pady=5)
            ttk.Label(frame, text=field, width=20).pack(side='left')
            var = tk.StringVar(value="Нет")
            self.vars[field] = var
            ttk.Radiobutton(frame, text="Да", variable=var, value="Да").pack(side='left')
            ttk.Radiobutton(frame, text="Нет", variable=var, value="Нет").pack(side='left')
        
        ttk.Button(self, text="Далее", command=self.process).pack(pady=15)

    def process(self):
        has_no = any(var.get() == "Нет" for var in self.vars.values())
        self.destroy()
        show_credit_options(self.app_data, has_no)

def calculate_coefficient(loan_amount, loan_term, rate, app_data):
    monthly_rate = rate / 100 / 12
    case = (monthly_rate * (1 + monthly_rate)**loan_term) / ((1 + monthly_rate)**loan_term - 1)
    loan_payment = loan_amount * case
    total_income = app_data['Зарплата'] + app_data.get('Зарплата супруга', 0)
    children = app_data['Дети']
    user_age = app_data['Возраст']
    gender = app_data['Пол']
    user_contribution = 27302 if (18 <= user_age <= (63 if gender == "Мужской" else 58)) else 17897
    spouse_contribution = 0
    
    if app_data['Семейное положение'] == "Женат/Замужем":
        spouse_age = app_data.get('Возраст супруга', 0)
        spouse_gender = app_data.get('Пол супруга', '')
        if spouse_gender:
            spouse_contribution = 27302 if (18 <= spouse_age <= (63 if spouse_gender == "Мужской" else 58)) else 17897
    
    denominator = (20663 * children) + user_contribution + spouse_contribution
    return (total_income - loan_payment) / denominator if denominator > 0 else 0

def adjust_loan_term(loan_amount, initial_term, rate, app_data):
    current_term = initial_term
    while True:
        coeff = calculate_coefficient(loan_amount, current_term, rate, app_data)
        if coeff >= 1.0:
            return current_term, coeff
        current_term += 1

def calculate_payment(amount, term, rate):
    monthly_rate = rate / 100 / 12
    case = (monthly_rate * (1 + monthly_rate)**term) / ((1 + monthly_rate)**term - 1)
    return amount * case

def show_credit_options(app_data, has_restriction):
    base_amount = app_data['Сумма кредита']
    base_term = app_data['Срок']
    base_rate = app_data['Процентная ставка']
    variants = []
    
    if has_restriction:
        new_amount = min(base_amount, 250000)
        adjusted_term, adjusted_coeff = adjust_loan_term(new_amount, base_term, base_rate, app_data)
        payment = calculate_payment(new_amount, adjusted_term, base_rate)
        variants.append(("Ограниченный", f"{new_amount:.2f} руб", adjusted_term, f"{base_rate:.1f}%", f"{payment:.2f} руб", adjusted_coeff))
    else:
        adjusted_term, adjusted_coeff = adjust_loan_term(base_amount, base_term, base_rate, app_data)
        payment = calculate_payment(base_amount, adjusted_term, base_rate)
        variants.append(("Исходный", f"{base_amount:.2f} руб", adjusted_term, f"{base_rate:.1f}%", f"{payment:.2f} руб", adjusted_coeff))
        amount_plus = base_amount * 1.15
        rate_plus = base_rate - 3
        term_plus, coeff_plus = adjust_loan_term(amount_plus, base_term, rate_plus, app_data)
        payment_plus = calculate_payment(amount_plus, term_plus, rate_plus)
        variants.append(("Увеличенный+", f"{amount_plus:.2f} руб", term_plus, f"{rate_plus:.1f}%", f"{payment_plus:.2f} руб", coeff_plus))
        amount_minus = base_amount * 0.85
        rate_minus = base_rate + 3
        term_minus, coeff_minus = adjust_loan_term(amount_minus, base_term, rate_minus, app_data)
        payment_minus = calculate_payment(amount_minus, term_minus, rate_minus)
        variants.append(("Уменьшенный-", f"{amount_minus:.2f} руб", term_minus, f"{rate_minus:.1f}%", f"{payment_minus:.2f} руб", coeff_minus))
    
    CreditOptionsWindow(app, variants, app_data)

class ApplicationWindow(tk.Toplevel):
    def __init__(self, parent, app_data):
        super().__init__(parent)
        self.app_data = app_data
        self.geometry("600x500")
        
        self.selection_vars = {
            'Зарплата': tk.StringVar(value="Нет"),
            'Зарплата супруга': tk.StringVar(value="Нет"),
            'Стаж': tk.StringVar(value="Нет"),
            'Округ': tk.StringVar(value="Нет")
        }
        
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=10, pady=10)
        
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scroll_frame = ttk.Frame(canvas)
        
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        row = 0
        ttk.Label(self.scroll_frame, text=f"Номер заявки: {app_data['Номер заявки']}").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(self.scroll_frame, text=f"{app_data['Пол']}, {app_data['Возраст']} лет").grid(row=row, column=0, sticky=tk.W)
        row += 1
        
        ttk.Label(self.scroll_frame, text=f"Стаж: {app_data['Стаж']} лет").grid(row=row, column=0, sticky=tk.W)
        self.add_radio_buttons(row, 'Стаж')
        row += 1
        
        ttk.Label(self.scroll_frame, text=f"Округ: {app_data['Округ']}").grid(row=row, column=0, sticky=tk.W)
        self.add_radio_buttons(row, 'Округ')
        row += 1
        
        spouse_status = 'Да' if app_data['Семейное положение'] == 'Женат/Замужем' else 'Нет'
        ttk.Label(self.scroll_frame, text=f"Наличие супруга: {spouse_status}").grid(row=row, column=0, sticky=tk.W)
        row += 1

        if spouse_status == 'Да':
            ttk.Label(self.scroll_frame, text=f"Пол супруга: {app_data.get('Пол супруга', '')}").grid(row=row, column=0, sticky=tk.W)
            row += 1
            ttk.Label(self.scroll_frame, text=f"Возраст супруга: {app_data.get('Возраст супруга', '')} лет").grid(row=row, column=0, sticky=tk.W)
            row += 1
            
            spouse_salary = app_data.get('Зарплата супруга', 0)
            ttk.Label(self.scroll_frame, text=f"Зарплата супруга: {spouse_salary} руб").grid(row=row, column=0, sticky=tk.W)
            self.add_radio_buttons(row, 'Зарплата супруга')
            row += 1

        ttk.Label(self.scroll_frame, text=f"Дети: {app_data['Дети']}").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(self.scroll_frame, text=f"Профессия: {app_data['Профессия']}").grid(row=row, column=0, sticky=tk.W)
        row += 1

        if app_data.get('salary_warning'):
            ttk.Label(self.scroll_frame, text=app_data['salary_warning'], foreground='red').grid(row=row, column=0, sticky=tk.W)
            row += 1

        ttk.Label(self.scroll_frame, text=f"Зарплата: {app_data['Зарплата']} руб").grid(row=row, column=0, sticky=tk.W)
        self.add_radio_buttons(row, 'Зарплата')
        row += 1

        ttk.Label(self.scroll_frame, text=f"Сумма кредита: {app_data['Сумма кредита']} руб").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(self.scroll_frame, text=f"Срок: {app_data['Срок']} мес").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(self.scroll_frame, text=f"Процентная ставка: {app_data['Процентная ставка']:.1f}%").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(self.scroll_frame, text=f"Ежемесячный платеж: {app_data['Ежемесячный платеж']:.2f} руб").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(self.scroll_frame, text=f"Коэффициент: {app_data['Коэффициент']:.2f}").grid(row=row, column=0, sticky=tk.W)
        row += 1
        ttk.Label(self.scroll_frame, text=f"Наличие судимостей: {app_data['Наличие судимостей']}").grid(row=row, column=0, sticky=tk.W)
        row += 1

        ttk.Button(self.scroll_frame, text="Отправить", command=self.submit).grid(row=row, column=0, pady=15)
        
    def add_radio_buttons(self, row, field):
        frame = ttk.Frame(self.scroll_frame)
        frame.grid(row=row, column=1, sticky=tk.E)
        ttk.Radiobutton(frame, text="Да", variable=self.selection_vars[field], value="Да").pack(side='left')
        ttk.Radiobutton(frame, text="Нет", variable=self.selection_vars[field], value="Нет").pack(side='left')

    def submit(self):
        selected_fields = [field for field, var in self.selection_vars.items() if var.get() == "Да"]
        self.save_to_excel("В обработке")
        self.destroy()
        
        if selected_fields:
            ThirdWindow(self.master, selected_fields, self.app_data)
        else:
            show_credit_options(self.app_data, False)

    def save_to_excel(self, status):
        try:
            new_data = {col: self.app_data.get(col, None) for col in COLUMNS}
            new_data.update({
                "Статус": status,
                "Причина отказа": "",
                "Выбранный тип кредита": ""
            })
            
            df = pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame(columns=COLUMNS)
            new_df = pd.DataFrame([new_data])
            df = pd.concat([df, new_df], ignore_index=True)
            df[COLUMNS].to_excel(EXCEL_FILE, index=False)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка сохранения: {str(e)}")

def validate_phone(phone):
    return re.match(r'^89\d{9}$', phone)

def update_children_field():
    if children_var.get() == "Да":
        children_label.grid(row=8, column=0)
        children_count_entry.grid(row=8, column=1)
    else:
        children_label.grid_remove()
        children_count_entry.grid_remove()

def update_spouse_fields():
    if marital_status_var.get() == "Женат/Замужем":
        spouse_salary_label.grid(row=5, column=0)
        spouse_salary_entry.grid(row=5, column=1)
        spouse_age_label.grid(row=6, column=0)
        spouse_age_entry.grid(row=6, column=1)
    else:
        spouse_salary_label.grid_remove()
        spouse_salary_entry.grid_remove()
        spouse_age_label.grid_remove()
        spouse_age_entry.grid_remove()

def update_convictions_field(*args):
    try:
        loan_amount = float(loan_amount_entry.get())
        if loan_amount > 250000:
            convictions_label.grid(row=15, column=0)
            convictions_yes.grid(row=15, column=1)
            convictions_no.grid(row=15, column=2)
        else:
            convictions_label.grid_remove()
            convictions_yes.grid_remove()
            convictions_no.grid_remove()
            convictions_var.set("Нет")
    except:
        convictions_label.grid_remove()
        convictions_yes.grid_remove()
        convictions_no.grid_remove()
        convictions_var.set("Нет")

def submit():
    try:
        app_data = {
            'Номер заявки': datetime.now().strftime("%Y%m%d%H%M%S"),
            'Телефон': phone_entry.get(),
            'Возраст': int(age_entry.get()),
            'Пол': gender_var.get(),
            'Семейное положение': marital_status_var.get(),
            'Дети': int(children_count_entry.get()) if children_var.get() == "Да" else 0,
            'Профессия': job_entry.get(),
            'Стаж': int(experience_entry.get()),
            'Округ': place_var.get().strip(),
            'Зарплата': float(salary_entry.get()),
            'Сумма кредита': float(loan_amount_entry.get()),
            'Срок': int(loan_term_entry.get()),
            'Наличие судимостей': convictions_var.get() if float(loan_amount_entry.get()) > 250000 else "Нет"
        }

        if not validate_phone(app_data['Телефон']):
            raise ValueError("Неверный формат телефона")

        if app_data['Возраст'] < 18 or app_data['Возраст'] > 100:
            raise ValueError("Некорректный возраст")

        exp_coeffs = load_experience_coefficients()
        experience_coeff = 1.0
        if exp_coeffs:
            valid_exp = max([e for e in exp_coeffs.keys() if e <= app_data['Стаж']], default=None)
            if valid_exp is not None:
                experience_coeff = exp_coeffs[valid_exp]

        place_coeff = 1.0
        place_coeffs = load_place_coefficients()
        if place_coeffs:
            place_coeff = place_coeffs.get(app_data['Округ'].lower(), 1.0)

        df_jobs = load_job_salaries()
        if not df_jobs.empty and app_data['Профессия']:
            job_clean = app_data['Профессия'].strip().lower()
            match = df_jobs[df_jobs.iloc[:, 0] == job_clean]
            if not match.empty:
                min_salary = match.iloc[0, 1] * experience_coeff * place_coeff
                max_salary = match.iloc[0, 2] * experience_coeff * place_coeff
                if app_data['Зарплата'] < min_salary or app_data['Зарплата'] > max_salary:
                    app_data['salary_warning'] = f"Несоответствие зарплаты! ({min_salary:.2f}-{max_salary:.2f} руб)"

        if app_data['Семейное положение'] == "Женат/Замужем":
            app_data['Зарплата супруга'] = float(spouse_salary_entry.get())
            app_data['Возраст супруга'] = int(spouse_age_entry.get())
            app_data['Пол супруга'] = "Женский" if app_data['Пол'] == "Мужской" else "Мужской"

        initial_rate = 30.0
        denominator_rate = place_coeff * experience_coeff * 0.8
        adjusted_rate = initial_rate / denominator_rate
        
        rate = adjusted_rate / 100 / 12
        case = (rate * (1+rate)**app_data['Срок']) / ((1+rate)**app_data['Срок'] - 1)
        loan_payment = app_data['Сумма кредита'] * case
        
        total_income = app_data['Зарплата'] + app_data.get('Зарплата супруга', 0)
        user_contribution = 27302 if (18 <= app_data['Возраст'] <= (63 if app_data['Пол'] == "Мужской" else 58)) else 17897
        spouse_contribution = 0
        
        if app_data['Семейное положение'] == "Женат/Замужем":
            spouse_age = app_data.get('Возраст супруга', 0)
            spouse_gender = app_data.get('Пол супруга', '')
            if spouse_gender:
                spouse_contribution = 27302 if (18 <= spouse_age <= (63 if spouse_gender == "Мужской" else 58)) else 17897
        
        denominator = (20663 * app_data['Дети']) + user_contribution + spouse_contribution
        result = (total_income - loan_payment) / denominator if denominator > 0 else 0

        app_data.update({
            'Процентная ставка': adjusted_rate,
            'Ежемесячный платеж': loan_payment,
            'Коэффициент': result
        })

        ApplicationWindow(app, app_data)

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

app = tk.Tk()
app.title("Кредитная анкета")
app.geometry("500x700")

main_frame = ttk.Frame(app)
main_frame.pack(fill='both', expand=True, padx=10, pady=10)

row = 0
ttk.Label(main_frame, text="Телефон (89XXXXXXXXX):").grid(row=row, column=0, sticky=tk.W)
phone_entry = ttk.Entry(main_frame)
phone_entry.grid(row=row, column=1)
row += 1

ttk.Label(main_frame, text="Возраст:").grid(row=row, column=0, sticky=tk.W)
age_entry = ttk.Entry(main_frame)
age_entry.grid(row=row, column=1)
row += 1

ttk.Label(main_frame, text="Пол:").grid(row=row, column=0, sticky=tk.W)
gender_var = tk.StringVar(value="Мужской")
ttk.Radiobutton(main_frame, text="Мужской", variable=gender_var, value="Мужской").grid(row=row, column=1, sticky=tk.W)
ttk.Radiobutton(main_frame, text="Женский", variable=gender_var, value="Женский").grid(row=row, column=2, sticky=tk.W)
row += 1

ttk.Label(main_frame, text="Семейное положение:").grid(row=row, column=0, sticky=tk.W)
marital_status_var = tk.StringVar(value="Холост")
ttk.Radiobutton(main_frame, text="Холост", variable=marital_status_var, value="Холост", command=update_spouse_fields).grid(row=row, column=1, sticky=tk.W)
ttk.Radiobutton(main_frame, text="Женат/Замужем", variable=marital_status_var, value="Женат/Замужем", command=update_spouse_fields).grid(row=row, column=2, sticky=tk.W)
row += 1

spouse_salary_label = ttk.Label(main_frame, text="Зарплата супруга:")
spouse_salary_entry = ttk.Entry(main_frame)
spouse_age_label = ttk.Label(main_frame, text="Возраст супруга:")
spouse_age_entry = ttk.Entry(main_frame)

ttk.Label(main_frame, text="Наличие детей:").grid(row=row, column=0, sticky=tk.W)
children_var = tk.StringVar(value="Нет")
ttk.Radiobutton(main_frame, text="Нет", variable=children_var, value="Нет", command=update_children_field).grid(row=row, column=1, sticky=tk.W)
ttk.Radiobutton(main_frame, text="Да", variable=children_var, value="Да", command=update_children_field).grid(row=row, column=2, sticky=tk.W)
row += 1

children_label = ttk.Label(main_frame, text="Количество детей:")
children_count_entry = ttk.Entry(main_frame)

ttk.Label(main_frame, text="Профессия:").grid(row=row, column=0, sticky=tk.W)
job_entry = ttk.Entry(main_frame)
job_entry.grid(row=row, column=1)
row += 1

ttk.Label(main_frame, text="Стаж работы (лет):").grid(row=row, column=0, sticky=tk.W)
experience_entry = ttk.Entry(main_frame)
experience_entry.grid(row=row, column=1)
row += 1

ttk.Label(main_frame, text="Административный округ:").grid(row=row, column=0, sticky=tk.W)
place_var = tk.StringVar()
place_combobox = ttk.Combobox(main_frame, textvariable=place_var)
place_combobox['values'] = [
    'Центральный', 'Северный', 'Северо-Восточный', 'Восточный',
    'Юго-Восточный', 'Южный', 'Юго-Западный', 'Западный',
    'Северо-Западный', 'Зеленоградский', 'Троицкий', 'Новомосковский'
]
place_combobox.grid(row=row, column=1)
row += 1

ttk.Label(main_frame, text="Зарплата:").grid(row=row, column=0, sticky=tk.W)
salary_entry = ttk.Entry(main_frame)
salary_entry.grid(row=row, column=1)
row += 1

ttk.Label(main_frame, text="Сумма кредита:").grid(row=row, column=0, sticky=tk.W)
loan_amount_entry = ttk.Entry(main_frame)
loan_amount_entry.grid(row=row, column=1)
loan_amount_entry.bind("<KeyRelease>", update_convictions_field)
row += 1

ttk.Label(main_frame, text="Срок (мес):").grid(row=row, column=0, sticky=tk.W)
loan_term_entry = ttk.Entry(main_frame)
loan_term_entry.grid(row=row, column=1)
row += 1

convictions_var = tk.StringVar(value="Нет")
convictions_label = ttk.Label(main_frame, text="Наличие судимостей:")
convictions_yes = ttk.Radiobutton(main_frame, text="Да", variable=convictions_var, value="Да")
convictions_no = ttk.Radiobutton(main_frame, text="Нет", variable=convictions_var, value="Нет")

ttk.Button(main_frame, text="Отправить заявку", command=submit).grid(row=row, columnspan=3, pady=15)

update_children_field()
update_spouse_fields()
update_convictions_field()

app.mainloop()