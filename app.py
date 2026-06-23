import pandas as pd
import tkinter as tk
import random
from datetime import datetime
import matplotlib.pyplot as plt
import os
from tkinter import messagebox  

# پیدا کردن مسیرِ دقیقِ فایل اکسل
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(current_dir, "prices.xlsx")

# ================== تعریف فونت‌ها (ظاهر مدرن) ==================
FONT = ("Segoe UI", 12)
FONT_B = ("Segoe UI", 12, "bold")
TITLE = ("Segoe UI", 22, "bold")
SMALL = ("Segoe UI", 10)
PRICE_FONT = ("Segoe UI", 16, "bold")

# ================== خواندن داده‌ها ==================
try:
    data = pd.read_excel(excel_path)
except FileNotFoundError:
    messagebox.showerror("خطا", f"فایل prices.xlsx پیدا نشد:\n{excel_path}")
    raise 

if 'Date' not in data.columns:
    data['Date'] = pd.date_range(end=datetime.today(), periods=len(data), freq='D')

# ================== لیست ذخیره تاریخچه قیمت‌ها ==================
# این دیکشنری تمام قیمت‌های لود شده از زمان باز شدن برنامه را در خود نگه می‌دارد
price_history = {
    'Time': [],
    'Gold18': [],
    'Gold24': [],
    'Dollar': [],
    'Coin': []
}

# ================== فرمت قیمت ==================
def format_price(value):
    try:
        value = round(float(value))
        return f"{value:,}"
    except:
        return "0"

# ================== Splash Screen ==================
def show_splash():
    splash = tk.Toplevel()
    splash.title("Loading...")
    splash.geometry("460x320")
    splash.configure(bg="#478EC1")
    splash.overrideredirect(True)
    
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - 460) // 2
    y = (screen_height - 320) // 2
    splash.geometry(f"460x320+{x}+{y}")

    tk.Label(splash, text="Gold Price", font=("Segoe UI", 28, "bold"), 
             fg="white", bg="#1e3a8a").pack(pady=(50,8))
    
    tk.Label(splash, text="نظارت قیمت طلا و دلار", font=("Segoe UI", 16), 
             fg="#eab308", bg="#1e3a8a").pack(pady=5)
    
    tk.Label(splash, text="پریسا قویدل", font=("Segoe UI", 15, "bold"), 
             fg="#a5b4fc", bg="#1e3a8a").pack(pady=12)
    
    tk.Label(splash, text="پروژه کارشناسی علوم کامپیوتر", font=("Segoe UI", 11), 
             fg="#cbd5e1", bg="#1a202c").pack(pady=5)

    progress = tk.Frame(splash, bg="#eab308", height=4, width=240)
    progress.pack(pady=25)
    
    tk.Label(splash, text="در حال بارگذاری...", font=("Segoe UI", 11), 
             fg="#cbd5e1", bg="#1a202c").pack()

    splash.after(3000, splash.destroy)
    return splash

# شروع برنامه و مخفی کردن پنجره اصلی برای اسپلش
root = tk.Tk()
root.withdraw()

previous_prices = None 

splash = show_splash()
root.after(3000, lambda: root.deiconify())

# ================== رابط کاربری اصلی ==================
root.title("Gold Price - پریسا قویدل")
root.geometry("400x740")  
root.configure(bg="#f8fafc")
root.resizable(False, False)
                    #نوار بالای برنامه
header = tk.Frame(root, bg="#3b82f6", height=90)
header.pack(fill="x")
header.pack_propagate(False)

tk.Label(header, text="قیمت طلا و دلار", font=TITLE,
         fg="white", bg="#3b82f6").pack(pady=25)

update_time_label = tk.Label(root, text="", font=SMALL, fg="#555", bg="#f8fafc")
update_time_label.pack(pady=5)

# فریم نمایش قیمت‌ها
price_frame = tk.Frame(root, bg="#f8fafc", padx=15, pady=12)
price_frame.pack(pady=12, padx=15, fill="x")

gold18_label = tk.Label(price_frame, text="", font=PRICE_FONT, fg="#d97706", bg="#f8fafc")
gold24_label = tk.Label(price_frame, text="", font=PRICE_FONT, fg="#b45309", bg="#f8fafc")
dollar_label = tk.Label(price_frame, text="", font=PRICE_FONT, fg="#15803d", bg="#f8fafc")
coin_label   = tk.Label(price_frame, text="", font=PRICE_FONT, fg="#854d0e", bg="#f8fafc")

for title, label in [
    ("طلای ۱۸ عیار", gold18_label),
    ("طلای ۲۴ عیار", gold24_label),
    ("دلار آمریکا", dollar_label),
    ("سکه امامی", coin_label)
]:
    tk.Label(price_frame, text=title, font=FONT, bg="#f8fafc", anchor="w").pack(fill="x", pady=(8,1))
    label.pack(fill="x", pady=(0,8))

# ================== توابع برنامه ==================

def get_current_prices():
    try:
        return {
            'Gold18': float(gold18_label.cget("text").replace(" تومان", "").replace(",", "")),
            'Gold24': float(gold24_label.cget("text").replace(" تومان", "").replace(",", "")),
            'Dollar': float(dollar_label.cget("text").replace(" تومان", "").replace(",", "")),
            'Coin':   float(coin_label.cget("text").replace(" تومان", "").replace(",", ""))
        }
    except:
        return None

def refresh_price():
    global previous_prices

    # ۱. ذخیره قیمت فعلی به عنوان قیمت قبلی (برای درصد تغییر)
    current = get_current_prices()
    if current is not None:
        previous_prices = current
    
    # ۲. خواندن ردیف تصادفی جدید
    index = random.randint(0, len(data) - 1)
    row = data.iloc[index]
    
    # ۳. استخراج مقادیر عددی جدید
    g18 = float(row['Gold18'])
    g24 = float(row['Gold24'])
    usd = float(row['Dollar'])
    coin = float(row['Coin'])

    # ۴. آپدیت لیبل‌های رابط کاربری
    gold18_label.config(text=f"{format_price(g18)} تومان")
    gold24_label.config(text=f"{format_price(g24)} تومان")
    dollar_label.config(text=f"{format_price(usd)} تومان")
    coin_label.config(text=  f"{format_price(coin)} تومان")
    
    # ۵. آپدیت زمان آخرین بروزرسانی
    now = datetime.now()
    now_str = now.strftime("%H:%M:%S")
    update_time_label.config(text=f"بروزرسانی شده در: {now.strftime('%H:%M:%S - %Y/%m/%d')}")

    # ۶. اضافه کردن این قیمت‌های جدید به تاریخچه برای نمودار صعودی/نزولی
    price_history['Time'].append(now_str)
    price_history['Gold18'].append(g18)
    price_history['Gold24'].append(g24)
    price_history['Dollar'].append(usd)
    price_history['Coin'].append(coin)

# نمودار درصد تغییر نسبت به قیمت قبلی (نمودار میله‌ای)
def show_percentage_chart():
    if previous_prices is None:
        messagebox.showinfo("اطلاعات ناکافی", "لطفاً حداقل یک بار دیگر دکمه بروزرسانی را بزنید.")
        return
    
    current = get_current_prices()
    if not current:
        return
    
    categories = ['Gold 18K', 'Gold 24K', 'USD', 'Coin']
    prev_values = [previous_prices['Gold18'], previous_prices['Gold24'], previous_prices['Dollar'], previous_prices['Coin']]
    curr_values = [current['Gold18'], current['Gold24'], current['Dollar'], current['Coin']]
    
    changes = []
    for prev, curr in zip(prev_values, curr_values):
        change = ((curr - prev) / prev * 100) if prev != 0 else 0
        changes.append(change)
    
    fig, ax = plt.subplots(figsize=(8, 5))
    colors = ['#22c55e' if c >= 0 else '#ef4444' for c in changes]
    bars = ax.bar(categories, changes, color=colors, alpha=0.85)
    
    ax.set_title("Percentage Change from Immediate Previous Price (%)", fontsize=12, fontweight='bold')
    ax.set_ylabel("Change (%)")
    ax.grid(axis='y', alpha=0.3)
    
    for bar, change in zip(bars, changes):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + (0.2 if height >= 0 else -0.5),
                f"{change:+.2f}%", ha='center', va='bottom' if height >= 0 else 'top', fontweight='bold')
    
    plt.tight_layout()
    plt.show()

# نمودار روند صعودی و نزولی (نمودار خطی تاریخچه)
def show_trend_chart():
    # برای کشیدن نمودار صعودی/نزولی حداقل باید ۲ نقطه (داده) ثبت شده باشد
    if len(price_history['Time']) < 2:
        messagebox.showinfo("تاریخچه ناکافی", "برای مشاهده نمودار روند صعودی/نزولی، باید حداقل ۲ بار دکمه بروزرسانی را زده باشید.")
        return

    # ساخت یک شکل با ۴ زیرنمودار مجزا (۲ در ۲)
    fig, axs = plt.subplots(2, 2, figsize=(11, 8))
    fig.suptitle("Historical Price Trends (Session)", fontsize=14, fontweight='bold')

    times = price_history['Time']

    # ۱. نمودار طلای ۱۸
    axs[0, 0].plot(times, price_history['Gold18'], marker='o', color='#d97706', linewidth=2)
    axs[0, 0].set_title("Gold 18K Trend")
    axs[0, 0].set_ylabel("Price (Toman)")
    axs[0, 0].grid(True, alpha=0.3)
    axs[0, 0].tick_params(axis='x', rotation=30)

    # ۲. نمودار طلای ۲۴
    axs[0, 1].plot(times, price_history['Gold24'], marker='o', color='#b45309', linewidth=2)
    axs[0, 1].set_title("Gold 24K Trend")
    axs[0, 1].grid(True, alpha=0.3)
    axs[0, 1].tick_params(axis='x', rotation=30)

    # ۳. نمودار دلار
    axs[1, 0].plot(times, price_history['Dollar'], marker='o', color='#15803d', linewidth=2)
    axs[1, 0].set_title("USD Trend")
    axs[1, 0].set_ylabel("Price (Toman)")
    axs[1, 0].grid(True, alpha=0.3)
    axs[1, 0].tick_params(axis='x', rotation=30)

    # ۴. نمودار سکه امامی
    axs[1, 1].plot(times, price_history['Coin'], marker='o', color='#854d0e', linewidth=2)
    axs[1, 1].set_title("Emami Coin Trend")
    axs[1, 1].grid(True, alpha=0.3)
    axs[1, 1].tick_params(axis='x', rotation=30)

    plt.tight_layout()
    plt.show()

# ================== دکمه‌ها ==================
btn_frame = tk.Frame(root, bg="#f8fafc")
btn_frame.pack(pady=15)

# دکمه بروزرسانی
tk.Button(btn_frame, text="🔄 بروزرسانی قیمت", font=FONT_B,
          bg="#3b82f6", fg="white",
          relief="flat", cursor="hand2",
          height=2, width=25,
          command=refresh_price).pack(pady=5)

# دکمه درصد تغییر نسبت به قیمت قبلی
tk.Button(btn_frame, text="📊 درصد تغییر لحظه‌ای", font=FONT_B,
          bg="#10b981", fg="white", # سبز ملایم
          relief="flat", cursor="hand2",
          height=2, width=25,
          command=show_percentage_chart).pack(pady=5)

# دکمه روند کلی صعودی و نزولی (جدید)
tk.Button(btn_frame, text="📈 نمودار روند صعودی/نزولی", font=FONT_B,
          bg="#6366f1", fg="white", # رنگ بنفش/آبی ملایم شیک
          relief="flat", cursor="hand2",
          height=2, width=25,
          command=show_trend_chart).pack(pady=5)

# اولین لود خودکار قیمت بعد از اسپلش اسکرین
root.after(3050, refresh_price)

root.mainloop()
