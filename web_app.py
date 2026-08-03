#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""وب‌اپ حسابداری هلیا بیوتی — Flask version"""
import os, sys, json
from datetime import datetime, timedelta
from collections import defaultdict
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.urandom(24)

# ─── Data Directory ───
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)
EMPLOYEES_FILE = os.path.join(DATA_DIR, "employees.xlsx")
SERVICES_FILE = os.path.join(DATA_DIR, "services.xlsx")
TRANSACTIONS_FILE = os.path.join(DATA_DIR, "transactions.xlsx")
CUSTOMERS_FILE = os.path.join(DATA_DIR, "customers.xlsx")

# ─── Persian Date ───
class PersianDate:
    @staticmethod
    def gregorian_to_jalali(gy, gm, gd):
        g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
        gy2 = gy + 1 if gm > 2 else gy
        days = 355666 + (365*gy) + ((gy2+3)//4) - ((gy2+99)//100) + ((gy2+399)//400) + gd + g_d_m[gm-1]
        jy = -1595 + (33*(days//12053)); days %= 12053
        jy += 4*(days//1461); days %= 1461
        if days > 365: jy += (days-1)//365; days = (days-1)%365
        if days < 186: jm = 1+(days//31); jd = 1+(days%31)
        else: jm = 7+((days-186)//30); jd = 1+((days-186)%30)
        return jy, jm, jd
    @staticmethod
    def today_str():
        now = datetime.now()
        jy, jm, jd = PersianDate.gregorian_to_jalali(now.year, now.month, now.day)
        return f"{jy}/{jm:02d}/{jd:02d}"
    @staticmethod
    def now_jalali():
        return PersianDate.today_str()

# ─── Default Data ───
DEFAULT_EMPLOYEES = [
    {"name":"مریم","specialty":"مو","phone":"","share_percent":0},
    {"name":"زهرا","specialty":"ناخن","phone":"","share_percent":0},
    {"name":"سارا","specialty":"ابرو","phone":"","share_percent":0},
    {"name":"نیلوفر","specialty":"مژه","phone":"","share_percent":0},
]
DEFAULT_SERVICES = [
    {"name":"رنگ مو","category":"مو","default_price":80000},
    {"name":"مش","category":"مو","default_price":120000},
    {"name":"کوتاهی مو","category":"مو","default_price":50000},
    {"name":"شینیون","category":"مو","default_price":100000},
    {"name":"فر دائمی","category":"مو","default_price":150000},
    {"name":"ابرو برداشتن","category":"ابرو","default_price":20000},
    {"name":"ابرو رنگ","category":"ابرو","default_price":15000},
    {"name":"تاتو ابرو","category":"ابرو","default_price":80000},
    {"name":"کاشت مژه","category":"مژه","default_price":100000},
    {"name":"اکستنشن مژه","category":"مژه","default_price":150000},
    {"name":"جلاس مژه","category":"مژه","default_price":50000},
    {"name":"مانیکور","category":"ناخن","default_price":60000},
    {"name":"پدیکور","category":"ناخن","default_price":70000},
    {"name":"ژل ناخن","category":"ناخن","default_price":80000},
    {"name":"کاشت ناخن","category":"ناخن","default_price":120000},
    {"name":"فرچ ناخن","category":"ناخن","default_price":40000},
]

# ─── Excel Manager ───
PINK = PatternFill(start_color="D81B60", end_color="D81B60", fill_type="solid")
HDR_FONT = Font(bold=True, color="FFFFFF")

def _ensure_workbook(fp, headers):
    if not os.path.exists(fp):
        wb = Workbook(); ws = wb.active; ws.title = "داده‌ها"
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = HDR_FONT; cell.fill = PINK; cell.alignment = Alignment(horizontal="center")
        wb.save(fp)
    return load_workbook(fp)

def init_all():
    _ensure_workbook(EMPLOYEES_FILE, ["نام","تخصص","تلفن","درصد سهم"])
    ws = load_workbook(EMPLOYEES_FILE).active
    if ws.max_row <= 1:
        wb = load_workbook(EMPLOYEES_FILE); ws = wb.active
        for e in DEFAULT_EMPLOYEES: ws.append([e["name"],e["specialty"],e["phone"],e["share_percent"]])
        wb.save(EMPLOYEES_FILE)
    _ensure_workbook(SERVICES_FILE, ["نام خدمت","دسته‌بندی","قیمت پیش‌فرض"])
    ws2 = load_workbook(SERVICES_FILE).active
    if ws2.max_row <= 1:
        wb2 = load_workbook(SERVICES_FILE); ws2 = wb2.active
        for s in DEFAULT_SERVICES: ws2.append([s["name"],s["category"],s["default_price"]])
        wb2.save(SERVICES_FILE)
    # Updated transaction headers with new fields
    _ensure_workbook(TRANSACTIONS_FILE, [
        "تاریخ","نام مشتری","نام خدمت","دسته‌بندی","نام کارمند",
        "مبلغ خالص","تخفیف","مبلغ نهایی","روش پرداخت","پورسانت کارمند","سهم سالن","یادداشت"
    ])
    _ensure_workbook(CUSTOMERS_FILE, ["نام","تلفن","تخصص مورد علاقه","تاریخ عضویت","تعداد مراجعه","یادداشت","امتیاز"])

def _read_rows(fp):
    if not os.path.exists(fp): return []
    wb = load_workbook(fp); ws = wb.active
    return [row for row in ws.iter_rows(min_row=2, values_only=True) if row[0]]

def get_employees():
    rows = _read_rows(EMPLOYEES_FILE)
    return [{"name":str(r[0]),"specialty":str(r[1] or ""),"phone":str(r[2] or ""),"share_percent":float(r[3] or 0)} for r in rows]

def save_employees(emps):
    wb = Workbook(); ws = wb.active; ws.title = "داده‌ها"
    for c, h in enumerate(["نام","تخصص","تلفن","درصد سهم"], 1):
        cell = ws.cell(row=1, column=c, value=h); cell.font = HDR_FONT; cell.fill = PINK
    for e in emps: ws.append([e["name"],e["specialty"],e["phone"],e["share_percent"]])
    wb.save(EMPLOYEES_FILE)

def get_services():
    rows = _read_rows(SERVICES_FILE)
    return [{"name":str(r[0]),"category":str(r[1] or ""),"default_price":int(r[2] or 0)} for r in rows]

def save_services(svcs):
    wb = Workbook(); ws = wb.active; ws.title = "داده‌ها"
    for c, h in enumerate(["نام خدمت","دسته‌بندی","قیمت پیش‌فرض"], 1):
        cell = ws.cell(row=1, column=c, value=h); cell.font = HDR_FONT; cell.fill = PINK
    for s in svcs: ws.append([s["name"],s["category"],s["default_price"]])
    wb.save(SERVICES_FILE)

def get_customers():
    rows = _read_rows(CUSTOMERS_FILE)
    result = []
    for r in rows:
        cust = {"name":str(r[0]),"phone":str(r[1] or ""),"specialty":str(r[2] or ""),"join_date":str(r[3] or ""),"visit_count":int(r[4] or 0),"note":str(r[5] or "")}
        # Handle optional points column
        cust["points"] = int(r[6]) if len(r) > 6 and r[6] else 0
        result.append(cust)
    return result

def save_customers(custs):
    wb = Workbook(); ws = wb.active; ws.title = "داده‌ها"
    for c, h in enumerate(["نام","تلفن","تخصص مورد علاقه","تاریخ عضویت","تعداد مراجعه","یادداشت","امتیاز"], 1):
        cell = ws.cell(row=1, column=c, value=h); cell.font = HDR_FONT; cell.fill = PINK
    for cu in custs: ws.append([cu["name"],cu["phone"],cu["specialty"],cu["join_date"],cu["visit_count"],cu["note"],cu.get("points",0)])
    wb.save(CUSTOMERS_FILE)

def get_transactions(start_date=None, end_date=None):
    if not os.path.exists(TRANSACTIONS_FILE): return []
    wb = load_workbook(TRANSACTIONS_FILE); ws = wb.active; txns = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]: continue
        t = {
            "date":str(row[0]),"customer":str(row[1]),"service":str(row[2]),
            "category":str(row[3]),"employee":str(row[4]),
            "amount":int(row[5] or 0),
            "discount":int(row[6] or 0),
            "final_amount":int(row[7] or 0),
            "payment_method":str(row[8] or "نقدی"),
            "commission":int(row[9] or 0),
            "salon_share":int(row[10] or 0),
            "note":str(row[11] or "")
        }
        if start_date and end_date:
            if start_date <= t["date"] <= end_date: txns.append(t)
        else: txns.append(t)
    return txns

def add_transaction(date_str, customer, service, category, employee, amount, discount=0, final_amount=0, payment_method="نقدی", commission=0, salon_share=0, note=""):
    wb = load_workbook(TRANSACTIONS_FILE); ws = wb.active
    ws.append([date_str, customer, service, category, employee, int(amount), int(discount), int(final_amount), payment_method, int(commission), int(salon_share), note])
    wb.save(TRANSACTIONS_FILE)

def delete_transaction(index):
    wb = load_workbook(TRANSACTIONS_FILE); ws = wb.active
    ws.delete_rows(index + 2); wb.save(TRANSACTIONS_FILE)

# ─── Backup Helper ───
def create_backup_excel(transactions, filename):
    wb = Workbook(); ws = wb.active; ws.title = "تراکنش‌ها"
    headers = ["تاریخ","مشتری","خدمت","دسته","کارمند","مبلغ خالص","تخفیف","مبلغ نهایی","روش پرداخت","پورسانت","سهم سالن","یادداشت"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h); cell.font = HDR_FONT; cell.fill = PINK
    for r, t in enumerate(transactions, 2):
        ws.cell(row=r,column=1,value=t["date"]); ws.cell(row=r,column=2,value=t["customer"])
        ws.cell(row=r,column=3,value=t["service"]); ws.cell(row=r,column=4,value=t["category"])
        ws.cell(row=r,column=5,value=t["employee"]); ws.cell(row=r,column=6,value=t["amount"])
        ws.cell(row=r,column=7,value=t.get("discount",0)); ws.cell(row=r,column=8,value=t.get("final_amount",t["amount"]))
        ws.cell(row=r,column=9,value=t.get("payment_method","نقدی")); ws.cell(row=r,column=10,value=t.get("commission",0))
        ws.cell(row=r,column=11,value=t.get("salon_share",0)); ws.cell(row=r,column=12,value=t.get("note",""))
    for c in range(1,13): ws.column_dimensions[get_column_letter(c)].width = 16
    fp = os.path.join(DATA_DIR, filename); wb.save(fp); return fp

def _file_size(fp):
    s = os.path.getsize(fp)
    return f"{s//1024} KB" if s < 1048576 else f"{s//1048576} MB"

# ─── Routes ───
@app.route("/")
def index():
    return render_template("index.html", today=PersianDate.today_str(),
        services=get_services(), employees=get_employees())

@app.route("/submit_transaction", methods=["POST"])
def submit_transaction():
    data = request.form
    customer = data.get("customer_name","").strip() or data.get("customer_name_manual","").strip()
    phone = data.get("customer_phone","").strip()
    services = data.getlist("service[]")
    employees = data.getlist("employee[]")
    amounts = data.getlist("amount[]")
    discount_type = data.get("discount_type", "amount")  # "amount" or "percent"
    discount_value = int(data.get("discount_value", 0) or 0)
    payment_method = data.get("payment_method", "نقدی")
    send_sms = data.get("send_sms", "")
    use_points = data.get("use_points", "")
    
    if not customer:
        flash("لطفاً نام مشتری را وارد کنید", "error"); return redirect(url_for("index"))
    if not services:
        flash("حداقل یک خدمت اضافه کنید", "error"); return redirect(url_for("index"))
    
    today = PersianDate.today_str()
    emps = get_employees()
    emp_shares = {e["name"]: e["share_percent"] for e in emps}
    
    # Calculate subtotal
    subtotal = 0
    items_data = []
    for svc, emp, amt in zip(services, employees, amounts):
        if not amt or not amt.isdigit(): continue
        amount = int(amt)
        subtotal += amount
        emp_name = emp.split(" (")[0] if " (" in emp else emp
        commission = int(amount * emp_shares.get(emp_name, 0) / 100)
        salon_share = amount - commission
        items_data.append({
            "service": svc.split(" - ")[0],
            "employee": emp_name,
            "amount": amount,
            "commission": commission,
            "salon_share": salon_share
        })
    
    # Calculate discount
    if discount_type == "percent":
        discount_amount = int(subtotal * discount_value / 100)
    else:
        discount_amount = discount_value
    
    final_amount = max(0, subtotal - discount_amount)
    
    # Use points if selected
    points_used = 0
    if use_points == "on" and customer:
        custs = get_customers()
        cust = next((c for c in custs if c["name"] == customer), None)
        if cust and cust.get("points", 0) > 0:
            points_used = min(cust["points"], final_amount // 1000)  # Each 1000 tomans = 1 point
            final_amount = max(0, final_amount - points_used * 1000)
            cust["points"] -= points_used
            save_customers(custs)
    
    # Add transactions
    for item in items_data:
        add_transaction(
            today, customer, item["service"], "", item["employee"],
            item["amount"], discount_amount // len(items_data) if items_data else 0,
            final_amount // len(items_data) if items_data else 0,
            payment_method, item["commission"], item["salon_share"], phone
        )
    
    # Update customer visit count and points
    if customer:
        custs = get_customers()
        cust = next((c for c in custs if c["name"] == customer), None)
        if cust:
            cust["visit_count"] += 1
            cust["points"] = cust.get("points", 0) + (final_amount // 10000)  # 1 point per 10000 tomans
            save_customers(custs)
        else:
            # Auto-add new customer
            custs.append({
                "name": customer, "phone": phone, "specialty": "",
                "join_date": today, "visit_count": 1, "note": "",
                "points": final_amount // 10000
            })
            save_customers(custs)
    
    # Handle credit/debt
    if payment_method == "نسیه":
        flash(f"⚠️ فاکتور نسیه ثبت شد! مشتری: {customer} — مبلغ: {final_amount:,} تومان", "info")
    else:
        flash(f"✅ تراکنش ثبت شد! مشتری: {customer} — مبلغ نهایی: {final_amount:,} تومان", "success")
    
    if send_sms == "on":
        flash(f"📱 پیامک فاکتور برای {customer} ارسال خواهد شد (قابلیت در دست ساخت)", "info")
    
    return redirect(url_for("index"))

@app.route("/reports")
def reports():
    today = PersianDate.today_str()
    txns = get_transactions(start_date=today, end_date=today)
    emps = get_employees()
    emp_shares = {e["name"]: e["share_percent"] for e in emps}
    total = sum(t.get("final_amount", t["amount"]) for t in txns)
    customers = len(set(t["customer"] for t in txns))
    commission = sum(t.get("commission", 0) for t in txns)
    for t in txns:
        t["commission"] = t.get("commission", int(t["amount"]*emp_shares.get(t["employee"],0)/100))
    return render_template("reports.html", transactions=list(reversed(txns)),
        total=total, customers=customers, services_count=len(txns),
        commission=commission, today=today)

@app.route("/monthly", methods=["GET","POST"])
def monthly():
    now = datetime.now()
    month_str = request.args.get("month", f"{now.year}/{now.month:02d}")
    try:
        parts = month_str.split("/"); y, m = int(parts[0]), int(parts[1])
    except: y, m = now.year, now.month; month_str = f"{y}/{m:02d}"
    start = f"{y}/{m:02d}/01"
    end = f"{y+1}/01/01" if m == 12 else f"{y}/{m+1:02d}/01"
    all_txns = get_transactions()
    monthly_txns = [t for t in all_txns if start <= t["date"] < end]
    emp_shares = {e["name"]: e["share_percent"] for e in get_employees()}
    emp_totals = defaultdict(lambda: {"count":0,"amount":0,"commission":0})
    for t in monthly_txns:
        c = t.get("commission", int(t["amount"]*emp_shares.get(t["employee"],0)/100))
        emp_totals[t["employee"]]["count"] += 1
        emp_totals[t["employee"]]["amount"] += t.get("final_amount", t["amount"])
        emp_totals[t["employee"]]["commission"] += c
    emp_cat = defaultdict(lambda: defaultdict(lambda: {"count":0,"amount":0,"commission":0}))
    for t in monthly_txns:
        c = t.get("commission", int(t["amount"]*emp_shares.get(t["employee"],0)/100))
        emp_cat[t["employee"]][t["category"]]["count"] += 1
        emp_cat[t["employee"]][t["category"]]["amount"] += t.get("final_amount", t["amount"])
        emp_cat[t["employee"]][t["category"]]["commission"] += c
    details = []
    for en in sorted(emp_cat.keys()):
        for cn in sorted(emp_cat[en].keys()):
            d = emp_cat[en][cn]
            details.append({"employee":en,"category":cn,"count":d["count"],"amount":d["amount"],"commission":d["commission"]})
    grand_total = sum(t.get("final_amount", t["amount"]) for t in monthly_txns)
    grand_comm = sum(t.get("commission", int(t["amount"]*emp_shares.get(t["employee"],0)/100)) for t in monthly_txns)
    months = [f"{yy}/{mm:02d}" for yy in range(1400,1410) for mm in range(1,13)]
    return render_template("monthly.html", month_str=month_str, months=months,
        emp_totals=dict(emp_totals), details=details,
        grand_total=grand_total, grand_comm=grand_comm, total_count=len(monthly_txns))

@app.route("/monthly/export/<month_str>")
def export_monthly(month_str):
    try:
        parts = month_str.split("/"); y, m = int(parts[0]), int(parts[1])
    except: flash("خطا", "error"); return redirect(url_for("monthly"))
    start = f"{y}/{m:02d}/01"
    end = f"{y+1}/01/01" if m == 12 else f"{y}/{m+1:02d}/01"
    monthly_txns = [t for t in get_transactions() if start <= t["date"] < end]
    if not monthly_txns: flash("تراکنشی وجود ندارد","error"); return redirect(url_for("monthly"))
    emp_shares = {e["name"]: e["share_percent"] for e in get_employees()}
    wb = Workbook(); ws = wb.active; ws.title = f"گزارش {month_str}"
    headers = ["تاریخ","مشتری","خدمت","دسته","کارمند","مبلغ","پورسانت","یادداشت"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h); cell.font = HDR_FONT; cell.fill = PINK
    for r, t in enumerate(monthly_txns, 2):
        ws.cell(row=r,column=1,value=t["date"]); ws.cell(row=r,column=2,value=t["customer"])
        ws.cell(row=r,column=3,value=t["service"]); ws.cell(row=r,column=4,value=t["category"])
        ws.cell(row=r,column=5,value=t["employee"]); ws.cell(row=r,column=6,value=t.get("final_amount",t["amount"]))
        ws.cell(row=r,column=7,value=t.get("commission",int(t["amount"]*emp_shares.get(t["employee"],0)/100)))
        ws.cell(row=r,column=8,value=t.get("note",""))
    for c in range(1,9): ws.column_dimensions[get_column_letter(c)].width = 18
    fp = os.path.join(DATA_DIR, f"report_{y}_{m:02d}.xlsx"); wb.save(fp)
    return send_file(fp, as_attachment=True)

@app.route("/customers", methods=["GET","POST"])
def customers_page():
    if request.method == "POST":
        action = request.form.get("action")
        custs = get_customers()
        if action == "add":
            custs.append({"name":request.form["name"],"phone":request.form.get("phone",""),"specialty":request.form.get("specialty",""),"join_date":PersianDate.today_str(),"visit_count":int(request.form.get("visits","0") or 0),"note":request.form.get("note",""),"points":0})
        elif action == "edit":
            idx = int(request.form["idx"])
            custs[idx] = {"name":request.form["name"],"phone":request.form.get("phone",""),"specialty":request.form.get("specialty",""),"join_date":custs[idx]["join_date"],"visit_count":int(request.form.get("visits","0") or 0),"note":request.form.get("note",""),"points":custs[idx].get("points",0)}
        elif action == "delete":
            custs.pop(int(request.form["idx"]))
        save_customers(custs)
        return redirect(url_for("customers_page"))
    query = request.args.get("q","").strip().lower()
    custs = get_customers()
    if query:
        custs = [c for c in custs if query in c["name"].lower() or query in c["phone"] or query in c["specialty"].lower()]
    return render_template("customers.html", customers=custs, query=query)

@app.route("/customers/export")
def export_customers():
    custs = get_customers()
    wb = Workbook(); ws = wb.active; ws.title = "دفترچه مشتریان"
    headers = ["نام","تلفن","تخصص مورد علاقه","تاریخ عضویت","تعداد مراجعه","یادداشت","امتیاز"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h); cell.font = HDR_FONT; cell.fill = PINK
    for r, cu in enumerate(custs, 2):
        ws.cell(row=r,column=1,value=cu["name"]); ws.cell(row=r,column=2,value=cu["phone"])
        ws.cell(row=r,column=3,value=cu["specialty"]); ws.cell(row=r,column=4,value=cu["join_date"])
        ws.cell(row=r,column=5,value=cu["visit_count"]); ws.cell(row=r,column=6,value=cu["note"])
        ws.cell(row=r,column=7,value=cu.get("points",0))
    for c in range(1,8): ws.column_dimensions[get_column_letter(c)].width = 18
    fp = os.path.join(DATA_DIR, f"customers_{PersianDate.today_str().replace('/','-')}.xlsx")
    wb.save(fp)
    return send_file(fp, as_attachment=True)

@app.route("/employees", methods=["GET","POST"])
def employees_page():
    if request.method == "POST":
        action = request.form.get("action")
        emps = get_employees()
        if action == "add":
            emps.append({"name":request.form["name"],"specialty":request.form.get("specialty",""),"phone":request.form.get("phone",""),"share_percent":float(request.form.get("share","0") or 0)})
        elif action == "edit":
            emps[int(request.form["idx"])] = {"name":request.form["name"],"specialty":request.form.get("specialty",""),"phone":request.form.get("phone",""),"share_percent":float(request.form.get("share","0") or 0)}
        elif action == "delete":
            emps.pop(int(request.form["idx"]))
        save_employees(emps)
        return redirect(url_for("employees_page"))
    return render_template("employees.html", employees=get_employees())

@app.route("/services", methods=["GET","POST"])
def services_page():
    if request.method == "POST":
        action = request.form.get("action")
        svcs = get_services()
        if action == "add":
            svcs.append({"name":request.form["name"],"category":request.form.get("category",""),"default_price":int(request.form.get("price","0") or 0)})
        elif action == "edit":
            svcs[int(request.form["idx"])] = {"name":request.form["name"],"category":request.form.get("category",""),"default_price":int(request.form.get("price","0") or 0)}
        elif action == "delete":
            svcs.pop(int(request.form["idx"]))
        save_services(svcs)
        return redirect(url_for("services_page"))
    return render_template("services.html", services=get_services())

@app.route("/backup", methods=["GET","POST"])
def backup_page():
    if request.method == "POST":
        action = request.form.get("action")
        today = PersianDate.today_str()
        all_txns = get_transactions()
        if action == "daily":
            txns = get_transactions(start_date=today, end_date=today)
            if txns:
                fp = create_backup_excel(txns, f"backup_daily_{today.replace('/','-')}.xlsx")
                flash(f"✅ بکاپ روزانه ذخیره شد! ({len(txns)} تراکنش)","success")
            else: flash("تراکنشی برای امروز نیست","info")
        elif action == "weekly":
            now = datetime.now(); week_ago = now - timedelta(days=7)
            s_jy,s_jm,s_jd = PersianDate.gregorian_to_jalali(week_ago.year,week_ago.month,week_ago.day)
            e_jy,e_jm,e_jd = PersianDate.gregorian_to_jalali(now.year,now.month,now.day)
            sd = f"{s_jy}/{s_jm:02d}/{s_jd:02d}"; ed = f"{e_jy}/{e_jm:02d}/{e_jd:02d}"
            txns = [t for t in all_txns if sd <= t["date"] <= ed]
            if txns:
                fp = create_backup_excel(txns, f"backup_weekly_{ed.replace('/','-')}.xlsx")
                flash(f"✅ بکاپ هفتگی! ({len(txns)} تراکنش)","success")
            else: flash("تراکنشی در ۷ روز اخیر نیست","info")
        elif action == "monthly":
            now = datetime.now(); jy,jm,jd = PersianDate.gregorian_to_jalali(now.year,now.month,now.day)
            sd = f"{jy}/{jm:02d}/01"; ed = f"{jy}/{jm+1:02d}/01" if jm<12 else f"{jy+1}/01/01"
            txns = [t for t in all_txns if sd <= t["date"] < ed]
            if txns:
                fp = create_backup_excel(txns, f"backup_monthly_{jy}_{jm:02d}.xlsx")
                flash(f"✅ بکاپ ماهانه! ({len(txns)} تراکنش)","success")
            else: flash("تراکنشی برای این ماه نیست","info")
        elif action == "full":
            if all_txns:
                fp = create_backup_excel(all_txns, f"backup_full_{today.replace('/','-')}.xlsx")
                flash(f"✅ بکاپ کامل! ({len(all_txns)} تراکنش)","success")
            else: flash("هیچ تراکنشی وجود ندارد","info")
        elif action == "custom":
            fd = request.form.get("from_date",""); td = request.form.get("to_date","")
            txns = [t for t in all_txns if fd <= t["date"] <= td]
            if txns:
                fn = f"backup_custom_{fd.replace('/','-')}_to_{td.replace('/','-')}.xlsx"
                fp = create_backup_excel(txns, fn)
                flash(f"✅ بکاپ ذخیره شد! ({len(txns)} تراکنش)","success")
            else: flash("تراکنشی در این بازه نیست","info")
        return redirect(url_for("backup_page"))
    backups = []
    if os.path.exists(DATA_DIR):
        for f in sorted(os.listdir(DATA_DIR), reverse=True):
            if f.startswith("backup_") and f.endswith(".xlsx"):
                fp = os.path.join(DATA_DIR, f)
                dt = datetime.fromtimestamp(os.path.getmtime(fp))
                jy,jm,jd = PersianDate.gregorian_to_jalali(dt.year,dt.month,dt.day)
                backups.append({"filename":f,"date":f"{jy}/{jm:02d}/{jd:02d} {dt.strftime('%H:%M')}","size":_file_size(fp)})
    return render_template("backup.html", backups=backups[:20], today=PersianDate.today_str())

@app.route("/download/<filename>")
def download(filename):
    fp = os.path.join(DATA_DIR, filename)
    if os.path.exists(fp): return send_file(fp, as_attachment=True)
    flash("فایل یافت نشد","error"); return redirect(url_for("backup_page"))

# ─── API ───
@app.route("/api/services")
def api_services():
    return jsonify(get_services())

@app.route("/api/employees")
def api_employees():
    return jsonify(get_employees())

@app.route("/api/customers")
def api_customers():
    return jsonify(get_customers())

@app.route("/api/calculate_commission", methods=["POST"])
def api_calculate_commission():
    data = request.json
    emp_name = data.get("employee", "")
    amount = int(data.get("amount", 0))
    emps = get_employees()
    emp = next((e for e in emps if e["name"] == emp_name), None)
    if emp:
        commission = int(amount * emp["share_percent"] / 100)
        salon_share = amount - commission
        return jsonify({"commission": commission, "salon_share": salon_share, "percent": emp["share_percent"]})
    return jsonify({"commission": 0, "salon_share": amount, "percent": 0})

# ─── Run ───
if __name__ == "__main__":
    init_all()
    app.run(host="0.0.0.0", port=5000, debug=True)
