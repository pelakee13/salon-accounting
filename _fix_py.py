p = "/data/data/com/termux/files/home/salon-accounting/web_app.py"
s = open(p, encoding="utf-8").read()
marker = "# ─── Data Access: Services ───"
start = s.index("def employee_payroll(")
next_def = s.index("\n" + marker, start)
func = s[start:next_def]
first_ret = func.index("return result")
new_func = func[:first_ret + len("return result")] + "\n"
s2 = s[:start] + new_func + s[next_def:]
open(p, "w", encoding="utf-8").write(s2)
print("OK, removed dead body. func len:", len(new_func))
