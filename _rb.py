p = "auth.py"
s = open(p, encoding="utf-8").read()
old = '''    "reception": {"/", "/submit_transaction", "/customers", "/customer/",
                 "/reports"},                       # front-desk only: customers, services, daily report
    "employee": {"/dashboard", "/reports", "/payroll"},  # own commission / today + payroll'''
new = '''    "reception": {"/", "/submit_transaction", "/customers", "/customer/",
                 "/reports", "/transaction/delete", "/transaction/edit"},
    "employee": {"/dashboard", "/reports", "/payroll", "/transaction/delete", "/transaction/edit"},'''
assert old in s
s = s.replace(old, new)
open(p, "w", encoding="utf-8").write(s)
print("auth: transaction routes allowed for reception/employee")
