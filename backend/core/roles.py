"""core.roles — SATU konstanta peran keuangan/approver (T-08).

Peran yang benar-benar di-seed di aplikasi ini: `accounting`, `staff_keuangan`,
`manager_keuangan` (bukan `finance`). Nama generik tetap dipertahankan untuk
kompatibilitas data lama.
"""
FINANCE_ROLES = ("superadmin", "admin", "owner", "accounting", "staff_keuangan",
                 "manager_keuangan", "finance", "finance_manager", "accountant")
APPROVER_ROLES = FINANCE_ROLES + ("hr", "hr_manager", "manager")
