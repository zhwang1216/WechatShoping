"""月度其他费用：按店铺、按月份手填，计入每月盈利。

金额以字符串保存 Decimal 结果。同一月份再次保存会覆盖金额和说明，并留下历史。
"""
import re
from decimal import Decimal, InvalidOperation

from .data_schemas import now_iso
from .errors import ApiError

MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")


class OtherExpenseService:
    def __init__(self, repo, op_log):
        self.repo = repo
        self.op_log = op_log

    def list(self, month=None):
        records = self.repo.read("other_expenses")["records"]
        if month:
            records = [r for r in records if r.get("month") == month]
        records = sorted(records, key=lambda r: r.get("month") or "")
        total = Decimal("0")
        for record in records:
            try:
                total += Decimal(record.get("amount") or "0")
            except InvalidOperation:
                continue
        return {
            "records": records,
            "month": month,
            "total": len(records),
            "amount_sum": str(total.quantize(Decimal("0.01"))),
        }

    def save(self, month, amount, note=None, operator="admin"):
        if not isinstance(month, str) or not MONTH_RE.match(month or ""):
            raise ApiError("month 必须是 YYYY-MM 格式（如 2026-06）",
                           code=400, http_status=400)
        if amount is None or isinstance(amount, bool) \
                or isinstance(amount, (list, dict)):
            raise ApiError("amount 必须是非负数字符串或数字",
                           code=400, http_status=400)
        try:
            amount_d = Decimal(str(amount))
        except InvalidOperation:
            raise ApiError("amount 不是合法金额", code=400, http_status=400)
        if amount_d < 0 or not amount_d.is_finite():
            raise ApiError("其他费用必须是非负数字", code=400, http_status=400)
        amount_s = str(amount_d.quantize(Decimal("0.01")))
        if note is not None and not isinstance(note, str):
            raise ApiError("note 必须是字符串", code=400, http_status=400)
        note = (note or "").strip() or None
        if note and len(note) > 40:
            raise ApiError("其他费用说明最多 40 个字", code=400, http_status=400)

        now = now_iso()
        old = None

        def mutator(data):
            nonlocal old
            rec = next((r for r in data["records"]
                        if r.get("month") == month), None)
            if rec is None:
                data["records"].append({
                    "record_id": "othexp_" + re.sub(r"\D", "", month)
                    + "_" + format(len(data["records"]) + 1, "04d"),
                    "month": month,
                    "amount": amount_s,
                    "note": note,
                    "created_at": now,
                    "updated_at": now,
                    "updated_by": operator,
                    "history": [],
                })
            else:
                old = dict(rec)
                rec.setdefault("history", []).append({
                    "amount": rec.get("amount"),
                    "note": rec.get("note"),
                    "changed_at": now,
                    "changed_by": operator,
                })
                rec["amount"] = amount_s
                rec["note"] = note
                rec["updated_at"] = now
                rec["updated_by"] = operator
            data["updated_at"] = now

        self.repo.update("other_expenses", mutator)
        self.op_log.append("other_expense.save", detail={
            "month": month, "amount": amount_s, "note": note,
            "previous_amount": (old or {}).get("amount"),
        }, operator=operator)
        return self.list(month)
