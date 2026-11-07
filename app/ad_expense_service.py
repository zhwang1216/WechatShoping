"""月度广告花费管理服务

定位（spec 5.x 广告费用管理页）：
1. 广告花费按月份手填；产品成本/快递费/综合成本/产品毛利/经营贡献利润等一律只读。
2. 按月份（YYYY-MM）登记"广告花费总金额"，用于利润页
   "扣除广告费后经营利润 = 经营贡献利润 − 当月广告花费总金额"。
3. 金额必须以字符串形式保存 Decimal 结果，禁止浮点直接计算。
4. 每次修改保留历史（history），并写入操作日志，可追溯。
5. 广告花费独立核算，不得混入供应商综合成本；利润计算仅扣一次。
"""
import re
from decimal import Decimal, InvalidOperation

from .data_schemas import now_iso
from .errors import ApiError

MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")


class AdExpenseService:
    def __init__(self, repo, op_log):
        self.repo = repo
        self.op_log = op_log

    # ==================== 查询 ====================
    def list(self, month=None):
        records = self.repo.read("ad_expenses")["records"]
        if month:
            records = [r for r in records if r.get("month") == month]
        records = sorted(records, key=lambda r: r.get("month") or "")
        return {
            "records": records,
            "month": month,
            "total": len(records),
            "amount_sum": self._sum_of(records),
        }

    def get_amount(self, month):
        """返回指定月份广告花费总金额（未登记按 0 计），Decimal"""
        if not month:
            return Decimal("0")
        for r in self.repo.read("ad_expenses")["records"]:
            if r.get("month") == month:
                return Decimal(r.get("amount") or "0")
        return Decimal("0")

    def months_amounts(self, months):
        """批量返回 {month: Decimal}，未登记月份为 0"""
        result = {m: Decimal("0") for m in months}
        for r in self.repo.read("ad_expenses")["records"]:
            m = r.get("month")
            if m in result:
                result[m] = Decimal(r.get("amount") or "0")
        return result

    @staticmethod
    def _sum_of(records):
        total = Decimal("0")
        for r in records:
            try:
                total += Decimal(r.get("amount") or "0")
            except InvalidOperation:
                continue
        return str(total.quantize(Decimal("0.01")))

    # ==================== 保存（唯一可编辑财务金额） ====================
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
            raise ApiError("广告花费总金额必须是非负数字", code=400, http_status=400)
        amount_s = str(amount_d.quantize(Decimal("0.01")))
        if note is not None and not isinstance(note, str):
            raise ApiError("note 必须是字符串", code=400, http_status=400)
        note = (note or "").strip() or None

        now = now_iso()
        old = None

        def mutator(d):
            nonlocal old
            rec = next((r for r in d["records"]
                        if r.get("month") == month), None)
            if rec is None:
                d["records"].append({
                    "record_id": "adexp_" + re.sub(r"\D", "", month)
                    + "_" + format(len(d["records"]) + 1, "04d"),
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
                rec["history"].append({
                    "amount": rec.get("amount"),
                    "note": rec.get("note"),
                    "changed_at": now,
                    "changed_by": operator,
                })
                rec["amount"] = amount_s
                rec["note"] = note
                rec["updated_at"] = now
                rec["updated_by"] = operator
            d["updated_at"] = now

        self.repo.update("ad_expenses", mutator)
        self.op_log.append("ad_expense.save", detail={
            "month": month, "amount": amount_s,
            "previous_amount": (old or {}).get("amount"),
        }, operator=operator)
        return self.list(month)
