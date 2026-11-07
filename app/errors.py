"""通用异常与错误码定义"""


class ApiError(Exception):
    """通用 API 参数/权限错误"""

    def __init__(self, message, code=400, http_status=400):
        super().__init__(message)
        self.code = code
        self.http_status = http_status


class ImportServiceError(Exception):
    """导入服务统一异常基类"""

    def __init__(self, message, code=2000, http_status=400, data=None):
        super().__init__(message)
        self.code = code
        self.http_status = http_status
        self.data = data


class ImportValidationError(ImportServiceError):
    """文件/模板/数据校验失败（阻止正式导入）"""

    def __init__(self, message, data=None):
        super().__init__(message, code=2001, http_status=400, data=data)


class ExcelParseError(ImportServiceError):
    """Excel 文件无法解析"""

    def __init__(self, message):
        super().__init__(message, code=2002, http_status=400)


class ImportConflictError(ImportServiceError):
    """重复文件上传冲突"""

    def __init__(self, message, data=None):
        super().__init__(message, code=2003, http_status=409, data=data)


class UploadNotFoundError(ImportServiceError):
    """上传暂存文件不存在或已过期"""

    def __init__(self, message):
        super().__init__(message, code=2004, http_status=404)


class BatchNotFoundError(ImportServiceError):
    """导入批次不存在"""

    def __init__(self, message):
        super().__init__(message, code=2005, http_status=404)


class ReconcileServiceError(Exception):
    """对账服务统一异常"""

    def __init__(self, message, code=3000, http_status=400, data=None):
        super().__init__(message)
        self.code = code
        self.http_status = http_status
        self.data = data


class OrderNotFoundError(ReconcileServiceError):
    """微信订单不存在"""

    def __init__(self, message):
        super().__init__(message, code=3001, http_status=404)


class SupplierRecordNotFoundError(ReconcileServiceError):
    """供应商成本记录不存在"""

    def __init__(self, message):
        super().__init__(message, code=3002, http_status=404)


class ManualMatchConflictError(ReconcileServiceError):
    """人工匹配关系冲突（成本记录已被其他订单占用）"""

    def __init__(self, message, data=None):
        super().__init__(message, code=3003, http_status=409, data=data)


class ModifyRejectedError(ReconcileServiceError):
    """已匹配订单拒绝人工修改"""

    def __init__(self, message, data=None):
        super().__init__(message, code=3004, http_status=403, data=data)


class FieldNotModifiableError(ReconcileServiceError):
    """字段为系统关键字段不可修改"""

    def __init__(self, message, data=None):
        super().__init__(message, code=3005, http_status=400, data=data)


class InvalidFieldValueError(ReconcileServiceError):
    """字段值格式非法"""

    def __init__(self, message, data=None):
        super().__init__(message, code=3006, http_status=400, data=data)


class ProfitServiceError(Exception):
    """利润核算服务统一异常"""

    def __init__(self, message, code=4000, http_status=400, data=None):
        super().__init__(message)
        self.code = code
        self.http_status = http_status
        self.data = data
