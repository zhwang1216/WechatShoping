"""统一 API 响应格式

所有接口返回:
    {"code": 0, "message": "success", "data": ...}   成功
    {"code": 非0, "message": "错误说明", "data": null} 失败
"""
import json


class ApiResponse:
    CODE_OK = 0

    @staticmethod
    def ok(data=None, message="success"):
        return {"code": ApiResponse.CODE_OK, "message": message, "data": data}

    @staticmethod
    def error(message, code=1, data=None):
        return {"code": code, "message": message, "data": data}

    @staticmethod
    def to_bytes(resp):
        return json.dumps(resp, ensure_ascii=False, indent=2).encode("utf-8")
