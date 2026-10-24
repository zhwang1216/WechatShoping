import json
import unittest

from app.api_response import ApiResponse


class TestApiResponse(unittest.TestCase):
    def test_ok_format(self):
        resp = ApiResponse.ok({"a": 1})
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["message"], "success")
        self.assertEqual(resp["data"], {"a": 1})

    def test_ok_custom_message(self):
        resp = ApiResponse.ok(None, message="写入成功")
        self.assertEqual(resp["code"], 0)
        self.assertEqual(resp["message"], "写入成功")
        self.assertIsNone(resp["data"])

    def test_error_format(self):
        resp = ApiResponse.error("参数错误", code=400)
        self.assertEqual(resp["code"], 400)
        self.assertEqual(resp["message"], "参数错误")
        self.assertIsNone(resp["data"])

    def test_to_bytes_utf8_chinese(self):
        payload = ApiResponse.to_bytes(ApiResponse.ok({"名称": "微信小店"}))
        text = payload.decode("utf-8")
        self.assertIn("微信小店", text)
        data = json.loads(text)
        self.assertEqual(data["data"]["名称"], "微信小店")


if __name__ == "__main__":
    unittest.main()
