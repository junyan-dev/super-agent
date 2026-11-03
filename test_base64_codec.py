"""Tests for :mod:`base64_codec`."""

import unittest

from base64_codec import base64_decode, base64_encode


class Base64EncodeTests(unittest.TestCase):
    def test_encode_text(self) -> None:
        self.assertEqual(base64_encode("hello"), "aGVsbG8=")

    def test_encode_unicode_text(self) -> None:
        self.assertEqual(base64_encode("你好"), "5L2g5aW9")

    def test_encode_bytes(self) -> None:
        self.assertEqual(base64_encode(b"\x00\xff"), "AP8=")

    def test_encode_empty(self) -> None:
        self.assertEqual(base64_encode(""), "")
        self.assertEqual(base64_encode(b""), "")

    def test_encode_rejects_unsupported_type(self) -> None:
        with self.assertRaises(TypeError):
            base64_encode(123)  # type: ignore[arg-type]


class Base64DecodeTests(unittest.TestCase):
    def test_decode_text_payload(self) -> None:
        self.assertEqual(base64_decode("aGVsbG8="), b"hello")

    def test_decode_unicode_payload(self) -> None:
        self.assertEqual(base64_decode("5L2g5aW9"), "你好".encode("utf-8"))

    def test_decode_accepts_bytes(self) -> None:
        self.assertEqual(base64_decode(b"aGVsbG8="), b"hello")

    def test_round_trip(self) -> None:
        payload = "round-trip 数据 \x00"
        self.assertEqual(
            base64_decode(base64_encode(payload)), payload.encode("utf-8")
        )

    def test_decode_rejects_invalid_input(self) -> None:
        for invalid in ("not base64!", "abc", "aGVsbG8"):
            with self.subTest(invalid=invalid):
                with self.assertRaises(ValueError):
                    base64_decode(invalid)

    def test_decode_rejects_unsupported_type(self) -> None:
        with self.assertRaises(TypeError):
            base64_decode(None)  # type: ignore[arg-type]


if __name__ == "__main__":
    unittest.main()
