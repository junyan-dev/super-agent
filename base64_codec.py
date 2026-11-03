"""Base64 encoding and decoding helpers.

A small, dependency-free wrapper around the standard library's
:mod:`base64` module.

>>> base64_encode("hello")
'aGVsbG8='
>>> base64_decode("aGVsbG8=")
b'hello'
"""

from __future__ import annotations

import base64

__all__ = ["base64_encode", "base64_decode"]


def base64_encode(data: str | bytes | bytearray | memoryview) -> str:
    """Return the Base64 representation of *data* as an ASCII string.

    Strings are encoded as UTF-8 before being Base64-encoded, so the result
    round-trips through :func:`base64_decode`.  Bytes-like inputs are encoded
    as-is.

    Args:
        data: Text or binary payload to encode.

    Returns:
        The Base64-encoded payload as an ASCII ``str``.

    Raises:
        TypeError: If *data* is neither a string nor bytes-like.
    """
    if isinstance(data, str):
        raw = data.encode("utf-8")
    elif isinstance(data, (bytes, bytearray, memoryview)):
        raw = bytes(data)
    else:
        raise TypeError(
            f"data must be str or bytes-like, not {type(data).__name__}"
        )

    return base64.b64encode(raw).decode("ascii")


def base64_decode(
    data: str | bytes | bytearray | memoryview,
) -> bytes:
    """Decode a Base64 payload and return the raw bytes.

    The input is validated strictly, so characters outside the standard
    Base64 alphabet, impossible lengths and incorrect padding all raise
    :class:`ValueError`.  Decode the result as UTF-8 to recover text that was
    passed to :func:`base64_encode`.

    Args:
        data: Base64 payload to decode.

    Returns:
        The decoded bytes.

    Raises:
        TypeError: If *data* is neither a string nor bytes-like.
        ValueError: If *data* is not valid Base64.
    """
    if isinstance(data, str):
        raw = data.encode("ascii", errors="strict")
    elif isinstance(data, (bytes, bytearray, memoryview)):
        raw = bytes(data)
    else:
        raise TypeError(
            f"data must be str or bytes-like, not {type(data).__name__}"
        )

    return base64.b64decode(raw, validate=True)
