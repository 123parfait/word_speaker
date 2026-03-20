# -*- coding: utf-8 -*-
from __future__ import annotations

import difflib


def build_diff_view(expected, actual):
    expected = str(expected or "")
    actual = str(actual or "")
    matcher = difflib.SequenceMatcher(a=expected.lower(), b=actual.lower())
    opcodes = matcher.get_opcodes()
    return {
        "expected_segments": _build_expected_segments(expected, opcodes),
        "actual_segments": _build_actual_segments(actual, opcodes),
        "legend": [
            ("missing", "missing"),
            (" / ", None),
            ("extra", "extra"),
            (" / ", None),
            ("wrong", "wrong"),
        ],
    }


def _build_expected_segments(expected, opcodes):
    segments = []
    for tag, i1, i2, _j1, _j2 in opcodes:
        if tag == "insert":
            continue
        segment = expected[i1:i2]
        style = None
        if tag == "replace":
            style = "wrong"
        elif tag == "delete":
            style = "missing"
        segments.append((segment, style))
    return segments


def _build_actual_segments(actual, opcodes):
    segments = []
    for tag, _i1, _i2, j1, j2 in opcodes:
        if tag == "delete":
            continue
        segment = actual[j1:j2]
        style = None
        if tag == "replace":
            style = "wrong"
        elif tag == "insert":
            style = "extra"
        segments.append((segment, style))
    return segments
