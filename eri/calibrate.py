"""회귀식 계수 재보정.

특허의 계수는 특허가 사용한 A/B/C 어휘 목록으로 추정한 값이다. 다른 어휘 목록
(예: 국립국어원 1~5등급)을 쓰면 X1·X2·Z 값의 분포가 달라지므로, 학년이 알려진
지문(교과서 지문 등)으로 같은 형태의 회귀식(상호작용항 포함)을 다시 추정할 수 있다.

    초등: 학년 = b0 + b1·Y + b2·X1 + b3·(X1·Y)
    중등: 학년 = b0 + b1·X2 + b2·Z + b3·K + b4·(K·Z)
"""
from __future__ import annotations

import numpy as np


def fit(results, grades):
    """분석 결과와 실제 학년 목록 → {'초등': (coef, R², n), '중등': ...}"""
    out = {}
    for level in ("초등", "중등"):
        rows = [(r, g) for r, g in zip(results, grades) if r.level == level and g is not None]
        need = 4 if level == "초등" else 5
        if len(rows) < need + 2:
            out[level] = (None, None, len(rows))
            continue
        if level == "초등":
            X = np.array([[1, r.Y, r.X1, r.X1 * r.Y] for r, _ in rows], float)
            names = ("intercept", "Y", "X1", "X1Y")
        else:
            X = np.array([[1, r.X2, r.Z, r.K, r.K * r.Z] for r, _ in rows], float)
            names = ("intercept", "X2", "Z", "K", "KZ")
        y = np.array([g for _, g in rows], float)
        beta, *_ = np.linalg.lstsq(X, y, rcond=None)
        pred = X @ beta
        ss_res = float(((y - pred) ** 2).sum())
        ss_tot = float(((y - y.mean()) ** 2).sum())
        r2 = 1 - ss_res / ss_tot if ss_tot else None
        out[level] = ({k: round(float(v), 4) for k, v in zip(names, beta)}, r2, len(rows))
    return out
