# -*- coding: utf-8 -*-
"""验证：overall.retest_due 与 /api/practice/retest 实际可出题数是否一致。

构造两条已过关的到期记录：
  A. 真实知识点（题库有同知识点题目）→ 应计入且能出题
  B. 虚构知识点（题库无题目）        → 修复后不应计入
运行前备份 mastery.xlsx，运行后还原，保证不留测试数据。
"""
import json
import os
import shutil
import sys
import urllib.request

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, "server"))

import excel_store as store          # noqa: E402
import engine                        # noqa: E402

BASE = "http://127.0.0.1:7100"
MASTERY = os.path.join(BASE_DIR, "server", "data", "mastery.xlsx")
BAK = MASTERY + ".verify_bak"

REAL_KP = "牛顿第二定律"
FAKE_KP = "__测试_题库无此题__"
TODAY = "2020-01-01"   # 远早于今天，必然到期


def api(path, token=None, method="GET", payload=None):
    h = {"Content-Type": "application/json"}
    if token:
        h["X-Auth-Token"] = token
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(BASE + path, data=data, headers=h, method=method)
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode())


def main():
    shutil.copy2(MASTERY, BAK)
    try:
        r = api("/api/login", method="POST",
                payload={"username": "student", "password": "123456"})
        token = r["data"]["token"]

        # 造两条到期记录
        store.upsert_mastery("student", REAL_KP, 88.0, "已过关",
                             last_review=TODAY, review_count=0)
        store.upsert_mastery("student", FAKE_KP, 90.0, "已过关",
                             last_review=TODAY, review_count=0)
        engine.store.upsert_mastery = store.upsert_mastery  # 确保同一份 store

        raw = engine.retest_due("student")
        prod = engine.retest_due("student", only_producible=True)
        print("到期记录（原始）      : %d" % len(raw))
        print("到期记录（可出题过滤）: %d" % len(prod))

        overall = api("/api/analytics/overall", token)["data"]
        retest = api("/api/practice/retest", token)["data"]
        if isinstance(retest, dict):
            retest = retest.get("items") or retest.get("questions") or []
        print("overall.retest_due    : %d" % overall["retest_due"])
        print("接口实际可出题数      : %d" % len(retest))
        for it in retest:
            print("    %s | qid %s" % (it.get("kp"), (it.get("question") or {}).get("qid")))

        ok = overall["retest_due"] == len(retest) and len(raw) > len(prod)
        print("\n结论:", "PASS 计数与实际出题一致" if ok else "FAIL 仍不一致")
        return 0 if ok else 1
    finally:
        shutil.copy2(BAK, MASTERY)
        os.remove(BAK)
        print("已还原 mastery.xlsx（测试数据未保留）")


if __name__ == "__main__":
    sys.exit(main())
