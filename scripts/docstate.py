# -*- coding: utf-8 -*-
"""记录与比对文件状态，发现用户在两轮之间手改过。
用法: python docstate.py record <file>   # 本轮写入后
      python docstate.py check  <file>   # 下轮动手前
状态存在文件同目录的 .tmp/.docstate.json。"""
import hashlib
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")


def state_path(f):
    d = os.path.join(os.path.dirname(os.path.abspath(f)), ".tmp")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, ".docstate.json")


def digest(f):
    with open(f, "rb") as fh:
        return {"mtime": os.path.getmtime(f), "sha256": hashlib.sha256(fh.read()).hexdigest()}


def main():
    mode, f = sys.argv[1], os.path.abspath(sys.argv[2])
    sp = state_path(f)
    db = json.load(open(sp, encoding="utf-8")) if os.path.exists(sp) else {}
    cur = digest(f)
    if mode == "record":
        db[f] = cur
        json.dump(db, open(sp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("已记录", f)
        return
    old = db.get(f)
    if old is None:
        print("无记录：首次处理，先完整读一遍")
    elif old["sha256"] == cur["sha256"]:
        print("未变：与上轮写入后一致")
    else:
        print("已变：用户在两轮之间改过。先 dump 比对并向用户复述差异，手改内容不得改回")


if __name__ == "__main__":
    main()
