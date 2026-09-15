#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 wps-split 目录发布到 GitHub（不依赖本机 git，直接走 GitHub REST API）。

用法：
  python3 tools/publish_to_github.py                # 默认 leifung/wps-split
  python3 tools/publish_to_github.py --repo 别的名字
  python3 tools/publish_to_github.py --private      # 建私有仓库
  python3 tools/publish_to_github.py --dry-run      # 只看要传哪些文件

令牌来源（按优先级）：
  1. 环境变量 GH_TOKEN
  2. 本机已保存的 ~/.workbuddy/.github_token
"""
import argparse
import base64
import json
import os
import sys
import time
import urllib.error
import urllib.request

API = "https://api.github.com"
SKIP_DIRS = {".git", "__pycache__", ".workbuddy", ".idea", ".vscode", "build"}
SKIP_FILES = {".DS_Store", "Thumbs.db"}
SKIP_REL_PREFIX = ()                      # 额外排除的相对路径前缀
SKIP_REL_CONTAINS = ("拆分结果",)           # 拆分产物目录不入库

_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
urllib.request.install_opener(_opener)


def read_saved_token():
    p = os.path.join(os.path.expanduser("~"), ".workbuddy", ".github_token")
    try:
        if os.path.isfile(p):
            return open(p, encoding="utf-8").read().strip()
    except Exception:
        pass
    return ""


def req(method, path, data=None, token=""):
    url = path if path.startswith("http") else API + path
    body = json.dumps(data, ensure_ascii=False).encode("utf-8") \
        if data is not None else None
    r = urllib.request.Request(url, data=body, method=method)
    r.add_header("Authorization", "Bearer " + token)
    r.add_header("Accept", "application/vnd.github+json")
    r.add_header("X-GitHub-Api-Version", "2022-11-28")
    r.add_header("User-Agent", "wps-split-publish")
    if body:
        r.add_header("Content-Type", "application/json; charset=utf-8")
    try:
        with urllib.request.urlopen(r, timeout=180) as resp:
            payload = resp.read()
            return json.loads(payload.decode("utf-8")) if payload else {}
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:600]
        raise RuntimeError("HTTP %s %s\n%s" % (e.code, e.reason, detail))


def collect(root):
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fn in sorted(filenames):
            if fn in SKIP_FILES or fn.endswith(".pyc"):
                continue
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, root).replace("\\", "/")
            if rel.startswith(SKIP_REL_PREFIX):
                continue
            if any(s in rel for s in SKIP_REL_CONTAINS):
                continue
            mode = "100755" if fn.endswith(".sh") else "100644"
            out.append((rel, mode, full))
    return sorted(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--owner", default="leifung")
    ap.add_argument("--repo", default="wps-split")
    ap.add_argument("--private", action="store_true")
    ap.add_argument("--message", default="")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--token-file", default="")
    args = ap.parse_args()

    token = os.environ.get("GH_TOKEN", "").strip()
    if args.token_file and os.path.isfile(args.token_file):
        token = open(args.token_file, encoding="utf-8").read().strip()
    if not token:
        token = read_saved_token()
    if not token and not args.dry_run:
        print("缺少 GitHub 令牌：")
        print("  https://github.com/settings/tokens 生成 classic token（勾 repo）")
        print("  然后 export GH_TOKEN=ghp_xxx，或存到 ~/.workbuddy/.github_token")
        return 2

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files = collect(root)
    print("待上传 %d 个文件：" % len(files))
    for rel, mode, full in files:
        print("  %s  %-40s %8d B" % (mode, rel, os.path.getsize(full)))
    if args.dry_run:
        return 0

    full_name = "%s/%s" % (args.owner, args.repo)

    # 1) 仓库：存在就用，不存在就建
    try:
        repo = req("GET", "/repos/" + full_name, token=token)
        print("仓库已存在：%s" % repo["html_url"])
    except Exception:
        repo = req("POST", "/user/repos", token=token, data={
            "name": args.repo,
            "description": "统信 UOS 下按指定列拆分 WPS 表格：新表以分类名命名，并完整沿用原表模板格式。零依赖，仅用 Python 标准库。",
            "homepage": "https://github.com/leifung/wps-split",
            "private": bool(args.private),
            "auto_init": True,          # 先生成初始提交，main 分支才存在
            "has_issues": True,
        })
        print("已创建仓库：%s" % repo["html_url"])
        time.sleep(2)

    # 2) 拿 main 分支最新提交
    base = None
    for _ in range(10):
        try:
            ref = req("GET", "/repos/%s/git/ref/heads/main" % full_name, token=token)
            base = ref["object"]["sha"]
            break
        except Exception:
            time.sleep(2)
    if base is None:
        raise RuntimeError("拿不到 main 分支，去网页确认仓库是否创建成功")

    # 3) 逐文件建 blob
    tree = []
    for rel, mode, full in files:
        with open(full, "rb") as f:
            content = f.read()
        blob = req("POST", "/repos/%s/git/blobs" % full_name, token=token, data={
            "content": base64.b64encode(content).decode("ascii"),
            "encoding": "base64",
        })
        tree.append({"path": rel, "mode": mode, "type": "blob", "sha": blob["sha"]})
        print("  blob %-40s %s" % (rel, blob["sha"][:8]))

    new_tree = req("POST", "/repos/%s/git/trees" % full_name,
                   token=token, data={"tree": tree})
    msg = args.message or """feat: 按列拆分 WPS 表格（保留原表模板格式，零依赖）

- 模板克隆引擎：复制原文件全部样式部件，只重写数据行
  字体/边框/底色/行高/列宽/数字格式/冻结/筛选/合并/条件格式/数据验证全部保留
- 公式行号引用自动改写（=D12*E12 -> =D5*E5）
- 支持列名 / 列字母 / 列序号定位，多列组合拆分
- 表头行可指定，表头上方多行标题可用 --template-rows 原样保留
- 文件名安全化（非法字符清理、超长截断、重名加 _2），清单 CSV 带 BOM
- 三种用法：tkinter 图形界面 / 交互式向导 / 命令行（支持批量目录）
- 纯 Python 标准库实现，不依赖 pandas/openpyxl，内网离线可用"""
    commit = req("POST", "/repos/%s/git/commits" % full_name, token=token, data={
        "message": msg, "tree": new_tree["sha"], "parents": [base]})
    req("PATCH", "/repos/%s/git/refs/heads/main" % full_name,
        token=token, data={"sha": commit["sha"]})
    print("已更新 main -> %s" % commit["sha"][:8])

    try:
        req("PUT", "/repos/%s/topics" % full_name, token=token, data={
            "names": ["uos", "deepin", "linux", "wps", "excel", "xlsx",
                      "spreadsheet", "split", "python", "zero-dependency"]})
        print("已设置 topics")
    except Exception as e:
        print("topics 设置跳过：%s" % e)

    print("\n完成： https://github.com/%s" % full_name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
