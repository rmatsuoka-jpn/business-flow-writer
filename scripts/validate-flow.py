#!/usr/bin/env python3
"""フロー定義JSONの静的検証ツール。

generate-flow.ps1 が描画時に例外を投げる条件（lane がlanesに無い／
flow の from/to が nodes に無い／未知の type など）を Excel を使わずに
事前チェックする。Excel/PowerShell の無い環境でも定義の妥当性を確認できる。

使い方:
    python scripts/validate-flow.py <flow.json のパス>

戻り値: エラーがあれば終了コード 1、無ければ 0。
"""
import json
import sys

# Windows の既定コンソール（日本語環境では cp932）では絵文字・em ダッシュ（—）を
# 出力できず UnicodeEncodeError で落ちるため、出力ストリームを UTF-8 に固定する。
# （Python 3.7+ の TextIOWrapper.reconfigure。PYTHONUTF8=1 が無い 3.13 でも安全に動く）
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

VALID_NODE_TYPES = {"start", "end", "task", "gateway", "datastore",
                    "dataobject", "document", "annotation"}
VALID_SITES = {"top", "left", "bottom", "right"}


def validate(path):
    errs, warns = [], []
    with open(path, encoding="utf-8") as f:
        d = json.load(f)

    # ---- lanes ----
    lane_ids = set()
    for i, ln in enumerate(d.get("lanes", [])):
        for req in ("id", "pool", "name"):
            if req not in ln:
                errs.append(f"lanes[{i}]: 必須フィールド '{req}' が無い")
        if ln.get("id") in lane_ids:
            errs.append(f"lanes[{i}]: id '{ln.get('id')}' が重複")
        lane_ids.add(ln.get("id"))
    if not lane_ids:
        errs.append("lanes が空")

    # ---- nodes （ps1: 未知type=throw / lane不在=throw）----
    node_ids = set()
    nodes_by_id = {}
    for i, n in enumerate(d.get("nodes", [])):
        nid = n.get("id")
        nodes_by_id[nid] = n
        for req in ("id", "type", "lane", "col"):
            if req not in n:
                errs.append(f"nodes[{i}] (id={nid}): 必須フィールド '{req}' が無い")
        if nid in node_ids:
            errs.append(f"nodes[{i}]: id '{nid}' が重複")
        node_ids.add(nid)
        if n.get("type") not in VALID_NODE_TYPES:
            errs.append(f"node '{nid}': 未知のノード種別 '{n.get('type')}'  # 描画時に例外")
        if n.get("lane") not in lane_ids:
            errs.append(f"node '{nid}': lane '{n.get('lane')}' が lanes に無い  # 描画時に例外")
        if n.get("type") in {"task", "datastore", "dataobject", "document", "annotation"}:
            if not n.get("label"):
                warns.append(f"node '{nid}': type={n.get('type')} だが label が空")

    # ---- flows （ps1: from/to が nodes に無い=throw）----
    for i, fl in enumerate(d.get("flows", [])):
        for end in ("from", "to"):
            if fl.get(end) not in node_ids:
                errs.append(f"flows[{i}]: {end} '{fl.get(end)}' が nodes に無い  # 描画時に例外")
        for site in ("fromSite", "toSite"):
            if site in fl and fl[site] not in VALID_SITES:
                errs.append(f"flows[{i}]: {site} '{fl[site]}' は top/left/bottom/right のいずれかでない")
        # 規約: association はデータオブジェクトにつながない（警告）
        if fl.get("type") == "association":
            for label in ("from", "to"):
                node = nodes_by_id.get(fl.get(label))
                if node and node.get("type") in {"dataobject", "document"}:
                    warns.append(
                        f"flows[{i}]: association が dataobject '{node.get('id')}' に接続（規約では非推奨）")

    return errs, warns


def main():
    if len(sys.argv) != 2:
        print("使い方: python scripts/validate-flow.py <flow.json のパス>", file=sys.stderr)
        return 2
    path = sys.argv[1]
    print(f"■ 検証対象: {path}")
    errs, warns = validate(path)
    for w in warns:
        print(f"  ⚠ 警告: {w}")
    for e in errs:
        print(f"  ✗ エラー: {e}")
    if errs:
        print(f"\n結果: NG（エラー {len(errs)} 件 / 警告 {len(warns)} 件）— この定義は描画時に落ちます")
        return 1
    print(f"\n結果: OK（エラー 0 件 / 警告 {len(warns)} 件）— 描画時の例外条件には該当しません")
    return 0


if __name__ == "__main__":
    sys.exit(main())
