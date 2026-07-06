# business-flow-writer

自治体業務の**業務フロー図（BPMN 簡易セット・分析レベル）**を、ヒアリング内容から Excel (.xlsx) として自動生成する [Claude](https://claude.com) スキルです。デジタル庁「地方公共団体の基幹業務システムの標準仕様における業務フロー」／ J-LIS「機能要件の表記方法利用ガイド」に準拠しています。

図はすべて Excel のオートシェイプで描画されるため、生成後に図形のドラッグ・テキスト編集・線の付け替えで自由に仕上げられます。**8割を自動生成し、最後の見栄えは手作業で仕上げる**運用を前提としています。

## 特徴

- ヒアリング → 中間定義表（Markdown）での確認 → JSON → Excel生成、という段階的ワークフロー
- プール／レーン、作業（作業番号付き）、排他分岐、データストア、データオブジェクト、注記、脚注（機能要件対応表）に対応
- 生成物は編集可能なオートシェイプ。元の `flow.json` を保管すれば再生成も可能
- 完全ローカルで動作（外部送信なし）

## 前提環境

| 依存 | 用途 | 必須 |
|---|---|---|
| Windows | Excel COM の利用 | ✓ |
| Microsoft Excel | オートシェイプ描画（COM経由） | ✓ |
| PowerShell 7+ | 生成スクリプトの実行 | ✓ |
| [Poppler](https://poppler.freedesktop.org/)（`pdftoppm`） | 生成PDFのPNG化による自己検証 | 任意 |

> Excel COM を使うため、本スキルは **Windows + デスクトップ版 Excel** が必要です。macOS / Linux / Excel 未インストール環境では動作しません。

## ディレクトリ構成

```
business-flow-writer/
├── SKILL.md                    # スキル本体（ワークフロー定義）
├── references/
│   └── notation.md             # 記法・JSONスキーマ・レイアウト規約
├── scripts/
│   └── generate-flow.ps1       # 描画エンジン（PowerShell 7 + Excel COM）
└── assets/
    └── sample-flow.json        # 動作確認用サンプル（デジ庁記載例「証明書の交付」）
```

## 使い方

### スキルとして（Claude Code など）

`business-flow-writer/` ディレクトリを Claude のスキル配置先に置くと、「業務フロー書いて」「BPMNで整理して」等の依頼で発動します。ワークフローの詳細は [SKILL.md](SKILL.md) を参照してください。

### スクリプトを直接実行

サンプル定義から xlsx（＋検証用PDF）を生成する例:

```powershell
pwsh -File scripts/generate-flow.ps1 `
    -Definition assets/sample-flow.json `
    -Output output/sample-flow.xlsx `
    -ExportPdf
```

| パラメータ | 説明 |
|---|---|
| `-Definition` | フロー定義JSONのパス |
| `-Output` | 出力する .xlsx のパス |
| `-ExportPdf` | 同名の PDF も出力（任意） |
| `-Visible` | Excel を画面に表示しながら実行（デバッグ用・任意） |

フロー定義JSONの書き方は [references/notation.md](references/notation.md) を参照してください。

## 既知の制限

- 差し戻しループ・線の交差回避の自動レイアウトはしない（手修正前提）
- 既存Excelフロー図の読み取り（As-Is取り込み）は未対応
- メッセージフロー（プール間の封筒付き破線）は association（破線）で代用

## ライセンス

[MIT License](LICENSE)
