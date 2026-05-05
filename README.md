# GAS 請求書自動生成ツール

> スプレッドシートへの入力から、請求書 PDF の自動生成・メール送付まで一気通貫で行う GAS ツール

[![GAS](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat-square&logo=google&logoColor=white)](https://www.google.com/script/start/)
[![Sheets](https://img.shields.io/badge/Google_Sheets-34A853?style=flat-square&logo=google-sheets&logoColor=white)](https://sheets.google.com)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)

---

## 📌 概要

フリーランス・小規模事業者向けの **請求書発行業務を完全自動化** する GAS ツールです。Google スプレッドシートに案件情報を入力するだけで、請求書を自動生成し、PDF に変換してメール送付まで行います。

---

## ✨ 主な機能

- **請求書の自動生成** — スプレッドシート入力から Google ドキュメントへ自動転記
- **PDF 変換** — 生成したドキュメントを PDF に自動変換
- **メール自動送付** — クライアントへ Gmail で自動送信
- **通し番号の自動付与** — 請求番号を自動でインクリメント
- **Google ドライブ保存** — 発行済み請求書を指定フォルダへ自動整理

---

## 🛠️ 使用技術

| 技術 | 用途 |
|---|---|
| Google Apps Script | メイン処理・PDF変換・メール送信 |
| Google スプレッドシート | 請求データの入力管理 |
| Google ドキュメント | 請求書テンプレート |
| Google ドライブ | 発行済みPDFの保存 |
| Gmail | 請求書メール送付 |

---

## 🚀 セットアップ

### 1. スプレッドシートの設定


### 2. GAS スクリプトの設置

1. スプレッドシートの「拡張機能」→「Apps Script」を開く
2. `請求書生成器.gs` を貼り付け
3. スクリプトプロパティに以下を設定：


### 3. 実行

スプレッドシートのカスタムメニュー「請求書を発行」をクリックするだけ。

---

## 📄 ライセンス

MIT License

---

<div align="right">

**Kei Assist** — 作業を仕組みに変える

</div>
