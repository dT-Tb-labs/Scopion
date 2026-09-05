---
layout: default
title: Scopion
---

# Scopion

**Trace a formula's inputs in Google Sheets.** Select a formula cell, press
<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>A</kbd>, and walk every cell it reads with
<kbd>↑</kbd> <kbd>↓</kbd> — across sheets, into hidden sheets, through named ranges.
Built for financial models.

**[Install from the Chrome Web Store](https://chromewebstore.google.com/search/Scopion)** ·
[Privacy policy](privacy/) · [Source and support](https://github.com/dT-Tb-labs/Scopion)

- **Origin row** — the list starts with the cell you audited; <kbd>↑</kbd> from the first input brings you home.
- **Drill** — <kbd>→</kbd> makes the highlighted cell the new origin, <kbd>←</kbd> goes back; breadcrumbs show the trail.
- **Blank inputs flagged** — an empty referenced cell is usually the error you were looking for.
- **Honest about dynamic references** — `OFFSET`, `INDEX`, `INDIRECT` are resolved from literals and single-cell values; the rest is listed as unresolved, never dropped.
- **One write, and only one** — Sheets unhides a hidden sheet when you jump into it; Scopion hides it again when the panel closes (on an Excel file it cannot, and says so). Nothing else is ever changed, and nothing leaves your browser except calls to Google's own API.

Scopion is a port of the Excel formula auditor generations of modellers relied on.

---

## 日本語

**Google スプレッドシートの数式が参照するセルをたどる拡張機能。** 数式セルを選んで
<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>A</kbd>。参照先が一覧され、<kbd>↑</kbd> <kbd>↓</kbd> で
シートをまたぎ、非表示シートの中へ、名前付き範囲を通って順にジャンプします。財務モデルのために。

**[Chrome ウェブストアからインストール](https://chromewebstore.google.com/search/Scopion)** ·
[プライバシーポリシー](privacy/) · [ソースコードとサポート](https://github.com/dT-Tb-labs/Scopion)

- **起点行** — 一覧は監査したセル自身から始まり、<kbd>↑</kbd> で元の位置へ。
- **ドリルダウン** — <kbd>→</kbd> でハイライト中のセルを新しい起点に、<kbd>←</kbd> で戻る。経路はブレッドクラムで表示。
- **空白入力を警告** — 参照先が空のセルは、たいてい探していた間違いそのもの。
- **動的参照に正直** — `OFFSET`・`INDEX`・`INDIRECT` はリテラルと単一セルの値から解決し、残りは「未解決」と明示。黙って省かない。
- **書き込みは一種類だけ** — ジャンプで表示状態になった非表示シートを、パネルを閉じるときに再び非表示へ（Excel ファイルでは戻せないため、その旨を表示）。それ以外は何も変更せず、Google の API 以外へデータは出ません。

Scopion は、Excel で長年使われてきた数式監査ツールの移植です。
