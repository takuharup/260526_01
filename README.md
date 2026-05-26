# test_260525_03 — VBA モジュールリファクタリング

## 概要

`test_260525_03.xlsm` の VBA プロジェクトを整理した。
重複していたモジュール・フォームを削除し、命名を統一した上で、
マクロがユーザーから正しく見える状態に修正した。

---

## 変更前の問題点

| 問題 | 詳細 |
|------|------|
| `UserForm2` の名前不一致 | `ParseSupportReaction` 内で `New FormNodeSelect` と参照しているが、フォームの VB_Name は `UserForm2` のままだったため実行時エラー |
| `Module1` が重複 | `ParseSupportReaction` に移植済みの古い実装が残っていた |
| `UserForm1` が重複 | `UserForm01` の古いコピーが残っていた |
| `Option Private Module` | `ParseSupportReaction` にこの宣言があり、`ParseAndSelectNodes` / `FilterByNode` が Alt+F8 のマクロ一覧に表示されなかった |

---

## 実施した変更

### VBA プロジェクト内（.xlsm バイナリ）

| 対象 | 変更内容 |
|------|----------|
| `UserForm2` | VB_Name・OLE ディレクトリ・dir ストリーム・PROJECT ストリームすべてで `FormNodeSelect` に改名 |
| `ParseSupportReaction` | 先頭の `Option Private Module` 行を削除 |
| `Module1` | 内容をスタブ（`Attribute VB_Name = "Module1"` のみ）に差し替え |
| `UserForm1` | 内容をスタブ（属性行のみ）に差し替え |
| `Module01` | 末尾に `RunParseAndSelectNodes` / `RunFilterByNode` の公開ラッパーサブを追加 |

### ソースファイル（`test_260525_03/` フォルダ）

| ファイル | 変更 |
|----------|------|
| `FormNodeSelect.frm` | 新規作成（`UserForm2.frm` を改名・VB_Name 修正） |
| `FormNodeSelect.frx` | 新規作成（`UserForm2.frx` をそのままコピー） |
| `Module01.bas` | ラッパーサブ 2 件を追記 |
| `ParseSupportReaction.bas` | `Option Private Module` 行を削除 |
| `Module1.bas` | 削除 |
| `UserForm1.frm` / `.frx` | 削除 |
| `UserForm2.frm` / `.frx` | 削除（→ FormNodeSelect に改名） |

### 変更後のモジュール構成

```
test_260525_03.xlsm
├── Module01          ← FormatValue / GenerateCombText / SaveDatFile / ShowForm
│                        + RunParseAndSelectNodes / RunFilterByNode (追加)
├── ParseSupportReaction  ← ParseAndSelectNodes / FilterByNode（Public・Alt+F8 から呼出可）
├── CopyToOpenBook    ← 変更なし
├── UserForm01        ← 変更なし
├── FormNodeSelect    ← 節点番号選択ダイアログ（旧 UserForm2）
├── Module1           ← スタブ（空）
└── UserForm1         ← スタブ（空）
```

---

## 生成したファイル

### `patch_vba.py`

xlsm バイナリを直接パッチする Python スクリプト。
Excel を起動せずに VBA ソースを書き換えるために実装した。

**主要な処理フロー:**

```
xlsm (ZIP) を読み込む
  └─ xl/vbaProject.bin を取り出す（OLE Compound File Binary）
       ├─ 各モジュールストリーム（VBA/Module01 等）を特定
       │    └─ MS-OVBA 圧縮を解凍して VBA ソースを取得
       ├─ ソースを編集（削除・追記・改名）
       ├─ MS-OVBA 圧縮で再圧縮してストリームに書き戻す
       ├─ OLE ディレクトリエントリのサイズフィールドを更新
       ├─ VBA/dir ストリームを再圧縮（MODULENAME を改名）
       ├─ PROJECT ストリームを更新（BaseClass= を改名）
       └─ OLE ディレクトリの UTF-16LE 名フィールドを改名
xlsm (ZIP) に書き戻す
```

**実装のポイント:**

- **MS-OVBA 圧縮**: LZ77 + フラグバイト方式（MS-OVBA 2.4.1 仕様）を Python で実装。
  チャンクヘッダの式は `header = (n_comp - 1) | 0xB000`（`n_comp` = トークンバイト数）。
- **OLE FAT チェーン走査**: olefile ライブラリの内部構造を利用してセクタチェーンを特定し、
  bytearray への直接書き込みでストリームを上書き。
- **ディレクトリエントリのサイズ更新**: ストリーム内容を縮小した際、
  OLE ディレクトリエントリのサイズフィールド（エントリ先頭 +120 バイト目）を
  新しい論理サイズに更新しないと、ゼロパディング部分をデコンプレッサが誤って
  次チャンクと解釈してクラッシュする。このフィールドの更新が核心的な修正点。

---

## 使い方

リファクタリングをやり直す場合（元の xlsm から再実行）:

```bash
# 元ファイルに戻してから実行
git checkout HEAD -- test_260525_03.xlsm
python3 patch_vba.py
```

依存ライブラリ:

```bash
pip install olefile oletools
```
