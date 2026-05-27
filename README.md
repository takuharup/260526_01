# test_260525_03 — VBA マクロ解説

## セットアップ手順

### ① VBS でモジュールをインポート（コマンドプロンプト）

```
cscript apply_refactor.vbs <xlsm のフルパス>
```

`ParseSupportReaction.bas` と `Module01.bas` を Excel に取り込む。

---

### ② フォームを生成（Excel 上でマクロ実行）

Alt+F8 から以下の 2 つを順番に実行する。

| マクロ名 | 生成されるフォーム |
|---|---|
| `CreateFormNodeSelect` | 節点番号選択ダイアログ |
| `CreateUserForm01` | DAT 出力プレビューダイアログ |

> **必要な設定（初回のみ）**  
> Excel → オプション → セキュリティセンター → セキュリティセンターの設定  
> → マクロの設定 →「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック

---

## マクロ一覧

### Module01

#### `ShowForm`
UserForm01（DAT 出力プレビュー）を開く。  
Alt+F8 から直接呼び出せるメインエントリポイント。

---

#### `RunParseAndSelectNodes`
支点反力テキストを読み込み、節点番号を選択して新規シートへ転記する。  
→ 内部で `ParseSupportReaction.ParseAndSelectNodes` を呼ぶラッパー。

---

#### `RunFilterByNode`
既存の支点反力シートから指定節点番号の行だけを抽出して新規シートへコピーする。  
→ 内部で `ParseSupportReaction.FilterByNode` を呼ぶラッパー。

---

#### `GenerateCombText` *(Function)*
アクティブシートの荷重組み合わせ表を読み取り、STRUDL DAT 形式のテキストを生成して返す。

**シートの想定レイアウト**

```
        | 組合番号1 | 組合番号2 | ...
荷重番号1|  係数      |  係数    | ...
荷重番号2|  係数      |  係数    | ...
```

- 1 行目の各列 = 組み合わせ番号（数値）
- 1 列目の各行 = 荷重番号（数値）
- セルの値が 0 の場合はその荷重を出力しない

**出力例**

```
Load comb 1 comb 1 1.2 2 1.0
Load comb 2 comb 1 1.0 3 0.8
```

---

#### `SaveDatFile`
`GenerateCombText` で生成したテキストをダイアログで指定したパスに `.dat` ファイルとして保存する。  
デフォルトのファイル名は `<ブック名>_comb.dat`。

---

#### `CreateFormNodeSelect` *(セットアップ用)*
`FormNodeSelect`（節点番号選択ダイアログ）を VBA コードで生成する。  
**セットアップ時に 1 回だけ実行する。**

生成されるコントロール：
- `lstNodes` — 節点番号リスト（複数選択可）
- `btnOK` / `btnCancel` ボタン

---

#### `CreateUserForm01` *(セットアップ用)*
`UserForm01`（DAT 出力プレビューダイアログ）を VBA コードで生成する。  
**セットアップ時に 1 回だけ実行する。**

生成されるコントロール：
- `cmbTemplate` — 出力テンプレート選択（現在は `STRUDL dat` のみ）
- `txtPreview` — 出力テキストのプレビュー表示欄
- `btnPreview` — プレビューを再生成
- `btnCreate` — DAT ファイルを保存
- `btnCancel` — 閉じる

---

### ParseSupportReaction

#### `ParseAndSelectNodes`
支点反力テキストファイルを読み込み、節点番号を選択して結果を新規シートへ転記する。

**操作フロー**

1. ファイル選択ダイアログで支点反力 `.txt` を選択
2. テキストを 1 パースして全行を収集、節点番号のユニーク一覧を作成
3. `FormNodeSelect` ダイアログで出力したい節点番号を選択（複数可）
4. 選択した節点の行だけを新規シート（`支点反力_YYYYMMDD_HHMMSS`）へ転記

**新規シートのヘッダー**

| 荷重番号 | 荷重名称 | 節点番号 | RX | RY | RZ | RMX | RMY | RMZ |
|---|---|---|---|---|---|---|---|---|

---

#### `FilterByNode`
既に Excel に読み込まれている支点反力シートを対象に、指定した節点番号の行だけを抽出する。  
（テキストファイルの再読み込みなしで繰り返しフィルタリングしたいときに使う）

**操作フロー**

1. InputBox で抽出元シート名を入力
2. InputBox で抽出する節点番号をカンマ区切りで入力  
   例：`1,3,合計`
3. 条件に一致する行を新規シート（`抽出_YYYYMMDD_HHMMSS`）へコピー

---

## ファイル構成

```
260526_01/
├── apply_refactor.vbs          ← セットアップ用 VBScript
├── test_260525_03.xlsm         ← 対象 Excel ファイル
└── test_260525_03/
    ├── Module01.bas            ← メインモジュール（CP932）
    ├── ParseSupportReaction.bas← 支点反力パーサ（CP932）
    ├── CopyToOpenBook.bas      ← 範囲コピーユーティリティ
    ├── FormNodeSelect.frm/.frx ← 節点選択フォーム（参照用）
    └── UserForm01.frm/.frx     ← DAT 出力フォーム（参照用）
```
