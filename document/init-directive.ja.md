# `@init:` 仕様と使い方（ユーザー向け）

`@init:` は、**その場で小さな初期化をまとめて実行**し、以降の同じブロック内で使える値（例：`{{title}}`）を用意するための記法です。
`- @init: ...` の**1行目に書き始め**、必要なら**複数行をハンギングインデント**（次行以降をインデント）で続けます。`@init:` の行自体は**最終出力に残りません**。

> **重要**：行末の継続記号 ` +` は**非推奨**です（将来的に廃止予定）。
> **ハンギングインデント方式**のみを使ってください。ハンギングインデント内では**空行の挿入も可能**です。

---

## まずは最短で

```text
- @init: $defaults({ title: "無題", count: 0 })

- タイトル：{{title}}
- 件数：{{count}}
```

* `@init:` の中では、以下の3つのヘルパーが使えます：

  * `$get(key)` … 既存の値を読む
  * `$set(key, value)` … 値を書き込む（上書き）
  * `$defaults(object)` … 未定義のキーだけを一括で埋める
* `@init:` で用意/更新した値は、そのブロック（見出し配下、テンプレ展開直下など）で参照できます。

---

## 使いどころ（`@init:` と簡易スカラーの使い分け）

### 簡易スカラー変数宣言（`&name: value`）

* 例：

  ```text
  - &title: "本日の献立"
  - &max: 10
  ```
* **リテラル値をサッと置く**ときに最適。
* **計算や条件分岐は不可**。シンプル・安全・見通しが良い。

### `@init:`（今回の主役）

* 例：

  ```text
  - @init: 
      $defaults({ title: "無題", pageSize: 20 })
      if ($get("env") === "prod") { $set("pageSize", 50) }
  ```
* **条件分岐・計算・一括初期化**など、**少し“処理”が必要**な場面で使う。
* まとめて書ける／後から読みやすい。

---

## 書き方リファレンス

### 基本構文

```text
- @init: <ここに1行で書いてもOK>
```

### 複数行（推奨：ハンギングインデント）

* **次の行からインデント**して続けます。
* **空行も入れてOK**（読みやすさのために自由に挿入可能）。

```text
- @init: 
    $defaults({ locale: "ja-JP", theme: "light", pageSize: 20 })

    if ($get("env") === "prod") {
        $set("pageSize", 50)
    }
```

### 非推奨（将来廃止予定）：行末の ` +`

```text
- @init: const s = "a" +   # ← この「行末 +」継続記法は使わない
```

**使わないでください。** 既存ドキュメントは後述の「移行ガイド」を参考に置き換えてください。

---

## ヘルパーの挙動

### `$get(key)`

* 現在のブロックで参照可能な `key` を返します。未定義なら `undefined` 相当。
* 例：環境によって上書きするか判断

  ```text
  - @init:
      if ($get("env") === "prod") $set("debug", false)
  ```

### `$set(key, value)`

* `key` に `value` を設定（上書き）。
* 例：計算して保存

  ```text
  - @init:
      const base = 100
      const tax = 0.1
      $set("price", base * (1 + tax))
  ```

### `$defaults(object)`

* 渡したオブジェクトの**未定義キーだけ**を埋めます（既存値は保持）。
* 例：初期値をまとめて

  ```text
  - @init: $defaults({ title: "無題", retries: 0, theme: "light" })
  ```

---

## レシピ集（コピペOK）

### 1) 既存値を壊さずにデフォルト注入

```text
- @init:
    $defaults({
        title:   "無題",
        locale:  "ja-JP",
        pageSize: 20
    })
```

### 2) 既存値の有無で分岐 & 計算

```text
- @init:
    const n = $get("pageSize") || 20
    $set("maxItems", n * 3)
```

### 3) 条件で一部だけ上書き

```text
- @init:
    $defaults({ retries: 0, theme: "light" })

    if ($get("env") === "prod") {
        $set("retries", 3)
    }
```

### 4) 長文の組み立て（配列 + `join` 推奨）

> かつて**行末 ` +` 継続**で書いていたケースは、**配列の `join("\n")`** で安全に置き換えられます。

```text
- @init:
    const lines = [
        "処理を開始しました。",
        "しばらくお待ちください。"
    ]
    $set("statusMessage", lines.join("\n"))
```

### 5) よくある初期化セット

```text
- @init:
    $defaults({
        env:        "dev",
        debug:      true,
        theme:      "light",
        pageSize:   20,
        dateFormat: "YYYY-MM-DD"
    })
```

### 6) 文字列の連結（改行不要）

```text
- @init:
    const msg = "Hello, " + "World!"
    $set("message", msg)
```

### 7) 長い式を読みやすく

```text
- @init:
    const total =
        $get("subtotal")
      + $get("tax")
      - ($get("discount") || 0)

    $set("total", total)
```

---

## 非推奨記法からの移行ガイド

**Before（非推奨：行末 ` +`）**

```text
- @init: const msg = "line1" +    
         "line2"
         $set("message", msg)
```

**After（推奨：配列 + join）**

```text
- @init:
    const msg = ["line1", "line2"].join("\n")
    $set("message", msg)
```

**Before（非推奨：行末 ` +` で長い式）**

```text
- @init: const total = base +    
         tax - discount
```

**After（推奨：ハンギングインデントで整形）**

```text
- @init:
    const total =
        base
      + tax
      - discount
```

---

## よくある質問（FAQ）

**Q. `@init:` の行は出力に出ますか？**
A. 出ません。初期化だけ行い、**行自体は非表示**です。

**Q. どこまでの範囲で値が使えますか？**
A. 通常は**同じブロック（見出しの配下、テンプレ展開の本体など）**で参照できます。迷ったら、使いたい直前で `@init:` しておくのが確実です。

**Q. `&name:` と `@init:` は併用できますか？**
A. はい。**固定のリテラルは `&name:`、条件や計算は `@init:`** と分けるのがおすすめです。

**Q. 空行を入れても大丈夫？**
A. **はい、大丈夫です。** ハンギングインデント内の空行は読みやすさ向上に有効です。

---

## ベストプラクティス

* **読みやすさ優先**：複数行にして、**空行やコメント**で意図を残す。
* **デフォルトは `$defaults` で一括**、環境差分や一時的な上書きは `$set`。
* **“その場で必要な初期化だけ”** を短く書く（長大なロジックは避ける）。
* **`&name:` と役割分担**（リテラル＝`&name:`、処理＝`@init:`）。
* **行末 ` +` は使わない**（将来の非互換を避ける）。
