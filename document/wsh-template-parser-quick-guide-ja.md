# WSH JScript テンプレートパーサー〈クイックガイド〉

> 本ドキュメントは、WSH JScript で実装した Markdown UL → JSON（→Excel）テンプレートエンジンの**現行仕様の要点**と、**最小サンプル**をまとめたものです。
>
> 前提：lodash 3.10.1 利用。`extendScope(parent, layer)` 実装。`evalExprWithScope`（`with(scope)`）キャッシュ。`$args`注入。`$params`は互換。`$scope`は撤去。

---

## 1. 基本構造
- **テンプレート宣言**：`&Name()`
- **テンプレート呼び出し**：`*Name({...})`
- **スコープ優先**：`call-args > node.params > ancestors (近い順) > globalScope(conf UPPER_SNAKE)`
- **式の評価**：`with(scope)` で `eval`。プレースホルダと @init で共用。
- **ネスト**：親ローカルを引き継ぎ、内側の `params` / `args` を重ねて評価・展開。

---

## 2. ディレクティブ / ヘルパー（@init 内）
- `@init`：ノード単位で**置換前に実行**。同ノード内に複数可、**記述順**に走る。
- @init で使えるヘルパ：
  - `$defaults(obj)`：未定義キーにだけ既定値を入れる。
  - `$set(key, value)`：現在スコープに代入。
  - `$get(key, [fallback])`：現在スコープから取得。
- これらヘルパは **@init 実行中のみ**注入され、終了後に掃除される（JSON 出力時にも除外）。

---

## 3. プレースホルダ
- **基本**：`{{path}}` は `_.get(scope, path)` と等価。`{{ expr }}` は式可。
- **空の `{{}}`**：デフォルトは「**最初の引数キー** or `$value`」。
- **モード**：
  - `{{? expr}}`：`expr` が偽（`false/null/undefined`…）なら **ノードをドロップ**。
    - 厳格真偽が必要なら `Q_BOOL_STRICT` フラグで `0/""` を偽に含めない等の調整。
  - `{{expr!}}`：**必須**。`undefined/null` なら **エラー（または警告）**。
- **レガシー互換**：フラグ `PLACEHOLDER_LEGACY_DROP` で**素の `{{…}}` 評価が falsy → ノードドロップ**の旧挙動を維持可。
- **関連フラグ（例）**：
  - `PLACEHOLDER_WARN_ON_UNDEFINED`：未定義参照で警告ログ。
  - `PLACEHOLDER_STRICT_DEFAULT`：空 `{{}}` を**必須扱い**に切り替え可。

---

## 4. 呼び出し・反復・ループメタ
- **配列引数**：`*Row([{...},{...}])` → 要素ごとに展開。
- **数値シュガー**：
  - `*Row(5)` ≡ `*Row(Array(5))`（データ不要で**回数だけ**ループ）。
  - `*Row({ $times: 5, label: "X" })` → 5回同条件でループ。
- **ループメタ**：各回のスコープに注入（JSON 出力では除外）。
  - `$index`（0初）、`$index1`（1初）、`$count`、`$isFirst`、`$isLast`、`$isOdd`、`$isEven`

---

## 5. JSON への出力
- **除外キー**：`$args,$params,$get,$set,$defaults,$index1,$count,$isFirst,$isLast,$isOdd,$isEven` … などは **replacer の drop keys** で落とす。
- シリアライズは**テンプレート実行後**のノードに対して行う。

---

## 6. 既知のフラグ（例）
- `PLACEHOLDER_WARN_ON_UNDEFINED`
- `PLACEHOLDER_LEGACY_DROP`
- `PLACEHOLDER_STRICT_DEFAULT`
- `Q_BOOL_STRICT`
- `ENABLE_NUMERIC_REPEAT`（数値リピート） / `MAX_REPEAT`

---

## 7. 最小サンプル（コピペで動作確認）

### 7.1 テンプレート定義
```
- &Row()
  - @init: $defaults({ label: "default", badge: false })
  - Label: {{label}}
  - Note:  {{? note}}        # note が無ければ行をドロップ
  - Badge: {{? badge}}

- &Wrap()
  - @init
    - $defaults({ prefix: "[" })
    - $defaults({ suffix: "]" })
  - {{prefix}}{{text}}{{suffix}}

- &Outer()
  - @init: $defaults({ label: "outer" }); $defaults({ text: label }); $set("state", ($args && $args.debug) ? "on" : "off")
  - *Wrap({text: text})
  - Debug: {{state}}

- &InitDemo()
  - @init
    - $defaults({ base: 10 })
    - $set("answer", $get("base", 0) * 2)
  - Answer={{answer}}

- &RequireDemo()
  - Title: {{ title! }}   # 必須（undefined/null でエラー or 警告）

- &LoopDemo()
  - @init: $defaults({ label: "L" })
  - idx={{$index}} / idx1={{$index1}} / count={{$count}} / first={{$isFirst}} / last={{$isLast}} / even={{$isEven}} / odd={{$isOdd}}
  - Label: {{label}}
```

### 7.2 呼び出しケース
```
- *Row({label:"Foo", note:"first", badge:"HOT"})
- *Row({note:"second"})
- *Outer({label:"X", debug:true})
- *Outer({text:"World", debug:false})
- *InitDemo({ base: 7 })
- *RequireDemo({ title: "Hello" })
- *LoopDemo(3)
- *LoopDemo([{label:"A"},{label:"B"}])
```

### 7.3 期待される主な結果イメージ（テキスト整形後）
```
Label: Foo
Note: first
Badge: HOT

Label: default
Note: second
# （badge 行は note 同様に存在しなければドロップ）

[X]
Debug: on

[World]
Debug: off

Answer=14

Title: Hello

idx=0 / idx1=1 / count=3 / first=true / last=false / even=true / odd=false
Label: L
idx=1 / idx1=2 / count=3 / first=false / last=false / even=false / odd=true
Label: L
idx=2 / idx1=3 / count=3 / first=false / last=true / even=true / odd=false
Label: L

idx=0 / idx1=1 / count=2 / first=true / last=false / even=true / odd=false
Label: A
idx=1 / idx1=2 / count=2 / first=false / last=true / even=false / odd=true
Label: B
```
> 注：`{{ title! }}` が未指定の場合は **エラー/警告**（実際の挙動はフラグ設定に依存）。

---

## 8. ベストプラクティス / よくある落とし穴
- **@init は置換より前**：`$defaults/$set` でスコープ整形→プレースホルダ評価が安定。
- **`{{?}}` は構造ドロップ**：行そのものが消える。スペースや記号だけの「空行」対策もルール化しておくと吉。
- **空 `{{}}` の意味**をチームで共有：`$value` を使うか「最初のキー」にするか、`PLACEHOLDER_STRICT_DEFAULT` を使うかを明確に。
- **数値リピート**には上限 `MAX_REPEAT` を設定（DoS 自衛）。
- ループメタは**表示用**で、**JSON には出さない**（drop keys に追加済みか確認）。

---

## 9. 変更フラグの例（conf）
```js
var conf = {
  ENABLE_NUMERIC_REPEAT: true,
  MAX_REPEAT: 500,
  PLACEHOLDER_WARN_ON_UNDEFINED: true,
  PLACEHOLDER_LEGACY_DROP: false,
  PLACEHOLDER_STRICT_DEFAULT: false,
  Q_BOOL_STRICT: false
};
```

---

## 10. チートシート
- 宣言：`&Name()` / 呼び出し：`*Name({...})`
- スコープ：`args > params > ancestors > global`
- @init：`$defaults/$set/$get`（実行順に注意）
- 置換：`{{path}}` / `{{ expr }}` / 空 `{{}}` は `$value` or 最初のキー
- 条件：`{{? expr}}`（偽でノード削除）
- 必須：`{{ expr! }}`（未定義/null でエラー/警告）
- 反復：`*T([…])` / `*T(5)` / `*T({$times:n,…})`
- ループメタ：`$index,$index1,$count,$isFirst,$isLast,$isOdd,$isEven`（JSON出力では除外）

---

## 11. 今後の拡張のタネ
- `@if/@else` の軽量ブロック
- シートテンプレ（`&SheetTpl()` + `*SheetTpl({$name:…})`）
- ドライラン／トレース（ドロップマーク・決定ログ）
- プレースホルダ式キャッシュの統計ダンプ

---

### 付記
- 本ガイドは**実装の現在地**に合わせた要点整理です。挙動差があればここに追記していきます。
- これを「Quick Handoff」に添えておけば、新規スレでも即合流できます。

