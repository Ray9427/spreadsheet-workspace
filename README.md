# スプレッドシート編集ワークスペース

Claude Code に Excel / CSV / Google スプレッドシートを編集してもらうための作業フォルダです。
Claude はこのフォルダを開くと `CLAUDE.md`(作業ルール)を自動で読み、次の手順で作業します。

1. 元のファイルの中身を確認する
2. 元ファイルは変更せず、`output/` にコピーを作って編集する
3. 数式エラーがないか・変えたいところだけが変わったかを自動でチェックする
4. 何を変えたかを報告する

---

## フォルダ構成

```
spreadsheet-workspace/
├── CLAUDE.md              ← Claude への作業ルール(自動で読まれる)
├── README.md              ← このファイル(人間向けの説明)
├── .claude/settings.json  ← 権限の設定(確認ダイアログを減らす・input/ を守る)
├── input/                 ← 編集してほしいファイルをここに入れる
│   └── sample_sales.xlsx  ← 動作確認用のサンプル(消してOK)
├── output/                ← 編集結果がここに保存される
├── scripts/
│   ├── inspect_xlsx.py    ← 中身をセル番地つきで表示
│   ├── recalc.py          ← 数式を再計算してエラー(#REF! など)を検出
│   ├── diff_xlsx.py       ← 編集前と編集後の差分を表示
│   └── work/              ← Claude が書いた編集用スクリプトが残る(あとから何をしたか追える)
└── requirements.txt       ← 必要な Python ライブラリ
```

---

## 使い方 A: パソコンの Claude Code で使う(いちばん簡単)

### 1. このフォルダをパソコンに置く

GitHub に置いてある場合は、リポジトリのページで **Code → Download ZIP** を押して展開します。
(git を使える人は `git clone <リポジトリのURL>` でもOK)

置き場所はどこでも構いません。例: `書類/spreadsheet-workspace`

### 2. Python とライブラリを入れる(初回だけ)

- Python が入っていなければ https://www.python.org/downloads/ から入れます。
  Windows ではインストール画面の **「Add python.exe to PATH」にチェック** を入れてください。
- ターミナル(Windows は PowerShell)でこのフォルダに移動して、ライブラリを入れます。

  ```bash
  # Mac
  cd ~/Documents/spreadsheet-workspace
  python3 -m pip install -r requirements.txt

  # Windows (PowerShell)
  cd $HOME\Documents\spreadsheet-workspace
  py -m pip install -r requirements.txt
  ```

### 3. LibreOffice を入れる(任意・おすすめ)

数式エラーのチェック(`scripts/recalc.py`)に使います。なくても編集はできます。
https://ja.libreoffice.org/ からダウンロード(Mac で Homebrew がある人は `brew install --cask libreoffice`)。

### 4. Claude Code を起動する

```bash
cd ~/Documents/spreadsheet-workspace   # このフォルダに移動
claude                                 # Claude Code を起動
```

Claude Code のインストール方法は公式ドキュメントを見てください: https://code.claude.com/docs
デスクトップアプリを使う場合は、このフォルダを作業フォルダとして選べばOKです。

### 5. ファイルを入れて頼む

編集したい Excel / CSV を `input/` に入れて、Claude に頼みます(頼み方の例は下にあります)。
結果は `output/` に `〇〇_編集済.xlsx` という名前で保存されます。

> **注意**: 編集中はそのファイルを Excel で閉じておいてください(開いたままだと保存に失敗することがあります)。

---

## 使い方 B: クラウド(claude.ai/code・スマホアプリ)で使う

クラウド版は毎回まっさらな環境で動くので、このフォルダを **GitHub のリポジトリ** にしておく必要があります。

### 1. GitHub にリポジトリを作る(初回だけ)

1. https://github.com/new を開く
2. Repository name に `spreadsheet-workspace` などと入れる
3. **Private(非公開)** を選ぶ(仕事のデータを入れるため)
4. 「Add a README file」などのチェックは **すべて外したまま** Create repository
5. このフォルダの中身をそのリポジトリに入れる(Claude に「このリポジトリに push して」と頼めます)

### 2. 編集したいファイルを入れる

GitHub のリポジトリのページで `input` フォルダを開き、**Add file → Upload files** でファイルをドラッグ → **Commit changes**。

### 3. Claude に頼む

claude.ai/code(またはスマホの Claude アプリ)で新しいセッションを作り、リポジトリに `spreadsheet-workspace` を選んで頼みます。

### 4. 結果を受け取る

Claude は作業が終わると、結果を GitHub の **`claude/` で始まるブランチ** に push します。

- GitHub のリポジトリページ左上のブランチ切り替え(`main` と書いてあるボタン)からそのブランチを選ぶ
- `output/` フォルダの中のファイルを開き、右上の **ダウンロードボタン(↓)** で保存

`main` に反映したいときは、Claude に「プルリクエストを作って」と頼み、GitHub 上で Merge します。

---

## Google スプレッドシートを編集する場合

ファイルの出し入れは不要です。claude.ai で Google Drive(Google スプレッドシート)を接続しておけば、URL を貼るだけで Claude が直接読み書きします。

```
https://docs.google.com/spreadsheets/d/xxxxx/edit
このスプレッドシートの「顧客」シートで、電話番号の書き方をハイフンありに統一して
```

Claude は「読む → 何を変えるか説明 → 書き込む」の順で作業し、変更した範囲を報告します。
大量の上書きや削除の前には確認を求めます。

---

## 頼み方の例

```
input/sample_sales.xlsx の中身を説明して
```
```
売上シートに「税込金額」の列を追加して。税率10%、1円未満は切り捨てで
```
```
担当者ごと・月ごとの売上を集計したシートを追加して
```
```
input/社員.csv と input/勤怠.csv を社員番号でくっつけて、1つの Excel にまとめて
```
```
重複している行を探して。消す前に一覧を見せて
```
```
B列の日付の書き方がバラバラなので、2026/10/15 の形にそろえて
```

**うまく頼むコツ**
- ファイル名・シート名・列(「B列」「金額の列」など)を具体的に書く
- 計算のルール(端数の切り捨て/四捨五入、税率など)を書く
- 「まず方針だけ見せて」と言えば、編集前に確認できる

---

## 安全のためのしくみ

| しくみ | 内容 |
|---|---|
| 元ファイルを守る | `CLAUDE.md` のルール+ `.claude/settings.json` の deny 設定で、`input/` のファイルは編集されません |
| 数式で残す | 合計などは数値ではなく Excel の数式で入るので、あとで数字を直しても自動で再計算されます |
| 自動チェック | 数式エラーの検出(`recalc.py`)と、変更箇所の一覧(`diff_xlsx.py`)を毎回行います |
| 作業の記録 | 編集に使ったスクリプトが `scripts/work/` に残ります |
| 削除は確認 | 行・列・シートの削除や大量の書き換えは、実行前に確認されます |

---

## 自分でもチェックしたいとき

```bash
python3 scripts/inspect_xlsx.py input/sample_sales.xlsx                 # 中身を見る
python3 scripts/recalc.py output/sample_sales_編集済.xlsx --show 10     # 数式エラーを調べる
python3 scripts/diff_xlsx.py input/sample_sales.xlsx output/sample_sales_編集済.xlsx   # 差分を見る
```

(Windows は `python3` を `py` に読み替えてください)

---

## よくある質問

**Q. 編集後のファイルをプレビューすると、数式のセルが空欄に見える**
A. Python(openpyxl)で保存したファイルには数式の「計算結果」が入っていないためです。Excel や Google スプレッドシートで開けば自動で計算されます。

**Q. 毎回「実行していいですか?」と聞かれる**
A. 確認用スクリプトと `scripts/work/` のスクリプトの実行は、確認なしで動くように設定済みです。それ以外で聞かれたら、内容を見て「Yes, and don't ask again」を選ぶと次回から聞かれません。

**Q. CSV を Excel で開くと文字化けする**
A. 「Excel で開くので UTF-8(BOM付き)で保存して」と頼んでください。

**Q. 自分で一からフォルダを作りたい**
A. 次のコマンドで空のフォルダ構成だけ作れます。そのあと、このリポジトリの `CLAUDE.md`・`.claude/settings.json`・`scripts/` をコピーしてください。

```bash
# Mac
mkdir -p ~/Documents/spreadsheet-workspace/{input,output,scripts/work,.claude}

# Windows (PowerShell)
cd $HOME\Documents
mkdir spreadsheet-workspace\input, spreadsheet-workspace\output, spreadsheet-workspace\scripts\work, spreadsheet-workspace\.claude
```

---

## 歩合表の軽量化(Google Apps Script)

Google スプレッドシートの歩合表を軽くするための道具と資料です。本番の実データ・ファイル ID は入れていません。

| 場所 | 内容 |
|---|---|
| `docs/手順書_1_バックアップと診断.md` | 最初にやること(バックアップと、変更しない診断)のクリック単位の手順 |
| `docs/報告_v2.md` | v1 の監査、v2 の作り、検証結果、外部連携の検討、判断してほしいこと |
| `gas/軽量化ツール_v2.gs` | 診断ツールのコード(新しい軽いスプレッドシートの Apps Script に貼る) |
| `gas/v1_参考_使わない/` | 以前の v1。監査の結果、使わない |
| `tests/` | ローカル試験(`node tests/test_v2.js`。匿名の作り物のデータだけを使う) |
