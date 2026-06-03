# HomeExpensesManager

Windows向けのデスクトップ家計簿アプリです。

Python（PySide6）とSQLiteを使用して開発しており、日々の収入・支出・貯蓄を管理しながら、月次・年次で収支分析を行うことができます。

---

## 主な機能

### 家計簿管理

- 収入登録
- 支出登録
- 金融貯蓄登録
- 現金／預金貯蓄登録
- データ編集
- データ削除

### マスタ管理

- 大区分管理
- 中区分管理
- 支払者管理

### 予算管理

- 月予算設定
- カテゴリ別予算管理
- 予算超過チェック

### 定期支出管理

- 定期支出登録
- 定期収入登録
- 定期積立登録
- 指定月への自動生成
- 重複登録防止

### 集計機能

- 総収入集計
- 総支出集計
- 総収益集計
- 総資産集計

### 月間収益分析

給与支給日を基準とした独自集計に対応しています。

例

| 給与支給日 | 集計期間 |
|------------|----------|
| 25日 | 2026/05/25 ～ 2026/06/24 |
| 25日 | 2026/06/25 ～ 2026/07/24 |

土曜日の場合

```text
2026/07/25(土)
↓
2026/07/24(金)
```

日曜日の場合

```text
2026/10/25(日)
↓
2026/10/26(月)
```

※現在は祝日補正には対応していません。

### 年間収益分析

- 年間収入集計
- 年間支出集計
- 年間収益集計
- 年間資産集計

### 分析グラフ

#### 月次収益推移

- 月ごとの収益推移を折れ線グラフ表示

#### カテゴリ別支出分析

- 円グラフ表示
- 支出割合分析

### Excel取込

Excel家計簿からデータを取り込み可能です。

対応形式

```text
.xlsx
.xlsm
```

---

## 使用技術

### 言語

- Python 3.9

### GUI

- PySide6

### データベース

- SQLite3

### Excel

- OpenPyXL

### グラフ

- Matplotlib

---

## インストール

### 1. リポジトリ取得

```bash
git clone https://github.com/akihotabata/HomeExpensesManager.git
cd HomeExpensesManager
```

### 2. 仮想環境作成

```bash
python -m venv venv
```

### 3. 仮想環境有効化

```bash
venv\Scripts\activate
```

### 4. 必要ライブラリインストール

```bash
pip install -r requirements.txt
```

---

## 起動方法

```bash
python household_gui.py
```

---

## EXE化

### PyInstallerインストール

```bash
pip install pyinstaller
```

### EXE生成

```bash
rmdir /s /q build dist

del /q *.spec

pyinstaller --onefile --windowed --clean --noconfirm --icon="%cd%\kakeibo_icon_v2.ico" --add-data "%cd%\kakeibo_icon_v2.ico;." --name "HouseHoldExpensesManager" household_gui.py
```

生成先

```text
dist\家計簿.exe
```

---

## データ保存先

アプリ起動フォルダにSQLiteデータベースが自動生成されます。

```text
household.db
```

バックアップ方法

```text
household.db
をコピーするだけ
```

---

## フォルダ構成

```text
HomeExpensesManager
│
├─ household_gui.py
├─ requirements.txt
├─ kakeibo_icon_v2.ico
├─ household.db
│
├─ build
├─ dist
└─ venv
```

---

## 画面一覧

### 総収益

- 全期間収入
- 全期間支出
- 全期間収益
- 全期間資産

### 入力

- 家計簿登録
- 家計簿更新
- 家計簿編集

### 家計簿一覧

- 全期間表示
- 月間表示
- 年間表示
- 削除機能

### 月間収益

- 月間集計
- 予算比較
- 予算超過判定

### 年間収益

- 年間集計
- 年間分析

### 分析グラフ

- 月次収益推移
- カテゴリ別支出分析

### コンフィグ

- 大区分設定
- 中区分設定
- 支払者設定
- 月予算設定
- 定期支出登録
- 給与支給日設定

---

## 動作環境

| 項目 | 内容 |
|------|------|
| OS | Windows 10 |
| OS | Windows 11 |
| Python | 3.9以上 |
| DB | SQLite3 |
| GUI | PySide6 |

---

## 注意事項

- Windows環境専用です
- SQLiteを使用しています
- ネットワーク接続は不要です
- データはローカルPC内のみ保存されます
- クラウド同期機能はありません
- 祝日補正は未対応です

---

## 今後の実装候補

- CSVインポート
- CSVエクスポート
- 検索機能強化
- 条件フィルタ
- グラフ機能強化
- ダークモード
- 自動バックアップ
- パスワード保護
