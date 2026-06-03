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

### 集計機能

- 総収入集計
- 総支出集計
- 総収益集計
- 総資産集計

### 月間収益分析

給与支給日を基準とした独自集計に対応しています。

例

給与支給日：25日

| 集計期間 | 集計対象月 |
|----------|-----------|
| 2026/05/25 ～ 2026/06/24 | 2026年6月 |
| 2026/06/25 ～ 2026/07/24 | 2026年7月 |

給与支給日が土日祝日の場合は、最も近い前営業日として扱います。

### 年間収益分析

指定した年の収入・支出・資産推移を集計します。

### コンフィグ機能

以下のマスタ情報を管理できます。

- 大区分管理
- 中区分管理
- 支払者管理
- 給与支給日設定
- 月予算設定
- 定期支出設定

---

## 使用技術

- Python 3.9
- PySide6
- SQLite3
- OpenPyXL
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

```bash
pyinstaller --onefile --windowed --clean --noconfirm --icon="kakeibo_icon_v2.ico" --add-data "kakeibo_icon_v2.ico;." --name "家計簿" household_gui.py
```

---

## データ保存先

アプリのデータはSQLiteデータベースに保存されます。

```text
household.db
```

バックアップを取得する場合は、上記ファイルをコピーしてください。

---

## フォルダ構成

```text
HomeExpensesManager
│
├─ household_gui.py
├─ household.db
├─ requirements.txt
├─ kakeibo_icon_v2.ico
│
├─ config
│   ├─ category_master.json
│   ├─ payer_master.json
│   ├─ salary_config.json
│   └─ budget_config.json
│
└─ backup
```

---

## 今後の実装予定

- CSVインポート
- CSVエクスポート
- 検索機能強化
- 条件フィルタ
- グラフ分析機能強化
- ダークモード対応
- 自動バックアップ
- パスワード保護

---

## 動作環境

| 項目 | 内容 |
|------|------|
| OS | Windows 10 / 11 |
| Python | 3.9以上 |
| DB | SQLite |
| GUI | PySide6 |

---

## ライセンス

個人利用・学習利用を想定しています。

---

## 開発者

- あきほ
- Python / PySide6
- HomeExpensesManager
