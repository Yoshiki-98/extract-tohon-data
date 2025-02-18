# 謄本データ抽出アプリ

## 使い方

1. 「Browse files」から処理したいPDFファイルを選択します（複数選択可）
2. 「処理開始」ボタンをクリックします
3. 処理が完了すると、Excelファイルがダウンロード可能になります

## 注意事項

* アップロード可能なファイル形式は PDF のみです
* 特殊文字を含む氏名は Excel ファイル内で □ で出力されます
* 処理には時間がかかる場合があります

---

# Python開発環境セットアップガイド

## 前提条件
- macOS または Linux環境
- インターネット接続があること

## 1. pyenvのインストール

### Homebrewを使用してpyenvをインストール
```bash
brew install pyenv
```

### シェル設定ファイルの編集
```bash
# zshの場合（~/.zshrc）
echo 'export PYENV_ROOT="$HOME/.pyenv"' >> ~/.zshrc
echo 'export PATH="$PYENV_ROOT/bin:$PATH"' >> ~/.zshrc
echo 'eval "$(pyenv init --path)"' >> ~/.zshrc
echo 'eval "$(pyenv init -)"' >> ~/.zshrc

# 設定の反映
source ~/.zshrc
```

## 2. Pythonのインストール

### 利用可能なバージョンの確認
```bash
pyenv install --list
```

### Python 3.11.0のインストール
```bash
pyenv install 3.11.0
```

### グローバルPythonバージョンの設定
```bash
pyenv global 3.11.0
```

### インストールの確認
```bash
python --version
```

## 3. 仮想環境の作成と管理

### プロジェクトディレクトリの作成
```bash
mkdir extract-tohon-data
cd extract-tohon-data
```

### 仮想環境の作成
```bash
python -m venv venv
```

### 仮想環境の有効化
```bash
# macOS/Linux
source venv/bin/activate

# Windows
.\venv\Scripts\activate
```

## 4. 依存パッケージのインストール

### 必要なパッケージのインストール
```bash
pip install pandas
pip install numpy
pip install streamlit
pip install pdfplumber
pip install openpyxl
```

### 依存関係の書き出し
```bash
pip freeze > requirements.txt
```

## 5. プロジェクトの構成

```
extract-tohon-data/
├── .streamlit/
├── .gitignore
├── README.md
├── requirements.txt
├── venv/
├── data-export.py
```

## 6. .gitignoreの設定

```
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
venv/
.env

# IDE
.idea/
.vscode/
*.swp
*.swo

# OS
.DS_Store
Thumbs.db
```

## 注意事項
- venv ディレクトリはGitにコミットしない
- 環境変数を含む .env ファイルはGitにコミットしない
- requirements.txt は定期的に更新する

## トラブルシューティング

### pyenvインストール時のエラー
必要な依存関係をインストール：
```bash
brew install openssl readline sqlite3 xz zlib tcl-tk
```

### パッケージインストール時のエラー
pip のアップグレード：
```bash
pip install --upgrade pip
```

## 参考文献
- [Python公式ドキュメント](https://docs.python.org/)
- [pyenv GitHub](https://github.com/pyenv/pyenv)

---
作成日: 2025年2月18日
更新日: 2025年2月18日