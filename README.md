# md2pptx-builder

Markdownファイルから会社ロゴと背景画像を重ねたPowerPointプレゼンテーションを自動生成するツール。
CLIとGUI（Streamlit）の両方に対応しています。

## 特徴

- Markdownファイルからスライドごとに分割して.pptxファイルを生成
- 背景画像とロゴを全スライドに適用
- 見出し、段落、箇条書き、コードブロックを適切にフォーマット
- テキスト強調（**太字**、*斜体*）のサポート
- 入れ子リスト（ネストされた箇条書き）のサポート
- 最適化されたフォントサイズとスペーシング
- CLI（コマンドライン）とGUI（Streamlit）の両方で利用可能

## インストール

### 推奨: 仮想環境を使用したインストール
```bash
# 1. リポジトリをクローン
git clone https://github.com/yourusername/md2pptx-builder.git
cd md2pptx-builder

# 2. 仮想環境を作成して有効化
## Windows
python -m venv venv
.\venv\Scripts\activate

## macOS/Linux
python -m venv venv
source venv/bin/activate

# 3. 依存パッケージをインストール
pip install -r requirements.txt

# 4. 開発モードでインストール
pip install -e .
```

### pipから直接インストール（予定）
```bash
pip install md2pptx-builder
```

### リポジトリから直接インストール
```bash
git clone https://github.com/yourusername/md2pptx-builder.git
cd md2pptx-builder
pip install -e .
```

## 使い方

### 推奨: 仮想環境内での実行

```bash
# 仮想環境が有効化されていることを確認
## Windows
.\venv\Scripts\activate
## macOS/Linux
source venv/bin/activate

# Streamlitアプリを起動
python -m streamlit run md2pptx_builder/app.py
```

### CLIから使用する場合

```bash
# 基本的な使い方
md2pptx-builder input.md -b background.jpg -l logo.png -o output.pptx

# 詳細ログを有効にする
md2pptx-builder input.md -b background.jpg -l logo.png -o output.pptx --verbose

# テンプレートを使用する場合
md2pptx-builder input.md -b background.jpg -l logo.png -o output.pptx -t template.pptx

# ヘルプを表示
md2pptx-builder --help
```

### GUIから使用する場合

```bash
# 仮想環境内でStreamlitアプリを起動（推奨）
python -m streamlit run md2pptx_builder/app.py

# または、モジュールとして実行
streamlit run -m md2pptx_builder.app
```

ブラウザで `http://localhost:8501/` にアクセスすると、GUIが表示されます。

## Markdownファイルの書き方

### スライド分割

Markdownファイルは以下の区切り文字でスライド分割されます：

```markdown
# スライド1のタイトル

コンテンツ

---

# スライド2のタイトル

コンテンツ
```

または、HTMLコメント形式も使用できます：

```markdown
# スライド1のタイトル

コンテンツ

<!-- pagebreak -->

# スライド2のタイトル

コンテンツ
```

### サポートされる書式

- **見出し**：`#`（スライドタイトル）、`##`（セクション見出し）、`###`（小見出し）
- **テキスト強調**：`**太字**`、`*斜体*`
- **リスト**：順序付き (`1. 項目`) ・順序なし (`- 項目`) リスト
- **コードブロック**：\`\`\` で囲まれたコードブロック

## レイアウト仕様

- **スライドサイズ**：16:9比率（標準的なワイドスクリーン）
- **フォント**：メイリオ（日本語）、Arial（英数字）
- **サイズ**：
  - タイトル：32pt
  - 見出し：22-18pt（レベルによる）
  - 本文：16pt
  - リスト：15pt（サブリスト：13pt）
  - コード：14pt

## 技術詳細

- **Markdownパーサー**：mistune 3.1.3
- **PowerPoint操作**：python-pptx
- **画像処理**：Pillow
- **GUI**：Streamlit

## 依存ライブラリ

- Python 3.8以上
- python-pptx >= 0.6.21
- mistune >= 3.0.0
- Pillow >= 9.0.0
- streamlit >= 1.20.0

## 開発

```bash
# 開発用インストール
git clone https://github.com/yourusername/md2pptx-builder.git
cd md2pptx-builder

# 仮想環境を作成して有効化
python -m venv venv
.\venv\Scripts\activate  # Windowsの場合
source venv/bin/activate  # macOS/Linuxの場合

# 依存パッケージをインストール
pip install -r requirements.txt

# 開発モードでインストール
pip install -e ".[dev]"

# テスト実行
pytest
```

## 最近の更新

- mistune 3.1.3との完全な互換性
- リスト項目とテキスト強調の表示問題を修正
- スライドレイアウトとスペーシングの改善
- フォントサイズの最適化

## ライセンス

MIT

## サンプル実行

```bash
# 仮想環境内でサンプルMarkdownからプレゼンテーションを生成
python -m md2pptx_builder samples/test.md -b background.jpg -l logo.jpg -o output.pptx
``` 