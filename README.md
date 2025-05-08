# md2pptx-builder

Markdownファイルから会社ロゴと背景画像を重ねたPowerPointプレゼンテーションを自動生成するツール。
CLIとGUI（Streamlit）の両方に対応しています。

## 機能

- Markdownファイルからスライドごとに分割して.pptxファイルを生成
- 背景画像とロゴを全スライドに適用
- 見出し、段落、箇条書きを適切にフォーマット
- CLI（コマンドライン）とGUI（Streamlit）の両方で利用可能

## インストール

```bash
pip install md2pptx-builder
```

## 使い方

### CLIから使用する場合

```bash
# 基本的な使い方
md2pptx-builder input.md -b background.jpg -l logo.png -o output.pptx

# ヘルプを表示
md2pptx-builder --help
```

### GUIから使用する場合

```bash
# Streamlitアプリを起動
streamlit run -m md2pptx_builder.app
```

ブラウザで `http://localhost:8501/` にアクセスすると、GUIが表示されます。

## 入力ファイル形式

Markdownファイルは以下の区切り文字でスライド分割されます：

```
# スライド1のタイトル

コンテンツ

---

# スライド2のタイトル

コンテンツ
```

## 必要環境

- Python 3.10以上
- 依存ライブラリ:
  - python-pptx
  - mistune
  - Pillow
  - streamlit

## ライセンス

MIT 

md2pptx-builder samples/test.md -b <背景画像パス> -l <ロゴ画像パス> -o output.pptx 