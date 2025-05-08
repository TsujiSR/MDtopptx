"""
md2pptx-builder - Streamlit GUI Application
"""

import os
import logging
import tempfile
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Any

import streamlit as st
from PIL import Image

# 相対インポートに変更
from parser import MarkdownParser
from builder import PPTXBuilder
from utils import (
    setup_logging, is_valid_image, is_valid_markdown, 
    create_temp_file, clean_temp_files
)

# ロギング設定
logger = logging.getLogger(__name__)
setup_logging(verbose=False)

# アプリケーション情報
APP_TITLE = "md2pptx-builder"
APP_DESCRIPTION = "Markdownからロゴと背景画像を重ねたPowerPointを生成"

# 一時ファイル管理
temp_files = []

def register_temp_file(file_path: str) -> None:
    """一時ファイルを登録する
    
    Args:
        file_path: 一時ファイルパス
    """
    global temp_files
    temp_files.append(file_path)

def cleanup() -> None:
    """一時ファイルをクリーンアップする"""
    global temp_files
    if temp_files:
        clean_temp_files(temp_files)
        temp_files = []

def save_uploaded_file(uploaded_file) -> Optional[str]:
    """アップロードされたファイルを一時ファイルとして保存する
    
    Args:
        uploaded_file: アップロードされたファイルオブジェクト
        
    Returns:
        Optional[str]: 保存されたファイルパス、またはNone
    """
    if uploaded_file is None:
        return None
        
    try:
        # 一時ファイル作成
        suffix = Path(uploaded_file.name).suffix
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        temp_file.write(uploaded_file.getvalue())
        temp_file.flush()
        temp_file.close()
        
        # 一時ファイル登録
        register_temp_file(temp_file.name)
        return temp_file.name
    except Exception as e:
        st.error(f"ファイル保存エラー: {e}")
        return None

def display_slide_preview(slides_data: List[Dict[str, Any]]) -> None:
    """スライドプレビューを表示する
    
    Args:
        slides_data: スライドデータのリスト
    """
    if not slides_data:
        st.warning("スライドがありません。Markdownテキストを確認してください。")
        return
        
    st.subheader("スライドプレビュー")
    
    for i, slide in enumerate(slides_data):
        with st.expander(f"スライド {i+1}: {slide['title']}"):
            st.markdown(slide["raw_text"])

def create_presentation(
    md_content: str, 
    background_path: str, 
    logo_path: str, 
    template_path: Optional[str] = None,
    output_filename: str = "output.pptx",
    font_family: str = "メイリオ"
) -> Optional[str]:
    """プレゼンテーションを作成する
    
    Args:
        md_content: Markdownテキスト
        background_path: 背景画像パス
        logo_path: ロゴ画像パス
        template_path: テンプレートパス（オプション）
        output_filename: 出力ファイル名
        font_family: 使用するフォント
        
    Returns:
        Optional[str]: 出力ファイルパスまたはNone
    """
    try:
        # Markdownパーサー初期化
        parser = MarkdownParser()
        
        # 一時Markdownファイル作成
        md_file = create_temp_file(md_content)
        register_temp_file(md_file)
        
        # Markdownを処理
        slides_data = parser.process_markdown_file(md_file)
        
        if not slides_data:
            st.warning("変換可能なスライドがありません。")
            return None
        
        # 出力ファイルパス設定
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, output_filename)
        register_temp_file(output_path)
        
        # PPTXビルダー初期化
        builder = PPTXBuilder(
            background_path=background_path,
            logo_path=logo_path,
            template_path=template_path,
            font_family=font_family,
            verbose=False
        )
        
        # プレゼンテーション構築
        builder.build_presentation(slides_data, output_path)
        
        return output_path
        
    except Exception as e:
        st.error(f"変換エラー: {e}")
        logger.error(f"変換エラー: {e}", exc_info=True)
        return None

def app():
    """Streamlitアプリケーションのメイン関数"""
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="📊",
        layout="wide"
    )
    
    # タイトル
    st.title(APP_TITLE)
    st.markdown(APP_DESCRIPTION)
    
    # サイドバー（入力フォーム）
    with st.sidebar:
        st.header("入力設定")
        
        # Markdownテキスト入力（ファイルアップロードまたはテキストエリア）
        md_upload = st.file_uploader("Markdownファイルをアップロード", type=["md", "markdown", "txt"])
        
        if md_upload:
            md_content = md_upload.getvalue().decode("utf-8")
            st.success(f"ファイルをアップロードしました: {md_upload.name}")
        else:
            md_content = st.text_area(
                "Markdownテキストを入力",
                height=200,
                placeholder="# スライド1\n\nコンテンツ\n\n---\n\n# スライド2\n\nコンテンツ"
            )
        
        # 背景画像
        background_file = st.file_uploader("背景画像をアップロード", type=["png", "jpg", "jpeg"])
        if background_file:
            background_path = save_uploaded_file(background_file)
            
            # プレビュー
            st.image(background_file, caption="背景画像", use_container_width=True)
        else:
            background_path = None
            st.warning("背景画像をアップロードしてください")
        
        # ロゴ画像
        logo_file = st.file_uploader("ロゴ画像をアップロード", type=["png", "jpg", "jpeg"])
        if logo_file:
            logo_path = save_uploaded_file(logo_file)
            
            # プレビュー
            st.image(logo_file, caption="ロゴ画像", use_container_width=True)
        else:
            logo_path = None
            st.warning("ロゴ画像をアップロードしてください")
        
        # フォント設定
        st.subheader("フォント設定")
        font_options = {
            "メイリオ": "メイリオ (Meiryo)",
            "游ゴシック": "游ゴシック (Yu Gothic)",
            "游明朝": "游明朝 (Yu Mincho)",
            "MS Pゴシック": "MS Pゴシック (MS PGothic)",
            "MS P明朝": "MS P明朝 (MS PMincho)",
            "BIZ UDゴシック": "BIZ UDゴシック (BIZ UDGothic)",
            "BIZ UD明朝": "BIZ UD明朝 (BIZ UDMincho)",
            "UD デジタル 教科書体": "UD デジタル 教科書体 (UD Digi Kyokasho)",
            "Mplus 1p": "Mplus 1p",
            "Noto Sans JP": "Noto Sans JP",
            "Noto Serif JP": "Noto Serif JP",
            "Kosugi Maru": "Kosugi Maru",
            "Sawarabi Gothic": "Sawarabi Gothic",
            "Sawarabi Mincho": "Sawarabi Mincho",
        }
        
        # フォントカテゴリ (ゴシック体と明朝体を分ける)
        gothic_fonts = ["メイリオ", "游ゴシック", "MS Pゴシック", "BIZ UDゴシック", 
                       "UD デジタル 教科書体", "Mplus 1p", "Noto Sans JP", "Kosugi Maru", "Sawarabi Gothic"]
        mincho_fonts = ["游明朝", "MS P明朝", "BIZ UD明朝", "Noto Serif JP", "Sawarabi Mincho"]
        
        # フォントカテゴリ選択
        font_category = st.radio(
            "フォントカテゴリ",
            ["ゴシック体 (Sans-serif)", "明朝体 (Serif)"],
            index=0,
            help="スライドで使用するフォントのカテゴリを選択してください"
        )
        
        # カテゴリに応じたフォントリスト
        if "ゴシック体" in font_category:
            font_list = gothic_fonts
            default_index = 0 if "メイリオ" in font_list else 0
        else:
            font_list = mincho_fonts
            default_index = 0 if "游明朝" in font_list else 0
        
        # フォント選択ドロップダウン
        selected_font = st.selectbox(
            "フォント選択",
            options=font_list,
            index=default_index,
            format_func=lambda x: font_options.get(x, x),
            help="スライドで使用するフォントを選択してください。英語テキストには自動的に適切なフォールバックフォントが使用されます。"
        )
        
        # フォールバック説明
        if "明朝体" in font_category:
            st.info("英語テキストには Times New Roman がフォールバックとして使用されます。")
        else:
            st.info("英語テキストには Arial がフォールバックとして使用されます。")
        
        # テンプレート（オプション）
        template_file = st.file_uploader("テンプレートPPTX（オプション）", type=["pptx"])
        if template_file:
            template_path = save_uploaded_file(template_file)
            st.success(f"テンプレートをアップロードしました: {template_file.name}")
        else:
            template_path = None
        
        # 出力ファイル名
        output_filename = st.text_input("出力ファイル名", "output.pptx")
        if not output_filename.endswith(".pptx"):
            output_filename += ".pptx"
    
    # メイン画面
    # スライドプレビューと変換ボタン
    if md_content and is_valid_markdown(md_content):
        # 入力内容のチェック
        if not background_path or not logo_path:
            st.warning("変換を実行するには、背景画像とロゴ画像をアップロードしてください。")
        
        # ボタン列を作成
        col1, col2 = st.columns(2)
        
        # プレビューボタン
        with col1:
            preview_button = st.button("スライドをプレビュー", type="secondary")
            
        # 変換ボタン
        with col2:
            convert_button = st.button("PowerPointに変換", type="primary", disabled=not (background_path and logo_path))
        
        # プレビューボタンが押された場合
        if preview_button:
            with st.spinner("プレビュー生成中..."):
                # パーサー初期化
                parser = MarkdownParser()
                
                # スライドデータ取得
                slides_data = parser.process_markdown_content(md_content)
                
                # プレビュー表示
                display_slide_preview(slides_data)
        
        # 変換ボタンが押された場合
        if convert_button and background_path and logo_path:
            with st.spinner("変換中..."):
                # パーサー初期化
                parser = MarkdownParser()
                
                output_path = create_presentation(
                    md_content=md_content,
                    background_path=background_path,
                    logo_path=logo_path,
                    template_path=template_path,
                    output_filename=output_filename,
                    font_family=selected_font
                )
                
                if output_path:
                    # ファイル読み込み
                    with open(output_path, "rb") as f:
                        pptx_data = f.read()
                    
                    # ダウンロードボタン
                    st.success("変換が完了しました！")
                    st.download_button(
                        label="PowerPointをダウンロード",
                        data=pptx_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
    else:
        st.info("Markdownテキストを入力または、ファイルをアップロードしてください。")
    
    # フッター
    st.markdown("---")
    st.markdown("md2pptx-builder | Markdown to PowerPoint Converter")

def main():
    """アプリケーションエントリポイント"""
    try:
        app()
    finally:
        cleanup()

if __name__ == "__main__":
    main() 