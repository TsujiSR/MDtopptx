"""
md2pptx-builder - PowerPoint builder
"""

import os
import logging
from typing import List, Dict, Any, Optional, Tuple, Union
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from md2pptx_builder.utils import is_valid_image, get_image_dimensions

logger = logging.getLogger(__name__)

class PPTXBuilder:
    """MarkdownからPowerPointを生成するクラス"""
    
    def __init__(self, 
                 background_path: str, 
                 logo_path: str, 
                 template_path: Optional[str] = None,
                 verbose: bool = False):
        """
        Args:
            background_path: 背景画像のパス
            logo_path: ロゴ画像のパス
            template_path: テンプレートPPTXのパス（オプション）
            verbose: 詳細ログを出力するかどうか
        """
        self.background_path = background_path
        self.logo_path = logo_path
        self.template_path = template_path
        self.verbose = verbose
        
        # 画像ファイルのチェック
        if not is_valid_image(background_path):
            raise ValueError(f"無効な背景画像: {background_path}")
        if not is_valid_image(logo_path):
            raise ValueError(f"無効なロゴ画像: {logo_path}")
        
        # プレゼンテーション作成
        if template_path and os.path.exists(template_path):
            self.prs = Presentation(template_path)
            logger.info(f"テンプレートを使用: {template_path}")
        else:
            self.prs = Presentation()
            logger.info("新規プレゼンテーションを作成")
            
        # デフォルトのスライドサイズを16:9に設定（テンプレートが無い場合）
        if not template_path:
            self.prs.slide_width = Inches(16 * 0.75)  # 16:9 比率
            self.prs.slide_height = Inches(9 * 0.75)
    
    def create_slide(self, slide_data: Dict[str, Any], total_slides: int) -> None:
        """スライドを作成する
        
        Args:
            slide_data: スライドデータ（タイトル、コンテンツなど）
            total_slides: スライドの総数
        """
        # レイアウトインデックス6は白紙のスライド
        layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(layout)
        
        # 背景画像設定
        self._apply_background(slide)
        
        # ロゴ設定
        self._add_logo(slide)
        
        # タイトル追加
        title = slide_data.get("title", f"スライド {slide_data['index'] + 1}")
        self._add_title(slide, title)
        
        # コンテンツ追加
        self._add_content(slide, slide_data["content"])
        
        # スライド番号追加
        current_slide = slide_data["index"] + 1
        self._add_slide_number(slide, current_slide, total_slides)
        
        logger.info(f"スライド {current_slide}/{total_slides} を作成: {title}")
    
    def _apply_background(self, slide) -> None:
        """スライドに背景画像を適用する
        
        Args:
            slide: スライドオブジェクト
        """
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)  # 白背景
        
        try:
            # 背景画像を全面に設定
            slide.shapes.add_picture(
                self.background_path,
                0, 0,
                width=self.prs.slide_width,
                height=self.prs.slide_height
            )
        except Exception as e:
            logger.error(f"背景画像の適用に失敗: {e}")
    
    def _add_logo(self, slide) -> None:
        """スライドにロゴを追加する
        
        Args:
            slide: スライドオブジェクト
        """
        try:
            # ロゴを右上に配置
            logo_width = Inches(1.2)  # ロゴサイズを調整 (1.5→1.2)
            logo = slide.shapes.add_picture(
                self.logo_path,
                self.prs.slide_width - logo_width - Inches(0.4),  # 右マージン (0.25→0.4)
                Inches(0.4),  # 上マージン (0.25→0.4)
                width=logo_width
            )
            if self.verbose:
                logger.debug(f"ロゴを追加: {logo.width} x {logo.height}")
        except Exception as e:
            logger.error(f"ロゴの追加に失敗: {e}")
    
    def _add_title(self, slide, title: str) -> None:
        """スライドにタイトルを追加する
        
        Args:
            slide: スライドオブジェクト
            title: タイトルテキスト
        """
        title_box = slide.shapes.add_textbox(
            Inches(1),  # 左マージン
            Inches(0.8),  # 上マージン（1→0.8に減少）
            self.prs.slide_width - Inches(2),  # 幅（両側マージン1インチずつ）
            Inches(1.3)  # 高さ（1.5→1.3に調整）
        )
        
        title_frame = title_box.text_frame
        title_frame.text = title
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # タイトル下の余白を追加
        title_frame.paragraphs[0].space_after = Pt(12)
        
        title_run = title_frame.paragraphs[0].runs[0]
        title_run.font.size = Pt(32)  # タイトルサイズ (36pt→32ptへ)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(0, 0, 0)  # 黒色
        title_run.font.name = "メイリオ"
        title_run.font.name_ascii = "Arial"  # フォールバック
    
    def _add_content(self, slide, content_ast: List[Dict[str, Any]]) -> None:
        """スライドにMarkdownコンテンツを追加する
        
        Args:
            slide: スライドオブジェクト
            content_ast: コンテンツのAST
        """
        # デバッグログ - AST構造を詳細に出力
        import json
        if logger.isEnabledFor(logging.INFO):
            logger.info(f"コンテンツAST: {json.dumps(content_ast, indent=2, ensure_ascii=False)}")
        
        # コンテンツ領域の定義 - マージン改善
        content_box = slide.shapes.add_textbox(
            Inches(1.0),  # 左マージン（0.8→1.0に増加）
            Inches(2.5),  # タイトル下から（2.8→2.5に減少でタイトルに近く）
            self.prs.slide_width - Inches(2.0),  # 幅（両側マージン1.0インチずつ）
            self.prs.slide_height - Inches(3.0)  # 高さ（下部マージン考慮、3.2→3.0で少し拡大）
        )
        
        text_frame = content_box.text_frame
        text_frame.word_wrap = True
        text_frame.auto_size = True  # テキストに合わせて自動調整
        
        # 段落間のスペーシング設定
        text_frame.paragraphs[0].space_after = Pt(6)  # 段落後の間隔
        
        # 最初の段落をクリア
        if text_frame.paragraphs:
            first_paragraph = text_frame.paragraphs[0]
            first_paragraph.text = ""
            current_paragraph = first_paragraph
        else:
            current_paragraph = text_frame.add_paragraph()
        
        # より直接的なアプローチでASTを処理
        for node in content_ast:
            self._process_node_direct(node, text_frame)
    
    def _process_node_direct(self, node: Dict[str, Any], text_frame) -> None:
        """ASTノードを直接処理して段落に変換する
        
        Args:
            node: ASTノード
            text_frame: テキストフレーム
        """
        node_type = node.get("type", "")
        
        # ヘッディング処理
        if node_type == "heading":
            level = node.get("attrs", {}).get("level", 2)
            if level > 1:  # H1はスライドタイトルで使用済み
                paragraph = text_frame.add_paragraph()
                text = self._node_to_text(node)
                paragraph.text = text
                
                # 段落間のスペーシング
                paragraph.space_before = Pt(12)  # 見出し前の間隔
                paragraph.space_after = Pt(6)    # 見出し後の間隔
                
                # 見出しスタイル設定
                for run in paragraph.runs:
                    run.font.size = Pt(22 - (level - 2) * 2)  # H2: 22pt, H3: 20pt, ... (前は24pt)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.font.name = "メイリオ"
                    run.font.name_ascii = "Arial"
        
        # 段落処理
        elif node_type == "paragraph":
            paragraph = text_frame.add_paragraph()
            text = self._node_to_text(node)
            paragraph.text = text
            
            # 段落間のスペーシング
            paragraph.space_after = Pt(6)  # 段落後の間隔
            
            # 段落スタイル設定
            for run in paragraph.runs:
                run.font.size = Pt(16)  # 通常テキスト
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.font.name = "メイリオ"
                run.font.name_ascii = "Arial"
        
        # リスト処理（最も重要な部分）
        elif node_type == "list":
            self._process_list_direct(node, text_frame)
        
        # コードブロック処理
        elif node_type == "block_code":
            code_para = text_frame.add_paragraph()
            code_text = node.get("raw", node.get("text", ""))
            code_para.text = code_text
            
            # コードブロックの間隔
            code_para.space_before = Pt(8)
            code_para.space_after = Pt(8)
            
            for run in code_para.runs:
                run.font.size = Pt(14)  # コードを小さく (16→14pt)
                run.font.name = "Consolas"
                run.font.name_ascii = "Consolas"
                run.font.color.rgb = RGBColor(30, 30, 30)  # 少し明るい黒
    
    def _process_list_direct(self, node: Dict[str, Any], text_frame) -> None:
        """リストを直接処理する
        
        Args:
            node: リストノード
            text_frame: テキストフレーム
        """
        ordered = node.get("ordered", False)
        list_items = node.get("children", [])
        
        # リスト開始前の余白
        if text_frame.paragraphs and text_frame.paragraphs[-1].text:
            space_para = text_frame.add_paragraph()
            space_para.space_before = Pt(4)
            space_para.space_after = Pt(0)
        
        item_number = 0
        for item in list_items:
            if item.get("type") != "list_item":
                continue
                
            item_number += 1
            
            # 各箇条書き項目に対する新しい段落
            para = text_frame.add_paragraph()
            
            # 箇条書き間のスペーシング
            para.space_before = Pt(0)
            para.space_after = Pt(2)  # 項目間の間隔を小さく
            
            # 箇条書きマーカーを作成
            if ordered:
                bullet = f"{item_number}. "
            else:
                bullet = "• "
            
            # 箇条書きの内容を抽出
            item_text = self._list_item_to_text(item)
            
            # 段落にテキストを設定
            para.text = bullet + item_text
            
            # スタイルを適用
            for run in para.runs:
                run.font.size = Pt(15)  # 箇条書きサイズ (14→15pt)
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.font.name = "メイリオ"
                run.font.name_ascii = "Arial"
                
                # 強調表示（**text**）
                if "**" in run.text:
                    parts = run.text.split("**")
                    new_text = ""
                    for i, part in enumerate(parts):
                        if i % 2 == 1:  # 強調部分
                            run.font.bold = True
                        new_text += part
                    run.text = new_text
            
            # ネストされたリストを処理
            self._process_nested_lists(item, text_frame)
        
        # リスト終了後の余白
        if list_items:
            space_para = text_frame.add_paragraph()
            space_para.space_before = Pt(2)
            space_para.space_after = Pt(0)
    
    def _process_nested_lists(self, parent_item: Dict[str, Any], text_frame) -> None:
        """ネストされたリストを処理する
        
        Args:
            parent_item: 親リスト項目
            text_frame: テキストフレーム
        """
        for child in parent_item.get("children", []):
            if child.get("type") == "list":
                # インデントを適用したサブリスト
                sub_items = child.get("children", [])
                ordered = child.get("ordered", False)
                
                for i, sub_item in enumerate(sub_items, 1):
                    if sub_item.get("type") != "list_item":
                        continue
                    
                    # サブ項目用の新しい段落
                    sub_para = text_frame.add_paragraph()
                    
                    # サブリスト項目の間隔
                    sub_para.space_before = Pt(1)
                    sub_para.space_after = Pt(1)
                    
                    # インデントを設定
                    sub_para.level = 1
                    
                    # サブリストのマーカー
                    if ordered:
                        sub_bullet = f"    {i}. "
                    else:
                        sub_bullet = "    ◦ "
                    
                    # サブアイテムのテキスト
                    sub_text = self._list_item_to_text(sub_item)
                    
                    # 段落にテキストを設定
                    sub_para.text = sub_bullet + sub_text
                    
                    # スタイルを適用
                    for run in sub_para.runs:
                        run.font.size = Pt(13)  # サブリストサイズ (12→13pt)
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.name = "メイリオ"
                        run.font.name_ascii = "Arial"
                        
                        # 強調表示を適用
                        if "**" in run.text:
                            run.text = run.text.replace("**", "")
                            run.font.bold = True
    
    def _node_to_text(self, node: Dict[str, Any]) -> str:
        """ノードからテキストを抽出する
        
        Args:
            node: ASTノード
            
        Returns:
            str: 抽出されたテキスト
        """
        if "children" not in node:
            return node.get("raw", node.get("text", ""))
        
        texts = []
        for child in node["children"]:
            child_type = child.get("type", "")
            
            if child_type == "text":
                texts.append(child.get("raw", child.get("text", "")))
            elif child_type == "strong":
                texts.append("**" + self._node_to_text(child) + "**")
            elif child_type == "emphasis":
                texts.append("*" + self._node_to_text(child) + "*")
            elif child_type == "code":
                texts.append("`" + child.get("raw", child.get("text", "")) + "`")
            elif "children" in child:
                texts.append(self._node_to_text(child))
        
        return "".join(texts)
    
    def _list_item_to_text(self, item: Dict[str, Any]) -> str:
        """リスト項目からテキストを抽出する
        
        Args:
            item: リスト項目ノード
            
        Returns:
            str: 抽出されたテキスト
        """
        texts = []
        if logger.isEnabledFor(logging.INFO):
            logger.info(f"処理中のリスト項目: {item}")
        
        for child in item.get("children", []):
            if logger.isEnabledFor(logging.INFO):
                logger.info(f"子要素タイプ: {child.get('type')}")
            
            if child.get("type") == "list":
                # ネストされたリストはスキップ
                continue
                
            elif child.get("type") == "block_text":
                # block_text要素を処理（mistune 3.1.3の特殊構造）
                if logger.isEnabledFor(logging.INFO):
                    logger.info(f"block_text要素処理: {child}")
                
                # 直接block_textの子要素または文字列を抽出
                if "children" in child:
                    for block_child in child.get("children", []):
                        if block_child.get("type") == "text":
                            text = block_child.get("raw", block_child.get("text", ""))
                            texts.append(text)
                            if logger.isEnabledFor(logging.INFO):
                                logger.info(f"抽出されたテキスト (block_text内): {text}")
                        elif block_child.get("type") == "strong":
                            strong_text = "**" + self._node_to_text(block_child) + "**"
                            texts.append(strong_text)
                        elif block_child.get("type") == "emphasis":
                            emph_text = "*" + self._node_to_text(block_child) + "*"
                            texts.append(emph_text)
                        # 他のインラインタイプも考慮
                        elif "children" in block_child:
                            texts.append(self._node_to_text(block_child))
                else:
                    # 子要素がない場合は直接テキストを抽出
                    text = child.get("raw", child.get("text", ""))
                    if text:
                        texts.append(text)
                
            elif child.get("type") == "paragraph":
                # 段落から各要素を抽出
                para_text = []
                
                for para_child in child.get("children", []):
                    child_type = para_child.get("type", "")
                    
                    if child_type == "text":
                        para_text.append(para_child.get("raw", para_child.get("text", "")))
                    elif child_type == "strong":
                        # 強調表示はマーカーつきで抽出（後で処理）
                        strong_children = para_child.get("children", [])
                        strong_text = "".join(c.get("raw", c.get("text", "")) for c in strong_children)
                        para_text.append(f"**{strong_text}**")
                    elif child_type == "emphasis":
                        # 斜体もマーカーつきで抽出
                        emph_children = para_child.get("children", [])
                        emph_text = "".join(c.get("raw", c.get("text", "")) for c in emph_children)
                        para_text.append(f"*{emph_text}*")
                    elif "children" in para_child:
                        # その他の複合要素
                        para_text.append(self._node_to_text(para_child))
                
                texts.append("".join(para_text))
                
            elif child.get("type") == "text":
                texts.append(child.get("raw", child.get("text", "")))
            
            elif child.get("type") == "strong":
                strong_text = self._node_to_text(child)
                texts.append(f"**{strong_text}**")
                
            elif child.get("type") == "emphasis":
                emph_text = self._node_to_text(child)
                texts.append(f"*{emph_text}*")
            
            # その他のノードタイプも対応
            elif "children" in child:
                texts.append(self._node_to_text(child))
        
        result = "".join(texts)
        if logger.isEnabledFor(logging.INFO):
            logger.info(f"リスト項目から抽出された最終テキスト: {result}")
        return result
    
    def _add_slide_number(self, slide, current: int, total: int) -> None:
        """スライド番号を追加する
        
        Args:
            slide: スライドオブジェクト
            current: 現在のスライド番号
            total: スライドの総数
        """
        # フッター領域にスライド番号を配置
        number_box = slide.shapes.add_textbox(
            self.prs.slide_width - Inches(1.5),  # 右マージン（2→1.5に調整）
            self.prs.slide_height - Inches(0.6),  # 下部に配置（0.8→0.6に調整）
            Inches(1.0),  # 幅
            Inches(0.3)   # 高さ
        )
        
        number_frame = number_box.text_frame
        number_frame.text = f"{current}/{total}"
        number_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        
        number_run = number_frame.paragraphs[0].runs[0]
        number_run.font.size = Pt(12)  # フォントサイズ小さく（14→12pt）
        number_run.font.color.rgb = RGBColor(80, 80, 80)  # グレー
        number_run.font.name = "メイリオ"
        number_run.font.name_ascii = "Arial"
    
    def build_presentation(self, slides_data: List[Dict[str, Any]], output_path: str) -> None:
        """スライドデータからプレゼンテーションを構築し保存する
        
        Args:
            slides_data: スライドデータのリスト
            output_path: 出力PPTXのパス
        """
        total_slides = len(slides_data)
        logger.info(f"{total_slides}枚のスライドを作成します")
        
        for slide_data in slides_data:
            self.create_slide(slide_data, total_slides)
        
        # 保存
        try:
            self.prs.save(output_path)
            logger.info(f"プレゼンテーションを保存しました: {output_path}")
        except Exception as e:
            logger.error(f"プレゼンテーション保存エラー: {e}")
            raise 