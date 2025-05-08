"""
md2pptx-builder - Markdown parser
"""

import re
import logging
from typing import List, Dict, Any, Tuple
import json

import mistune

logger = logging.getLogger(__name__)

class MarkdownParser:
    """Markdownをパースし、スライドに分割するクラス"""
    
    def __init__(self, pagebreak: str = "---"):
        """
        Args:
            pagebreak: スライド区切り文字
        """
        self.pagebreak = pagebreak
        self.parser = mistune.create_markdown(renderer='ast')
    
    def split_to_slides(self, markdown_content: str) -> List[str]:
        """Markdownコンテンツをスライドごとに分割する
        
        Args:
            markdown_content: Markdownテキスト
            
        Returns:
            List[str]: スライドごとに分割されたMarkdownテキストのリスト
        """
        # <!-- pagebreak --> 形式も対応
        alt_pagebreak = "<!-- pagebreak -->"
        
        # 両方の区切り文字を統一形式に変換
        normalized_content = markdown_content.replace(alt_pagebreak, self.pagebreak)
        
        # 区切り文字で分割
        pattern = f"(?:^|\n){re.escape(self.pagebreak)}(?:\n|$)"
        slides = re.split(pattern, normalized_content)
        
        # 空のスライドを除去
        slides = [slide.strip() for slide in slides if slide.strip()]
        
        logger.info(f"{len(slides)}枚のスライドに分割しました")
        return slides
    
    def parse_slide(self, slide_content: str) -> List[Dict[str, Any]]:
        """スライドのMarkdownをパースしてAST（抽象構文木）に変換する
        
        Args:
            slide_content: スライドのMarkdownテキスト
            
        Returns:
            List[Dict[str, Any]]: ASTノードのリスト
        """
        try:
            ast = self.parser(slide_content)
            # デバッグ用：ASTログ
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(f"ASTパース結果: {json.dumps(ast[:3], ensure_ascii=False)[:200]}...")
            return ast
        except Exception as e:
            logger.error(f"Markdownパースエラー: {e}")
            # 最低限、テキストとして扱えるよう空のドキュメントを返す
            return [{"type": "paragraph", "children": [{"type": "text", "text": slide_content}]}]
    
    def get_slide_title(self, ast: List[Dict[str, Any]]) -> Tuple[str, List[Dict[str, Any]]]:
        """スライドからタイトル（h1）を抽出する
        
        Args:
            ast: スライドのAST
            
        Returns:
            Tuple[str, List[Dict[str, Any]]]: タイトルと残りのコンテンツのAST
        """
        title = ""
        remaining_ast = []
        
        for node in ast:
            if not title and node["type"] == "heading" and node.get("attrs", {}).get("level") == 1:
                # H1を見つけたらタイトルとして使用
                title_parts = []
                for child in node["children"]:
                    if child["type"] == "text":
                        # mistune 3.1.3では主にrawキーが使われる
                        text = child.get("raw", child.get("text", ""))
                        title_parts.append(text)
                title = "".join(title_parts)
            else:
                remaining_ast.append(node)
        
        return title, remaining_ast
    
    def process_markdown_file(self, file_path: str) -> List[Dict[str, Any]]:
        """Markdownファイルを処理し、スライド情報のリストを返す
        
        Args:
            file_path: Markdownファイルパス
            
        Returns:
            List[Dict[str, Any]]: スライド情報（タイトル、コンテンツのAST）のリスト
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            return self.process_markdown_content(content)
            
        except Exception as e:
            logger.error(f"Markdownファイル処理エラー: {e}")
            raise
    
    def process_markdown_content(self, content: str) -> List[Dict[str, Any]]:
        """Markdownコンテンツを処理し、スライド情報のリストを返す
        
        Args:
            content: Markdownテキスト
            
        Returns:
            List[Dict[str, Any]]: スライド情報（タイトル、コンテンツのAST）のリスト
        """
        slides = []
        slide_texts = self.split_to_slides(content)
        
        for index, slide_text in enumerate(slide_texts):
            ast = self.parse_slide(slide_text)
            # デバッグ用：ASTをログ出力
            self.debug_ast(ast, f"スライド{index+1}")
            title, content_ast = self.get_slide_title(ast)
            
            if not title:
                title = f"スライド {index + 1}"
                
            slides.append({
                "title": title,
                "content": content_ast,
                "index": index,
                "raw_text": slide_text
            })
        
        return slides
        
    def debug_ast(self, ast: List[Dict[str, Any]], prefix: str = ""):
        """ASTをデバッグのためにログ出力する
        
        Args:
            ast: ASTノード
            prefix: ログプレフィックス
        """
        if logger.isEnabledFor(logging.DEBUG):
            logger.debug(f"{prefix} AST構造: {json.dumps(ast, ensure_ascii=False, indent=2)}") 