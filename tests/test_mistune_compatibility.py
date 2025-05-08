"""
md2pptx-builder - mistune 3.1.3互換性テスト
"""

import pytest
import mistune
from md2pptx_builder.parser import MarkdownParser


def test_mistune_version():
    """mistune 3.1.3以上のバージョンが使用されていることを確認"""
    version = mistune.__version__
    major, minor, patch = map(int, version.split('.'))
    assert major >= 3
    assert (major == 3 and minor >= 1) or major > 3


def test_ast_format():
    """mistune ASTの形式が期待通りであることを確認"""
    parser = MarkdownParser()
    markdown = "# タイトル\n\nテキスト\n\n- リスト1\n- リスト2"
    ast = parser.parse_slide(markdown)
    
    # ヘッディングノードの構造確認
    heading_node = ast[0]
    assert heading_node["type"] == "heading"
    assert heading_node.get("attrs", {}).get("level") == 1
    assert "children" in heading_node
    
    # リストノードの構造確認
    list_node = [node for node in ast if node["type"] == "list"][0]
    assert "children" in list_node
    assert len(list_node["children"]) == 2
    
    # リスト項目の構造確認
    list_item = list_node["children"][0]
    assert list_item["type"] == "list_item"
    assert "children" in list_item


def test_block_text_processing():
    """block_text要素の処理が正しいことを確認"""
    parser = MarkdownParser()
    markdown = "- リスト項目"
    ast = parser.parse_slide(markdown)
    
    # リスト構造の確認
    assert ast[0]["type"] == "list"
    list_item = ast[0]["children"][0]
    assert list_item["type"] == "list_item"
    
    # block_text要素またはその他の有効な構造があることを確認
    child = list_item["children"][0]
    assert child["type"] in ["block_text", "paragraph", "text"]


def test_title_extraction():
    """タイトル抽出が正しく機能することを確認"""
    parser = MarkdownParser()
    markdown = "# タイトル\n\n内容"
    ast = parser.parse_slide(markdown)
    
    title, content = parser.get_slide_title(ast)
    assert title == "タイトル"
    assert len(content) == 1


def test_markdown_splitting():
    """Markdownの分割が正しく機能することを確認"""
    parser = MarkdownParser()
    markdown = """# スライド1

内容1

---

# スライド2

内容2"""
    
    slides = parser.split_to_slides(markdown)
    assert len(slides) == 2
    assert "スライド1" in slides[0]
    assert "スライド2" in slides[1]


def test_html_comment_pagebreak():
    """HTMLコメント形式のページ区切りが機能することを確認"""
    parser = MarkdownParser()
    markdown = """# スライド1

内容1

<!-- pagebreak -->

# スライド2

内容2"""
    
    slides = parser.split_to_slides(markdown)
    assert len(slides) == 2
    assert "スライド1" in slides[0]
    assert "スライド2" in slides[1] 