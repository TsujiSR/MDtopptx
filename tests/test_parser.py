"""
md2pptx-builder - Markdownパーサーのテスト
"""

import os
import tempfile
import unittest
from pathlib import Path

from md2pptx_builder.parser import MarkdownParser

class TestMarkdownParser(unittest.TestCase):
    """Markdownパーサーのテスト"""
    
    def setUp(self):
        """テスト開始前の準備"""
        self.parser = MarkdownParser()
        
        # テスト用Markdownテキスト
        self.test_md_content = """# スライド1タイトル

これは最初のスライドです。

- 箇条書き1
- 箇条書き2

---

# スライド2タイトル

## 見出し2

これは2つ目のスライドです。

```python
print("Hello, World!")
```

<!-- pagebreak -->

# スライド3タイトル

最後のスライドです。
"""
        
        # 一時ファイル作成
        self.temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".md")
        self.temp_file.write(self.test_md_content.encode('utf-8'))
        self.temp_file.close()
    
    def tearDown(self):
        """テスト終了後のクリーンアップ"""
        if os.path.exists(self.temp_file.name):
            os.unlink(self.temp_file.name)
    
    def test_split_to_slides(self):
        """スライド分割のテスト"""
        slides = self.parser.split_to_slides(self.test_md_content)
        
        # 期待されるスライド数
        self.assertEqual(len(slides), 3, "スライド数が3でなければならない")
        
        # 最初のスライドが正しく分割されているか
        self.assertIn("スライド1タイトル", slides[0])
        self.assertIn("これは最初のスライドです。", slides[0])
        
        # 2つ目のスライドが正しく分割されているか
        self.assertIn("スライド2タイトル", slides[1])
        self.assertIn("見出し2", slides[1])
        
        # 3つ目のスライドが正しく分割されているか
        self.assertIn("スライド3タイトル", slides[2])
        self.assertIn("最後のスライドです。", slides[2])
    
    def test_parse_slide(self):
        """スライドパースのテスト"""
        slide_content = "# タイトル\n\nこれはテスト段落です。"
        ast = self.parser.parse_slide(slide_content)
        
        # ASTが正しく生成されているか
        self.assertTrue(isinstance(ast, list))
        self.assertTrue(len(ast) > 0)
        
        # 見出しノードの確認
        heading_node = ast[0]
        self.assertEqual(heading_node["type"], "heading")
        self.assertEqual(heading_node["level"], 1)
        
        # 段落ノードの確認
        paragraph_node = ast[1]
        self.assertEqual(paragraph_node["type"], "paragraph")
    
    def test_get_slide_title(self):
        """スライドタイトル抽出のテスト"""
        slide_content = "# スライドタイトル\n\nコンテンツ"
        ast = self.parser.parse_slide(slide_content)
        title, remaining = self.parser.get_slide_title(ast)
        
        # タイトルが正しく抽出されているか
        self.assertEqual(title, "スライドタイトル")
        
        # 残りのコンテンツが正しく処理されているか
        self.assertEqual(len(remaining), 1)
        self.assertEqual(remaining[0]["type"], "paragraph")
    
    def test_process_markdown_file(self):
        """Markdownファイル処理のテスト"""
        slides_data = self.parser.process_markdown_file(self.temp_file.name)
        
        # スライドデータの確認
        self.assertEqual(len(slides_data), 3)
        
        # スライドタイトルの確認
        self.assertEqual(slides_data[0]["title"], "スライド1タイトル")
        self.assertEqual(slides_data[1]["title"], "スライド2タイトル")
        self.assertEqual(slides_data[2]["title"], "スライド3タイトル")
        
        # インデックスの確認
        self.assertEqual(slides_data[0]["index"], 0)
        self.assertEqual(slides_data[1]["index"], 1)
        self.assertEqual(slides_data[2]["index"], 2)

if __name__ == "__main__":
    unittest.main() 