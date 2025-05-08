"""
md2pptx-builder - CLIモジュールのテスト
"""

import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock
from pathlib import Path

from md2pptx_builder.cli import validate_inputs, run

class TestCLI(unittest.TestCase):
    """CLIモジュールのテスト"""
    
    def setUp(self):
        """テスト開始前の準備"""
        # テスト用のダミーファイル作成
        self.temp_md = tempfile.NamedTemporaryFile(delete=False, suffix=".md")
        self.temp_md.write(b"# Test\n\nContent\n\n---\n\n# Slide 2")
        self.temp_md.close()
        
        # ダミー画像ファイル作成
        self.temp_bg = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        self.temp_bg.close()
        
        self.temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        self.temp_logo.close()
        
        # 出力ディレクトリ
        self.output_dir = tempfile.mkdtemp()
        self.output_file = os.path.join(self.output_dir, "output.pptx")
    
    def tearDown(self):
        """テスト終了後のクリーンアップ"""
        # 一時ファイルの削除
        for temp_file in [self.temp_md, self.temp_bg, self.temp_logo]:
            if os.path.exists(temp_file.name):
                os.unlink(temp_file.name)
        
        if os.path.exists(self.output_file):
            os.unlink(self.output_file)
    
    @patch("md2pptx_builder.cli.is_valid_image")
    def test_validate_inputs_valid(self, mock_is_valid_image):
        """有効な入力の検証テスト"""
        # モックの設定
        mock_is_valid_image.return_value = True
        
        # テスト用の引数
        args = {
            "input_md": self.temp_md.name,
            "background": self.temp_bg.name,
            "logo": self.temp_logo.name,
            "template": None,
            "output": self.output_file,
            "verbose": False,
            "pagebreak": "---",
            "dry_run": False
        }
        
        # 検証実行
        result = validate_inputs(args)
        self.assertTrue(result, "有効な入力が検証を通過すべき")
    
    @patch("md2pptx_builder.cli.is_valid_image")
    def test_validate_inputs_missing_md(self, mock_is_valid_image):
        """入力Markdownファイルが存在しない場合のテスト"""
        # モックの設定
        mock_is_valid_image.return_value = True
        
        # 存在しないファイルパス
        non_existent_file = "/path/to/nonexistent.md"
        
        # テスト用の引数
        args = {
            "input_md": non_existent_file,
            "background": self.temp_bg.name,
            "logo": self.temp_logo.name,
            "template": None,
            "output": self.output_file,
            "verbose": False,
            "pagebreak": "---",
            "dry_run": False
        }
        
        # 検証実行
        result = validate_inputs(args)
        self.assertFalse(result, "存在しないMarkdownファイルは検証に失敗すべき")
    
    @patch("md2pptx_builder.cli.validate_inputs")
    @patch("md2pptx_builder.cli.MarkdownParser")
    @patch("md2pptx_builder.cli.PPTXBuilder")
    def test_run_successful(self, mock_builder, mock_parser, mock_validate):
        """正常実行のテスト"""
        # モックの設定
        mock_validate.return_value = True
        
        mock_parser_instance = MagicMock()
        mock_parser_instance.process_markdown_file.return_value = [
            {"title": "Slide 1", "content": [], "index": 0, "raw_text": "# Slide 1"}
        ]
        mock_parser.return_value = mock_parser_instance
        
        mock_builder_instance = MagicMock()
        mock_builder.return_value = mock_builder_instance
        
        # テスト用の引数
        args = {
            "input_md": self.temp_md.name,
            "background": self.temp_bg.name,
            "logo": self.temp_logo.name,
            "template": None,
            "output": self.output_file,
            "verbose": False,
            "pagebreak": "---",
            "dry_run": False
        }
        
        # 実行
        result = run(args)
        
        # 検証
        self.assertEqual(result, 0, "成功した実行は0を返すべき")
        mock_validate.assert_called_once_with(args)
        mock_parser_instance.process_markdown_file.assert_called_once_with(self.temp_md.name)
        mock_builder.assert_called_once()
        mock_builder_instance.build_presentation.assert_called_once()
    
    @patch("md2pptx_builder.cli.validate_inputs")
    def test_run_validation_failed(self, mock_validate):
        """入力検証失敗時のテスト"""
        # モックの設定
        mock_validate.return_value = False
        
        # テスト用の引数
        args = {
            "input_md": "/nonexistent.md",
            "background": self.temp_bg.name,
            "logo": self.temp_logo.name,
            "template": None,
            "output": self.output_file,
            "verbose": False,
            "pagebreak": "---",
            "dry_run": False
        }
        
        # 実行
        result = run(args)
        
        # 検証
        self.assertEqual(result, 1, "検証失敗時は1を返すべき")
        mock_validate.assert_called_once_with(args)
    
    @patch("md2pptx_builder.cli.validate_inputs")
    @patch("md2pptx_builder.cli.MarkdownParser")
    def test_run_dry_run(self, mock_parser, mock_validate):
        """ドライラン実行のテスト"""
        # モックの設定
        mock_validate.return_value = True
        
        mock_parser_instance = MagicMock()
        mock_parser_instance.process_markdown_file.return_value = [
            {"title": "Slide 1", "content": [], "index": 0, "raw_text": "# Slide 1"}
        ]
        mock_parser.return_value = mock_parser_instance
        
        # テスト用の引数
        args = {
            "input_md": self.temp_md.name,
            "background": self.temp_bg.name,
            "logo": self.temp_logo.name,
            "template": None,
            "output": self.output_file,
            "verbose": False,
            "pagebreak": "---",
            "dry_run": True  # ドライラン
        }
        
        # 実行
        result = run(args)
        
        # 検証
        self.assertEqual(result, 0, "ドライランは0を返すべき")
        mock_validate.assert_called_once_with(args)
        mock_parser_instance.process_markdown_file.assert_called_once_with(self.temp_md.name)
        # PPTXビルダーは呼ばれないはず

if __name__ == "__main__":
    unittest.main() 