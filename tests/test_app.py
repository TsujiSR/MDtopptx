"""
md2pptx-builder - Streamlit GUIアプリケーションのテスト
"""

import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock

from md2pptx_builder.app import save_uploaded_file, create_presentation, register_temp_file

class TestStreamlitApp(unittest.TestCase):
    """StreamlitアプリケーションのGUIテスト"""
    
    def setUp(self):
        """テスト開始前の準備"""
        # モックファイルコンテンツ
        self.mock_md_content = "# Test Slide\n\nThis is a test."
        
        # テスト用一時ファイル
        self.temp_dir = tempfile.mkdtemp()
        
        # モックアップロードファイル
        self.mock_upload_md = MagicMock()
        self.mock_upload_md.name = "test.md"
        self.mock_upload_md.getvalue.return_value = self.mock_md_content.encode('utf-8')
        
        self.mock_upload_bg = MagicMock()
        self.mock_upload_bg.name = "background.png"
        self.mock_upload_bg.getvalue.return_value = b"test image data"
        
        self.mock_upload_logo = MagicMock()
        self.mock_upload_logo.name = "logo.png"
        self.mock_upload_logo.getvalue.return_value = b"test logo data"
    
    def tearDown(self):
        """テスト終了後のクリーンアップ"""
        # 一時ファイルのクリーンアップは別途行われる
        pass
    
    @patch("streamlit.error")
    @patch("tempfile.NamedTemporaryFile")
    def test_save_uploaded_file(self, mock_temp_file, mock_st_error):
        """アップロードファイル保存のテスト"""
        # モックの設定
        mock_temp = MagicMock()
        mock_temp.name = "/tmp/test1234.md"
        mock_temp_file.return_value = mock_temp
        
        # 関数呼び出し
        result = save_uploaded_file(self.mock_upload_md)
        
        # 検証
        self.assertEqual(result, mock_temp.name)
        mock_temp.write.assert_called_once_with(self.mock_md_content.encode('utf-8'))
        mock_temp.flush.assert_called_once()
        mock_temp.close.assert_called_once()
    
    @patch("streamlit.error")
    def test_save_uploaded_file_none(self, mock_st_error):
        """Noneアップロードのテスト"""
        result = save_uploaded_file(None)
        self.assertIsNone(result)
    
    @patch("streamlit.error")
    @patch("tempfile.NamedTemporaryFile")
    def test_save_uploaded_file_error(self, mock_temp_file, mock_st_error):
        """アップロードエラーのテスト"""
        # 例外を発生させる
        mock_temp_file.side_effect = Exception("Test error")
        
        # 関数呼び出し
        result = save_uploaded_file(self.mock_upload_md)
        
        # 検証
        self.assertIsNone(result)
        mock_st_error.assert_called_once()
    
    @patch("md2pptx_builder.app.MarkdownParser")
    @patch("md2pptx_builder.app.create_temp_file")
    @patch("md2pptx_builder.app.register_temp_file")
    @patch("md2pptx_builder.app.PPTXBuilder")
    @patch("streamlit.warning")
    def test_create_presentation(self, mock_st_warning, mock_builder, 
                               mock_register, mock_create_temp, mock_parser):
        """プレゼンテーション作成のテスト"""
        # モックの設定
        temp_md = "/tmp/temp1234.md"
        mock_create_temp.return_value = temp_md
        
        mock_parser_instance = MagicMock()
        slides_data = [
            {"title": "Test Slide", "content": [], "index": 0, "raw_text": "# Test Slide"}
        ]
        mock_parser_instance.process_markdown_file.return_value = slides_data
        mock_parser.return_value = mock_parser_instance
        
        mock_builder_instance = MagicMock()
        mock_builder.return_value = mock_builder_instance
        
        # 背景とロゴのパス
        bg_path = "/tmp/bg.png"
        logo_path = "/tmp/logo.png"
        output_filename = "test_output.pptx"
        
        # 関数呼び出し
        result = create_presentation(
            md_content=self.mock_md_content,
            background_path=bg_path,
            logo_path=logo_path,
            output_filename=output_filename
        )
        
        # 検証
        self.assertIsNotNone(result)
        mock_create_temp.assert_called_once_with(self.mock_md_content)
        mock_register.assert_called()
        mock_parser_instance.process_markdown_file.assert_called_once_with(temp_md)
        mock_builder.assert_called_once_with(
            background_path=bg_path,
            logo_path=logo_path,
            template_path=None,
            verbose=False
        )
        mock_builder_instance.build_presentation.assert_called_once()
    
    @patch("md2pptx_builder.app.MarkdownParser")
    @patch("md2pptx_builder.app.create_temp_file")
    @patch("md2pptx_builder.app.register_temp_file")
    @patch("streamlit.warning")
    def test_create_presentation_no_slides(self, mock_st_warning, 
                                         mock_register, mock_create_temp, mock_parser):
        """スライドなしの場合のテスト"""
        # モックの設定
        temp_md = "/tmp/temp1234.md"
        mock_create_temp.return_value = temp_md
        
        mock_parser_instance = MagicMock()
        mock_parser_instance.process_markdown_file.return_value = []  # 空のスライドデータ
        mock_parser.return_value = mock_parser_instance
        
        # 背景とロゴのパス
        bg_path = "/tmp/bg.png"
        logo_path = "/tmp/logo.png"
        
        # 関数呼び出し
        result = create_presentation(
            md_content=self.mock_md_content,
            background_path=bg_path,
            logo_path=logo_path
        )
        
        # 検証
        self.assertIsNone(result)
        mock_st_warning.assert_called_once()
        mock_parser_instance.process_markdown_file.assert_called_once_with(temp_md)
    
    def test_register_temp_file(self):
        """一時ファイル登録のテスト"""
        # グローバル変数をリセット
        import md2pptx_builder.app
        md2pptx_builder.app.temp_files = []
        
        # テスト用ファイルパス
        test_path = "/tmp/test.txt"
        
        # 関数呼び出し
        register_temp_file(test_path)
        
        # 検証
        self.assertIn(test_path, md2pptx_builder.app.temp_files)
        self.assertEqual(len(md2pptx_builder.app.temp_files), 1)

if __name__ == "__main__":
    unittest.main() 