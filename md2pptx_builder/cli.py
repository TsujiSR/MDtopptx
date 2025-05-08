"""
md2pptx-builder - CLI interface
"""

import os
import sys
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, Optional

from md2pptx_builder.parser import MarkdownParser
from md2pptx_builder.builder import PPTXBuilder
from md2pptx_builder.utils import setup_logging, is_valid_image, is_valid_markdown

logger = logging.getLogger(__name__)

def parse_arguments() -> Dict[str, Any]:
    """コマンドライン引数をパースする
    
    Returns:
        Dict[str, Any]: パースされた引数
    """
    parser = argparse.ArgumentParser(
        description="Markdownファイルから背景画像とロゴを重ねたPowerPointを生成します",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument(
        "input_md",
        help="入力Markdownファイルパス"
    )
    
    parser.add_argument(
        "-b", "--background",
        required=True,
        help="背景画像ファイルパス（.png/.jpg）"
    )
    
    parser.add_argument(
        "-l", "--logo",
        required=True,
        help="ロゴ画像ファイルパス"
    )
    
    parser.add_argument(
        "-t", "--template",
        help="テンプレートPPTXファイルパス"
    )
    
    parser.add_argument(
        "-o", "--output",
        default="output.pptx",
        help="出力PPTXファイルパス"
    )
    
    parser.add_argument(
        "--pagebreak",
        default="---",
        help="Markdownスライド区切り文字"
    )
    
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="解析のみを行い、ファイルは書き出しません"
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="詳細ログを出力します"
    )
    
    return vars(parser.parse_args())

def validate_inputs(args: Dict[str, Any]) -> bool:
    """入力ファイルを検証する
    
    Args:
        args: パースされた引数
        
    Returns:
        bool: 検証結果
    """
    # 入力Markdownファイル
    if not os.path.exists(args["input_md"]):
        logger.error(f"入力Markdownファイルが見つかりません: {args['input_md']}")
        return False
    
    # 背景画像
    if not os.path.exists(args["background"]):
        logger.error(f"背景画像ファイルが見つかりません: {args['background']}")
        return False
    
    if not is_valid_image(args["background"]):
        logger.error(f"無効な背景画像ファイルです: {args['background']}")
        return False
    
    # ロゴ画像
    if not os.path.exists(args["logo"]):
        logger.error(f"ロゴ画像ファイルが見つかりません: {args['logo']}")
        return False
    
    if not is_valid_image(args["logo"]):
        logger.error(f"無効なロゴ画像ファイルです: {args['logo']}")
        return False
    
    # テンプレート
    if args["template"] and not os.path.exists(args["template"]):
        logger.error(f"テンプレートPPTXファイルが見つかりません: {args['template']}")
        return False
    
    # 出力先ディレクトリ
    output_dir = os.path.dirname(args["output"])
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            logger.info(f"出力ディレクトリを作成しました: {output_dir}")
        except Exception as e:
            logger.error(f"出力ディレクトリ作成エラー: {e}")
            return False
    
    return True

def run(args: Dict[str, Any]) -> int:
    """メイン処理を実行する
    
    Args:
        args: パースされた引数
        
    Returns:
        int: 終了コード
    """
    # 入力ファイル検証
    if not validate_inputs(args):
        return 1
    
    try:
        # Markdownパーサー初期化
        parser = MarkdownParser(pagebreak=args["pagebreak"])
        
        # Markdownファイルを処理
        slides_data = parser.process_markdown_file(args["input_md"])
        
        # スライドが存在するか確認
        if not slides_data:
            logger.warning("変換可能なスライドがありません")
            return 0
        
        logger.info(f"{len(slides_data)}枚のスライドを検出しました")
        
        # ドライランの場合はここで終了
        if args["dry_run"]:
            logger.info("ドライラン: PPTXファイルは生成されません")
            return 0
        
        # PPTXビルダー初期化
        builder = PPTXBuilder(
            background_path=args["background"],
            logo_path=args["logo"],
            template_path=args["template"],
            verbose=args["verbose"]
        )
        
        # プレゼンテーション構築
        builder.build_presentation(slides_data, args["output"])
        
        logger.info(f"変換が完了しました: {args['output']}")
        return 0
        
    except Exception as e:
        logger.error(f"変換エラー: {e}")
        if args["verbose"]:
            import traceback
            logger.error(traceback.format_exc())
        return 1

def main() -> None:
    """CLIのエントリーポイント"""
    # 引数解析
    args = parse_arguments()
    
    # ロギング設定
    setup_logging(args["verbose"])
    
    # 実行
    sys.exit(run(args))

if __name__ == "__main__":
    main() 