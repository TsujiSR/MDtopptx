"""
md2pptx-builder - Utility functions
"""

import os
import tempfile
import logging
from pathlib import Path
from typing import Optional, Union, Tuple

from PIL import Image

# ロギング設定
logger = logging.getLogger(__name__)

def setup_logging(verbose: bool = False) -> None:
    """ロギングを設定する
    
    Args:
        verbose: 詳細ログを表示するかどうか
    """
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    
    # デバッグモードでは外部ライブラリのログレベルを調整
    if verbose:
        # mistune、python-pptxなどの外部ライブラリのログレベルを上げる
        logging.getLogger('mistune').setLevel(logging.INFO)
        logging.getLogger('pptx').setLevel(logging.INFO)
    else:
        # 外部ライブラリのログレベルを下げる
        logging.getLogger('mistune').setLevel(logging.WARNING)
        logging.getLogger('pptx').setLevel(logging.WARNING)

def is_valid_image(file_path: Union[str, Path]) -> bool:
    """有効な画像ファイルかどうかを確認する
    
    Args:
        file_path: 画像ファイルパス
        
    Returns:
        bool: 有効な画像かどうか
    """
    try:
        Image.open(file_path).verify()
        return True
    except Exception as e:
        logger.error(f"無効な画像ファイル: {file_path}, エラー: {e}")
        return False

def get_image_dimensions(file_path: Union[str, Path]) -> Tuple[int, int]:
    """画像のサイズを取得する
    
    Args:
        file_path: 画像ファイルパス
        
    Returns:
        Tuple[int, int]: 幅と高さのタプル
    """
    with Image.open(file_path) as img:
        return img.size

def create_temp_file(content: str, suffix: str = '.md') -> str:
    """一時ファイルを作成する
    
    Args:
        content: ファイル内容
        suffix: ファイル拡張子
        
    Returns:
        str: 一時ファイルパス
    """
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        temp_file.write(content.encode('utf-8'))
        temp_file.flush()
        return temp_file.name
    finally:
        temp_file.close()

def clean_temp_files(file_paths: list) -> None:
    """一時ファイルを削除する
    
    Args:
        file_paths: 削除する一時ファイルのリスト
    """
    for path in file_paths:
        try:
            if os.path.exists(path):
                os.unlink(path)
                logger.debug(f"一時ファイルを削除しました: {path}")
        except Exception as e:
            logger.error(f"一時ファイル削除エラー: {path}, エラー: {e}")

def is_valid_markdown(content: str) -> bool:
    """有効なMarkdownコンテンツかどうかを確認する
    
    Args:
        content: Markdownテキスト
        
    Returns:
        bool: 有効かどうか
    """
    # 簡易チェック - 空でないかどうか
    if not content or not content.strip():
        return False
    return True

def get_file_extension(file_path: Union[str, Path]) -> str:
    """ファイル拡張子を取得する
    
    Args:
        file_path: ファイルパス
        
    Returns:
        str: 拡張子（ドット付き）
    """
    return os.path.splitext(str(file_path))[1].lower() 