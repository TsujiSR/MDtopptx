"""
md2pptx-builder - テスト用のfixtures
"""

import os
import pytest
import tempfile
from pathlib import Path


@pytest.fixture
def sample_markdown():
    """サンプルのMarkdownテキストを提供するfixture"""
    return """# テストスライド1

テスト段落です。
**太字**と*斜体*をサポートします。

- リスト項目1
- リスト項目2
  - サブリスト項目

---

# テストスライド2

## 見出し2

```python
def hello_world():
    print("Hello, World!")
```
"""


@pytest.fixture
def temp_markdown_file(sample_markdown):
    """一時的なMarkdownファイルを作成するfixture"""
    with tempfile.NamedTemporaryFile(suffix=".md", delete=False, mode="w", encoding="utf-8") as f:
        f.write(sample_markdown)
        temp_path = f.name
    
    yield temp_path
    
    # テスト後にファイルを削除
    if os.path.exists(temp_path):
        os.unlink(temp_path)


@pytest.fixture
def temp_output_pptx():
    """一時的な出力PPTXファイルパスを提供するfixture"""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        temp_path = f.name
    
    # 先に削除してからパスだけ返す
    os.unlink(temp_path)
    
    yield temp_path
    
    # テスト後にファイルを削除（存在する場合）
    if os.path.exists(temp_path):
        os.unlink(temp_path)


@pytest.fixture
def sample_images_dir():
    """サンプル画像のディレクトリを返すfixture"""
    # プロジェクトルートからの相対パス
    samples_dir = Path(__file__).parent.parent / "samples"
    return samples_dir 