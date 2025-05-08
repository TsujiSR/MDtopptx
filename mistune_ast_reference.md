# Mistune によるMarkdown要素の変換構造

Mistune 3.xでは、Markdownテキストを抽象構文木（AST）に変換します。このドキュメントは、md2pptx-builder プロジェクトにおける各Markdown要素がMistune 3.xでどのように変換されるかを整理したものです。

## 基本構造

```json
[
  {
    "type": "要素タイプ",
    "children": [...],  // 子要素がある場合
    "attrs": {...},     // 属性が必要な場合
    "raw": "元のテキスト"  // 生のテキスト
  },
  // 他の要素...
]
```

## 見出し (Heading)

**Markdown:**
```markdown
# 見出し1
## 見出し2
```

**AST:**
```json
{
  "type": "heading",
  "attrs": {
    "level": 1  // H1～H6のレベル
  },
  "children": [
    {
      "type": "text",
      "raw": "見出し1"
    }
  ]
}
```

## 段落 (Paragraph)

**Markdown:**
```markdown
これは段落です。
```

**AST:**
```json
{
  "type": "paragraph",
  "children": [
    {
      "type": "text",
      "raw": "これは段落です。"
    }
  ]
}
```

## リスト (List)

**Markdown:**
```markdown
- 項目1
- 項目2
  - ネスト項目
```

**AST:**
```json
{
  "type": "list",
  "attrs": {
    "ordered": false,
    "depth": 0
  },
  "children": [
    {
      "type": "list_item",
      "children": [
        {
          "type": "block_text",
          "children": [
            {
              "type": "text",
              "raw": "項目1"
            }
          ]
        }
      ]
    },
    {
      "type": "list_item",
      "children": [
        {
          "type": "block_text",
          "children": [
            {
              "type": "text",
              "raw": "項目2"
            }
          ]
        },
        {
          "type": "list",
          "attrs": {
            "ordered": false,
            "depth": 1
          },
          "children": [
            // ネスト項目...
          ]
        }
      ]
    }
  ]
}
```

## コードブロック (Block Code)

**Markdown:**
````markdown
```python
def hello():
    print("Hello")
```
````

**AST:**
```json
{
  "type": "block_code",
  "attrs": {
    "info": "python"
  },
  "raw": "def hello():\n    print(\"Hello\")"
}
```

## インラインコード (CodeSpan)

**Markdown:**
```markdown
`コード`
```

**AST:**
```json
{
  "type": "codespan",
  "raw": "コード"
}
```

## 強調・太字

**Markdown:**
```markdown
*斜体* **太字**
```

**AST:**
```json
{
  "type": "paragraph",
  "children": [
    {
      "type": "emphasis",
      "children": [
        {
          "type": "text",
          "raw": "斜体"
        }
      ]
    },
    {
      "type": "text",
      "raw": " "
    },
    {
      "type": "strong",
      "children": [
        {
          "type": "text",
          "raw": "太字"
        }
      ]
    }
  ]
}
```

## リンク

**Markdown:**
```markdown
[リンクテキスト](https://example.com)
```

**AST:**
```json
{
  "type": "link",
  "attrs": {
    "url": "https://example.com"
  },
  "children": [
    {
      "type": "text",
      "raw": "リンクテキスト"
    }
  ]
}
```

## 画像

**Markdown:**
```markdown
![代替テキスト](image.jpg "タイトル")
```

**AST:**
```json
{
  "type": "image",
  "attrs": {
    "url": "image.jpg",
    "title": "タイトル",
    "alt": "代替テキスト"
  }
}
```

## 引用

**Markdown:**
```markdown
> これは引用です
```

**AST:**
```json
{
  "type": "block_quote",
  "children": [
    {
      "type": "paragraph",
      "children": [
        {
          "type": "text",
          "raw": "これは引用です"
        }
      ]
    }
  ]
}
```

## 水平線

**Markdown:**
```markdown
---
```

**AST:**
```json
{
  "type": "thematic_break"
}
```

## 特殊ケース：コロンを含むリスト項目

**Markdown:**
```markdown
- key: value
```

**AST:**
```json
{
  "type": "list_item",
  "children": [
    {
      "type": "block_text",
      "children": [
        {
          "type": "text",
          "raw": "key: value"
        }
      ]
    }
  ]
}
```

## 複合要素の例：リンク付きテキストと太字の混在

**Markdown:**
```markdown
**太字[リンク](https://example.com)と通常テキスト**
```

**AST:**
```json
{
  "type": "strong",
  "children": [
    {
      "type": "text",
      "raw": "太字"
    },
    {
      "type": "link",
      "attrs": {
        "url": "https://example.com"
      },
      "children": [
        {
          "type": "text",
          "raw": "リンク"
        }
      ]
    },
    {
      "type": "text",
      "raw": "と通常テキスト"
    }
  ]
}
```

## md2pptx-builder での実装上の注意点

1. **コードスパン処理**: 
   - コードスパンは`raw`プロパティに直接テキストが格納される
   - リスト内のコードスパンは特に注意が必要で、適切に`raw`からテキストを抽出する

2. **ネストされたリスト**: 
   - 親リスト項目の`children`に子リストが含まれる形で表現される
   - 深さ（depth）を適切に処理して、インデントレベルを管理する

3. **コロン処理**: 
   - `key: value`形式のテキストを正しく表示するために、コロンの後にスペースを挿入する特殊処理が必要

4. **空白処理**: 
   - Mistune 3.xでは空行は`blank_line`タイプで表現される
   - 適切な余白を作るために空行を段落として処理する

5. **テキスト抽出**: 
   - 複雑な要素からテキストを抽出する場合は、再帰的に`children`を探索する必要がある
   - `_node_to_text`や`_extract_list_item_text`などのヘルパーメソッドを活用する

## デバッグのヒント

実際のASTはMarkdownの構造によって複雑に変化するため、以下の手法が有効です：

1. ASTをJSONとしてログ出力して構造を確認する
   ```python
   import json
   logger.info(f"AST: {json.dumps(ast, indent=2, ensure_ascii=False)}")
   ```

2. 特定の要素タイプごとに処理を切り分け、段階的にデバッグする

3. リスト項目などの複雑な要素は、最初に単純化したケースで処理を確認する

4. パスや正規表現を組み合わせて必要なテキストを抽出する場合は、中間結果もログ出力する

5. 想定外の結果が出た場合は、ASTの構造変化を詳細に追跡する 