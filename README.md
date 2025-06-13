# Slide Agent - 自律型スライド生成システム

マークダウン形式で作成したスライド情報から、PowerPointスライドを自動生成するLangGraphベースの自律型AIエージェントです。

## 概要

このシステムは以下の処理を自動で行います：

1. **マークダウン解析**: `---`区切りでスライドを分割
2. **コンテンツ整形**: AIがプレゼンテーション用に内容を最適化
3. **テンプレート選択**: 各スライドに最適なテンプレートを自動選択
4. **コンテンツ割り当て**: テンプレートの各要素に内容を自動割り当て
5. **PowerPoint生成**: 最終的なPowerPointファイル(.pptx)を出力

## 特徴

- **自律型AI**: LangGraphベースのワークフローで全自動処理
- **テンプレート管理**: JSON形式でテンプレートを管理、随時追加可能
- **中間ファイル保存**: 処理過程を全て記録、デバッグ・検証が容易
- **柔軟な構成**: 各テンプレートはフォルダ単位で管理

## システム要件

- Python 3.8+
- OpenAI API キー

## インストール

1. リポジトリをクローン
```bash
git clone <repository-url>
cd slide-agent
```

2. 依存関係をインストール
```bash
pip install -r requirements.txt
```

3. 環境変数を設定
```bash
cp .env.sample .env
# .envファイルを編集してOpenAI API キーを設定
```

## 使用方法

### 基本的な使用方法

```bash
python main.py sample_slides.md
```

### オプション

```bash
python main.py input.md -o output.pptx --debug
```

- `-o, --output`: 出力ファイル名を指定
- `--debug`: デバッグモードで実行

## テンプレート構成

テンプレートは `templates/` ディレクトリ内で管理されます：

```
templates/
├── タイトルスライド/
│   ├── template.json
│   └── template.pptx
├── 1カラムテキスト/
│   ├── template.json
│   └── template.pptx
└── ...
```

### 利用可能なテンプレート

1. **タイトルスライド**: プレゼンテーションの冒頭用
2. **1カラムテキスト**: 標準的な本文スライド
3. **リード文+1カラムテキスト**: 要点と詳細を分けて表示
4. **2カラム（画像＋テキスト）**: 左に画像、右にテキスト
5. **エンドスライド**: 発表の締めくくり用

## マークダウン形式

入力マークダウンは `---` でスライドを区切ります：

```markdown
# タイトルスライド

発表者: 田中太郎
日付: 2024年12月

---

# 本文スライド

ここに内容を記述します。

- 箇条書き1
- 箇条書き2

---

# まとめ

ご清聴ありがとうございました
```

## 出力ファイル

### 生成されるファイル

- **PowerPointファイル**: `output/generated_slides_YYYYMMDD_HHMMSS.pptx`
- **中間ファイル**: `intermediate/` ディレクトリに処理過程を保存
  - `01_parsed_slides_*.json`: 解析されたスライド
  - `02_formatted_slides_*.json`: 整形されたコンテンツ
  - `03_template_selection_*.json`: テンプレート選択結果
  - `04_content_assignment_*.json`: コンテンツ割り当て結果

## カスタマイズ

### 新しいテンプレートの追加

1. `templates/` 内に新しいフォルダを作成
2. `template.json` でテンプレート構成を定義
3. `template.pptx` でPowerPointテンプレートを作成

### template.json の例

```json
{
  "template_name": "カスタムテンプレート",
  "description": "カスタムテンプレートの説明",
  "use_case_examples": ["利用ケース1", "利用ケース2"],
  "objects": [
    {
      "type": "text",
      "name": "title",
      "role": "タイトル"
    },
    {
      "type": "text",
      "name": "content",
      "role": "本文"
    }
  ]
}
```

## トラブルシューティング

### よくある問題

1. **OpenAI API キーエラー**
   - `.env` ファイルにAPIキーが正しく設定されているか確認

2. **テンプレートファイルが見つからない**
   - `templates/` ディレクトリ内の構成を確認
   - `template.pptx` ファイルが存在するか確認

3. **PowerPoint生成エラー**
   - 中間ファイルを確認してどの段階でエラーが発生したか特定
   - `--debug` オプションで詳細なエラー情報を確認

## 技術仕様

- **フレームワーク**: LangGraph
- **AIモデル**: OpenAI GPT-4o
- **PowerPoint生成**: python-pptx
- **テンプレート管理**: JSON形式

## ライセンス

MIT License

## 貢献

プルリクエストやイシューの報告を歓迎します。

## 更新履歴

- v1.0.0: 初回リリース
  - 基本的なマークダウン→PowerPoint変換機能
  - 5つの基本テンプレート
  - LangGraphベースのワークフロー
