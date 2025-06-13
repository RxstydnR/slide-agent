import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import argparse

from src.workflow import SlideGenerationWorkflow
from src.models import SlideGenerationResponse

def load_environment():
    """環境変数を読み込み"""
    load_dotenv()
    
    openai_api_key = os.getenv("OPENAI_API_KEY")
    if not openai_api_key:
        print("Error: OPENAI_API_KEY environment variable is not set.")
        print("Please create a .env file with your OpenAI API key:")
        print("OPENAI_API_KEY=your_api_key_here")
        sys.exit(1)
    
    return openai_api_key

def read_markdown_file(file_path: str) -> str:
    """マークダウンファイルを読み込み"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        print(f"Error: File not found: {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(description="マークダウンからPowerPointスライドを自動生成")
    parser.add_argument("input_file", help="入力マークダウンファイルのパス")
    parser.add_argument("-o", "--output", help="出力ファイル名（省略時は自動生成）")
    parser.add_argument("--debug", action="store_true", help="デバッグモード")
    
    args = parser.parse_args()
    
    # 環境変数を読み込み
    openai_api_key = load_environment()
    
    # 入力ファイルを読み込み
    print(f"マークダウンファイルを読み込み中: {args.input_file}")
    markdown_content = read_markdown_file(args.input_file)
    
    # ワークフローを初期化
    print("スライド生成ワークフローを初期化中...")
    workflow = SlideGenerationWorkflow(openai_api_key)
    
    try:
        # ワークフローを実行
        print("スライド生成を開始します...")
        print("=" * 50)
        
        result = workflow.run(markdown_content)
        
        print("=" * 50)
        
        # 結果を表示
        if result.error_message:
            print(f"エラーが発生しました: {result.error_message}")
            return SlideGenerationResponse(
                success=False,
                error_message=result.error_message,
                intermediate_files=result.intermediate_files
            )
        
        if result.output_file:
            print(f"✅ PowerPointファイルが正常に生成されました!")
            print(f"📁 出力ファイル: {result.output_file}")
            
            if result.intermediate_files:
                print(f"\n📋 中間ファイル ({len(result.intermediate_files)}個):")
                for file in result.intermediate_files:
                    print(f"   - {file}")
            
            return SlideGenerationResponse(
                success=True,
                output_file=result.output_file,
                intermediate_files=result.intermediate_files
            )
        else:
            print("❌ PowerPointファイルの生成に失敗しました")
            return SlideGenerationResponse(
                success=False,
                error_message="出力ファイルが生成されませんでした",
                intermediate_files=result.intermediate_files
            )
            
    except Exception as e:
        error_message = f"予期しないエラーが発生しました: {str(e)}"
        print(f"❌ {error_message}")
        
        if args.debug:
            import traceback
            traceback.print_exc()
        
        return SlideGenerationResponse(
            success=False,
            error_message=error_message
        )

if __name__ == "__main__":
    main()
