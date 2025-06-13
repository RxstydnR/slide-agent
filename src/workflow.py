import json
import os
from typing import Dict, Any, List
from datetime import datetime
from pathlib import Path

from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from langgraph.graph import StateGraph, END
from langgraph.graph.message import add_messages

from src.models import WorkflowState, SlideContent, ProcessedSlide
from src.template_manager import TemplateManager
from src.markdown_parser import MarkdownParser
from src.powerpoint_generator import PowerPointGenerator

class SlideGenerationWorkflow:
    """スライド生成ワークフロー"""
    
    def __init__(self, openai_api_key: str):
        self.llm = ChatOpenAI(
            model="gpt-4o",
            api_key=openai_api_key,
            temperature=0.1
        )
        self.template_manager = TemplateManager()
        self.markdown_parser = MarkdownParser()
        self.powerpoint_generator = PowerPointGenerator(self.template_manager)
        
        # 中間ファイル保存用ディレクトリ
        self.intermediate_dir = Path("intermediate")
        self.intermediate_dir.mkdir(exist_ok=True)
        
        # ワークフローグラフを構築
        self.workflow = self._build_workflow()
    
    def _build_workflow(self) -> StateGraph:
        """ワークフローグラフを構築"""
        
        workflow = StateGraph(WorkflowState)
        
        # ノードを追加
        workflow.add_node("parse_markdown", self._parse_markdown_node)
        workflow.add_node("format_content", self._format_content_node)
        workflow.add_node("select_templates", self._select_templates_node)
        workflow.add_node("assign_content", self._assign_content_node)
        workflow.add_node("generate_powerpoint", self._generate_powerpoint_node)
        
        # エッジを追加
        workflow.set_entry_point("parse_markdown")
        workflow.add_edge("parse_markdown", "format_content")
        workflow.add_edge("format_content", "select_templates")
        workflow.add_edge("select_templates", "assign_content")
        workflow.add_edge("assign_content", "generate_powerpoint")
        workflow.add_edge("generate_powerpoint", END)
        
        return workflow.compile()
    
    def _parse_markdown_node(self, state: WorkflowState) -> Dict[str, Any]:
        """マークダウン解析ノード"""
        print("Step 1: マークダウンを解析中...")
        
        # マークダウンを解析
        slides = self.markdown_parser.parse_markdown(state.markdown_content)
        
        # 中間ファイルとして保存
        intermediate_file = self.intermediate_dir / f"01_parsed_slides_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(intermediate_file, 'w', encoding='utf-8') as f:
            slides_data = [slide.model_dump() for slide in slides]
            json.dump(slides_data, f, ensure_ascii=False, indent=2)
        
        return {
            "slides": slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _format_content_node(self, state: WorkflowState) -> Dict[str, Any]:
        """コンテンツ整形ノード"""
        print("Step 2: コンテンツを整形中...")
        
        formatted_slides = []
        
        for slide in state.slides:
            # AIにコンテンツの整形を依頼
            system_prompt = """
あなたはプレゼンテーション資料作成の専門家です。
与えられたスライドコンテンツを、プレゼンテーションに適した形式に整形してください。

整形の方針：
1. 内容を簡潔で分かりやすくする
2. 箇条書きや構造化された形式にする
3. 重要なポイントを明確にする
4. 元の内容の意図を保持する

元のコンテンツをそのまま返すのではなく、プレゼンテーション用に最適化してください。
"""
            
            human_prompt = f"""
以下のスライドコンテンツを整形してください：

{slide.content}
"""
            
            messages = [
                SystemMessage(content=system_prompt),
                HumanMessage(content=human_prompt)
            ]
            
            response = self.llm.invoke(messages)
            formatted_content = response.content
            
            # 整形されたスライドを作成
            formatted_slide = SlideContent(
                content=formatted_content,
                slide_number=slide.slide_number
            )
            formatted_slides.append(formatted_slide)
        
        # 中間ファイルとして保存
        intermediate_file = self.intermediate_dir / f"02_formatted_slides_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(intermediate_file, 'w', encoding='utf-8') as f:
            slides_data = [slide.model_dump() for slide in formatted_slides]
            json.dump(slides_data, f, ensure_ascii=False, indent=2)
        
        return {
            "slides": formatted_slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _select_templates_node(self, state: WorkflowState) -> Dict[str, Any]:
        """テンプレート選択ノード"""
        print("Step 3: 各スライドに最適なテンプレートを選択中...")
        
        # テンプレート情報を取得
        templates_info = self.template_manager.get_templates_info_for_ai()
        
        processed_slides = []
        
        for slide in state.slides:
            # AIにテンプレート選択を依頼
            system_prompt = f"""
あなたはプレゼンテーション資料作成の専門家です。
与えられたスライドコンテンツに最適なテンプレートを選択してください。

利用可能なテンプレート：
{templates_info}

選択の方針：
1. スライドの内容と目的に最も適したテンプレートを選ぶ
2. テンプレートの利用ケース例を参考にする
3. スライドの位置（最初、中間、最後）も考慮する

回答は選択したテンプレート名のみを返してください。
"""
            
            human_prompt = f"""
以下のスライドコンテンツに最適なテンプレートを選択してください：

スライド番号: {slide.slide_number}
コンテンツ:
{slide.content}
"""
            
            messages = [
                SystemMessage(content=system_prompt),
                HumanMessage(content=human_prompt)
            ]
            
            response = self.llm.invoke(messages)
            selected_template = response.content.strip()
            
            # 選択されたテンプレートが存在するかチェック
            if selected_template not in self.template_manager.get_template_names():
                # デフォルトテンプレートを使用
                selected_template = "1カラムテキスト"
            
            processed_slide = ProcessedSlide(
                slide_number=slide.slide_number,
                template_name=selected_template,
                content_mapping={}  # 次のステップで設定
            )
            processed_slides.append(processed_slide)
        
        # 中間ファイルとして保存
        intermediate_file = self.intermediate_dir / f"03_template_selection_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(intermediate_file, 'w', encoding='utf-8') as f:
            slides_data = [slide.model_dump() for slide in processed_slides]
            json.dump(slides_data, f, ensure_ascii=False, indent=2)
        
        return {
            "processed_slides": processed_slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _assign_content_node(self, state: WorkflowState) -> Dict[str, Any]:
        """コンテンツ割り当てノード"""
        print("Step 4: テンプレートにコンテンツを割り当て中...")
        
        updated_processed_slides = []
        
        for i, processed_slide in enumerate(state.processed_slides):
            slide_content = state.slides[i]
            template = self.template_manager.get_template(processed_slide.template_name)
            
            if not template:
                continue
            
            # AIにコンテンツ割り当てを依頼
            objects_info = []
            for obj in template.objects:
                objects_info.append(f"- {obj.name}: {obj.role}")
            
            system_prompt = f"""
あなたはプレゼンテーション資料作成の専門家です。
与えられたスライドコンテンツを、指定されたテンプレートの各オブジェクトに適切に割り当ててください。

テンプレート: {template.template_name}
オブジェクト構成:
{chr(10).join(objects_info)}

回答はJSON形式で、各オブジェクト名をキーとし、割り当てるテキストを値として返してください。
例: {{"title": "スライドタイトル", "main_text": "本文内容"}}
"""
            
            human_prompt = f"""
以下のスライドコンテンツを、テンプレートの各オブジェクトに割り当ててください：

{slide_content.content}
"""
            
            messages = [
                SystemMessage(content=system_prompt),
                HumanMessage(content=human_prompt)
            ]
            
            response = self.llm.invoke(messages)
            
            try:
                # JSONレスポンスを解析
                content_mapping = json.loads(response.content)
            except json.JSONDecodeError:
                # JSON解析に失敗した場合はデフォルトマッピング
                content_mapping = {"title": f"スライド {slide_content.slide_number}", "main_text": slide_content.content}
            
            # ProcessedSlideを更新
            updated_slide = ProcessedSlide(
                slide_number=processed_slide.slide_number,
                template_name=processed_slide.template_name,
                content_mapping=content_mapping
            )
            updated_processed_slides.append(updated_slide)
        
        # 中間ファイルとして保存
        intermediate_file = self.intermediate_dir / f"04_content_assignment_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(intermediate_file, 'w', encoding='utf-8') as f:
            slides_data = [slide.model_dump() for slide in updated_processed_slides]
            json.dump(slides_data, f, ensure_ascii=False, indent=2)
        
        return {
            "processed_slides": updated_processed_slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _generate_powerpoint_node(self, state: WorkflowState) -> Dict[str, Any]:
        """PowerPoint生成ノード"""
        print("Step 5: PowerPointファイルを生成中...")
        
        try:
            # PowerPointファイルを生成
            output_file = self.powerpoint_generator.generate_presentation(
                state.processed_slides,
                f"generated_slides_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            )
            
            print(f"PowerPointファイルが生成されました: {output_file}")
            
            return {
                "output_file": output_file,
                "intermediate_files": state.intermediate_files
            }
            
        except Exception as e:
            error_message = f"PowerPoint生成エラー: {str(e)}"
            print(error_message)
            
            return {
                "error_message": error_message,
                "intermediate_files": state.intermediate_files
            }
    
    def run(self, markdown_content: str) -> WorkflowState:
        """ワークフローを実行"""
        
        initial_state = WorkflowState(
            markdown_content=markdown_content,
            templates=self.template_manager.get_all_templates()
        )
        
        # ワークフローを実行
        final_state = self.workflow.invoke(initial_state)
        
        return WorkflowState(**final_state)
