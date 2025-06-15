import json
import os
import logging
import time
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


format_content_system_prompt = """
あなたはプレゼンテーション資料作成の専門家です。
与えられたスライドコンテンツを、プレゼンテーションに適した形式に整形してください。

整形の方針：
1. 内容を簡潔で分かりやすくする
2. 箇条書きや構造化された形式にする
3. 重要なポイントを明確にする
4. 元の内容の意図を保持する

元のコンテンツをそのまま返すのではなく、プレゼンテーション用に最適化してください。
"""
            
format_content_human_prompt = """
以下のスライドコンテンツを整形してください：
{content}
"""

select_template_system_prompt = """
あなたはプレゼンテーション資料作成の専門家です。
与えられたスライドコンテンツに最適なテンプレートを選択してください。

# 利用可能なテンプレート：
{templates_info}

# 選択の方針：
1. スライドの内容と目的に最も適したテンプレートを選ぶ
2. テンプレートの利用ケース例を参考にする
3. スライドの位置（最初、中間、最後）も考慮する

回答は選択したテンプレート名のみを返してください。
"""
            
select_template_human_prompt = """
以下のスライドコンテンツに最適なテンプレートを選択してください：

スライド番号: {slide_number}
コンテンツ:
{slide_content}
"""

assign_content_system_prompt = """
あなたはプレゼンテーション資料作成の専門家です。
与えられたスライドコンテンツを、指定されたテンプレートの各オブジェクトに適切に割り当ててください。

テンプレート: {template_name}
オブジェクト構成: {objects_info}
"""
# 回答はJSON形式で、各オブジェクト名をキーとし、割り当てるテキストを値として返してください。
# 例: {{"title": "スライドタイトル", "main_text": "本文内容"}}
            
assign_content_human_prompt = """
以下のスライドコンテンツを、テンプレートの各オブジェクトに割り当ててください：
{slide_content}
"""

import logging

class SlideGenerationWorkflow:
    """スライド生成ワークフロー"""

    def __init__(self):
        self.llm = ChatOpenAI(model="gpt-4.1", temperature=0.1)
        self.template_manager = TemplateManager()
        self.markdown_parser = MarkdownParser()
        self.powerpoint_generator = PowerPointGenerator(self.template_manager)
        
        # ログの出力形式を定義
        FORMATTER = '%(asctime)s - (%(filename)s) - [%(levelname)s] - %(message)s'

        # ロガーを作成
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

        # DEBUGレベルのログを'log/debug_handler.log'へ出力するハンドラを作成
        os.makedirs('log', exist_ok=True)
        log_filename = f'log/debug_handler_{int(time.time())}.log'
        file_handler = logging.FileHandler(log_filename, mode='w')
        file_handler.setLevel(logging.DEBUG)
        self.logger.addHandler(file_handler)

        # StreamHandlerも追加（formatterは無視）
        stream_handler = logging.StreamHandler()
        stream_handler.setLevel(logging.DEBUG)
        self.logger.addHandler(stream_handler)
    
        
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
        self.logger.info("Step 1: マークダウンを解析中...")
        self.logger.debug(f"入力マークダウン: \n<INPUT>\n{state.markdown_content}\n</INPUT>\n")
        
        # マークダウンを解析
        slides = self.markdown_parser.parse_markdown(state.markdown_content)
        self.logger.debug(f"解析結果スライド数: {len(slides)}")
        
        # 中間ファイルとして保存
        intermediate_file = self.intermediate_dir / f"01_parsed_slides_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(intermediate_file, 'w', encoding='utf-8') as f:
            slides_data = [slide.model_dump() for slide in slides]
            json.dump(slides_data, f, ensure_ascii=False, indent=2)
        self.logger.debug(f"中間ファイル保存: {intermediate_file}")
        
        return {
            "slides": slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _format_content_node(self, state: WorkflowState) -> Dict[str, Any]:
        """コンテンツ整形ノード"""
        self.logger.info("Step 2: コンテンツを整形中...")
        
        formatted_slides = []
        
        for slide in state.slides:
            self.logger.debug(f"スライド番号: {slide.slide_number} の整形前コンテンツ: {slide.content}")
            # AIにコンテンツの整形を依頼
            messages = [
                SystemMessage(content=format_content_system_prompt),
                HumanMessage(content=format_content_human_prompt.format(content=slide.content))
            ]
            
            response = self.llm.invoke(messages)
            formatted_content = response.content
            self.logger.debug(f"スライド番号: {slide.slide_number} の整形後コンテンツ: {formatted_content}")
            
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
        self.logger.debug(f"中間ファイル保存: {intermediate_file}")
        
        return {
            "slides": formatted_slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _select_templates_node(self, state: WorkflowState) -> Dict[str, Any]:
        """テンプレート選択ノード"""
        self.logger.info("Step 3: 各スライドに最適なテンプレートを選択中...")
        
        # テンプレート情報を取得
        templates_info = self.template_manager.get_templates_info_for_ai()
        self.logger.debug(f"テンプレート情報: {templates_info}")
        
        processed_slides = []
        
        for slide in state.slides:
            self.logger.debug(f"スライド番号: {slide.slide_number} のテンプレート選択開始")
            # AIにテンプレート選択を依頼   
            messages = [
                SystemMessage(content=select_template_system_prompt.format(templates_info=templates_info)),
                HumanMessage(content=select_template_human_prompt.format(slide_number=slide.slide_number, slide_content=slide.content))
            ]
            
            # pydanticモデルを定義し、llm.with_structured_outputを使って構造的に出力を取得
            from pydantic import BaseModel,Field
            class TemplateSelectionOutput(BaseModel):
                template_name: str
                reason: str = Field(..., description="reasonは1行で完結に")

            response = self.llm.with_structured_output(TemplateSelectionOutput).invoke(messages)
            selected_template = response.template_name.strip()
            self.logger.debug(f"AI選択テンプレート: {selected_template} (理由: {response.reason})")
            
            # 選択されたテンプレートが存在するかチェック
            if selected_template not in self.template_manager.get_template_names():
                self.logger.warning(f"選択されたテンプレート名「{selected_template}」が見つかりません。補正を試みます。")
                # テンプレート名一覧を取得
                template_names = self.template_manager.get_template_names()
                # 補正用プロンプト
                correction_system_prompt = (
                    "あなたはプレゼン資料作成AIです。"
                    "以下のテンプレート名一覧から、最も近いものを1つだけ選び、テンプレート名のみを返してください。"
                    "必ず一覧に含まれる名前を返してください。"
                    "\nテンプレート名一覧:\n" + "\n".join(template_names)
                )
                correction_human_prompt = (
                    f"AIが選択したテンプレート名: {response.template_name}\n"
                    "正しいテンプレート名を1つだけ返してください。\n"
                    "正しいテンプレート名:"
                )
                correction_messages = [
                    SystemMessage(content=correction_system_prompt),
                    HumanMessage(content=correction_human_prompt)
                ]
                # 補正結果を取得
                corrected_template = self.llm.invoke(correction_messages).content.strip()
                self.logger.debug(f"補正後テンプレート名: {corrected_template}")
                if corrected_template in template_names:
                    self.logger.info("選択されたテンプレート名への補正が完了しました。")
                    selected_template = corrected_template
                else:
                    self.logger.error("再度選択されたテンプレート名も不正です。デフォルトのテンプレート「1カラムテキスト」を選択します。")
                    # それでも見つからない場合はデフォルト
                    selected_template = "1カラムテキスト"
            
            processed_slide = ProcessedSlide(
                slide_number=slide.slide_number,
                template_name=selected_template,
                content_mapping={}  # 次のステップで設定
            )
            processed_slides.append(processed_slide)
            self.logger.debug(f"スライド番号: {slide.slide_number} のテンプレート選択完了: {selected_template}")
        
        # 中間ファイルとして保存
        intermediate_file = self.intermediate_dir / f"03_template_selection_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(intermediate_file, 'w', encoding='utf-8') as f:
            slides_data = [slide.model_dump() for slide in processed_slides]
            json.dump(slides_data, f, ensure_ascii=False, indent=2)
        self.logger.debug(f"中間ファイル保存: {intermediate_file}")
        
        return {
            "processed_slides": processed_slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _assign_content_node(self, state: WorkflowState) -> Dict[str, Any]:
        """コンテンツ割り当てノード"""
        self.logger.info("Step 4: テンプレートにコンテンツを割り当て中...")
        
        updated_processed_slides = []
        
        for i, processed_slide in enumerate(state.processed_slides):
            slide_content = state.slides[i]
            template = self.template_manager.get_template(processed_slide.template_name)
            
            if not template:
                self.logger.error(f"テンプレート「{processed_slide.template_name}」が見つかりません。スキップします。")
                continue
            
            # AIにコンテンツ割り当てを依頼
            objects_info = []
            for obj in template.objects:
                objects_info.append(f"- {obj.name}: {obj.role}")
            
            messages = [
                SystemMessage(
                    content=assign_content_system_prompt.format(
                        template_name=template.template_name, 
                        objects_info=chr(10).join(objects_info)
                )),
                HumanMessage(
                    content=assign_content_human_prompt.format(
                        slide_content=slide_content.content
                    )
                )
            ]
            
            # Pydanticでnameとcontentのペアをリストとして構造化出力させる
            from pydantic import BaseModel, Field
            from typing import List

            class ObjectContentPair(BaseModel):
                name: str = Field(description="オブジェクト名")
                content: str = Field(description="割り当てる内容")

            class ContentMappingList(BaseModel):
                items: List[ObjectContentPair]

            # LLMに構造化出力を要求
            response = self.llm.with_structured_output(ContentMappingList).invoke(messages)

            # nameとcontentのペアのリストをdictに変換（nameをキー、contentを値）
            content_mapping = {item.name: item.content for item in response.items}
            self.logger.debug(f"スライド番号: {slide_content.slide_number} のcontent_mapping: {content_mapping}")
            
            # もし失敗した場合デフォルトマッピング
            # content_mapping = {
            #     "title": f"スライド {slide_content.slide_number}", 
            #     "main_text": slide_content.content
            # }
        
            # ProcessedSlideを更新
            updated_slide = ProcessedSlide(
                slide_number=processed_slide.slide_number,
                template_name=processed_slide.template_name,
                content_mapping=content_mapping
            )
            updated_processed_slides.append(updated_slide)
            self.logger.debug(f"スライド番号: {slide_content.slide_number} の割り当て完了")
        
        # 中間ファイルとして保存
        intermediate_file = self.intermediate_dir / f"04_content_assignment_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(intermediate_file, 'w', encoding='utf-8') as f:
            slides_data = [slide.model_dump() for slide in updated_processed_slides]
            json.dump(slides_data, f, ensure_ascii=False, indent=2)
        self.logger.debug(f"中間ファイル保存: {intermediate_file}")
        
        return {
            "processed_slides": updated_processed_slides,
            "intermediate_files": state.intermediate_files + [str(intermediate_file)]
        }
    
    def _generate_powerpoint_node(self, state: WorkflowState) -> Dict[str, Any]:
        """PowerPoint生成ノード"""
        self.logger.info("Step 5: PowerPointファイルを生成中...")
        
        try:
            # PowerPointファイルを生成
            output_file = self.powerpoint_generator.generate_presentation(
                state.processed_slides,
                f"generated_slides_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            )
            
            self.logger.info(f"PowerPointファイルが生成されました: {output_file}")
            
            return {
                "output_file": output_file,
                "intermediate_files": state.intermediate_files
            }
            
        except Exception as e:
            error_message = f"PowerPoint生成エラー: {str(e)}"
            self.logger.error(error_message)
            
            return {
                "error_message": error_message,
                "intermediate_files": state.intermediate_files
            }
    
    def run(self, markdown_content: str) -> WorkflowState:
        """ワークフローを実行"""
        
        self.logger.info("ワークフローの実行を開始します。")
        initial_state = WorkflowState(
            markdown_content=markdown_content,
            templates=self.template_manager.get_all_templates()
        )
        
        # ワークフローを実行
        self.logger.debug("StateGraphのinvokeを呼び出します。")
        final_state = self.workflow.invoke(initial_state)
        self.logger.info("ワークフローの実行が完了しました。")
        
        return WorkflowState(**final_state)
