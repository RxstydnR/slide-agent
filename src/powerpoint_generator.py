import os
import shutil
from typing import List, Dict, Optional
from pathlib import Path
from pptx import Presentation
from src.models import ProcessedSlide, Template
from src.template_manager import TemplateManager

class PowerPointGenerator:
    """PowerPoint生成クラス"""
    
    def __init__(self, template_manager: TemplateManager):
        self.template_manager = template_manager
        self.output_dir = Path("output")
        self.output_dir.mkdir(exist_ok=True)
    
    def generate_presentation(self, processed_slides: List[ProcessedSlide], output_filename: str = "generated_slides.pptx") -> str:
        """処理済みスライドからPowerPointプレゼンテーションを生成"""
        
        # 新しいプレゼンテーションを作成
        prs = Presentation()
        
        # 最初のデフォルトスライドを削除
        if len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[0]
        
        for processed_slide in processed_slides:
            self._add_slide_from_template(prs, processed_slide)
        
        # 出力ファイルパス
        output_path = self.output_dir / output_filename
        
        # プレゼンテーションを保存
        prs.save(str(output_path))
        
        return str(output_path)
    
    def _add_slide_from_template(self, prs: Presentation, processed_slide: ProcessedSlide) -> None:
        """テンプレートからスライドを追加"""
        
        # テンプレートのPowerPointファイルパスを取得
        template_pptx_path = self.template_manager.get_template_pptx_path(processed_slide.template_name)
        
        if not template_pptx_path or not template_pptx_path.exists():
            print(f"Warning: Template file not found for {processed_slide.template_name}")
            return
        
        # テンプレートプレゼンテーションを読み込み
        template_prs = Presentation(str(template_pptx_path))
        
        if len(template_prs.slides) == 0:
            print(f"Warning: No slides found in template {processed_slide.template_name}")
            return
        
        # テンプレートの最初のスライドを取得
        template_slide = template_prs.slides[0]
        
        # テンプレートスライドをコピーしてメインプレゼンテーションに追加
        self._copy_slide_with_content(prs, template_slide, processed_slide.content_mapping)
    
    def _copy_slide_with_content(self, prs: Presentation, template_slide, content_mapping: Dict[str, str]) -> None:
        """テンプレートスライドをコピーしてコンテンツを埋め込み"""
        
        # 新しいスライドを追加（空のレイアウトを使用）
        blank_layout = prs.slide_layouts[6]  # Blank layout
        new_slide = prs.slides.add_slide(blank_layout)
        
        # テンプレートスライドの全てのshapeをコピー
        for shape in template_slide.shapes:
            self._copy_shape_with_content(new_slide, shape, content_mapping)
    
    def _copy_shape_with_content(self, slide, template_shape, content_mapping: Dict[str, str]) -> None:
        """shapeをコピーしてコンテンツを埋め込み"""
        
        # shapeの基本情報を取得
        left = template_shape.left
        top = template_shape.top
        width = template_shape.width
        height = template_shape.height
        
        # テキストshapeの場合
        if hasattr(template_shape, 'text'):
            # 元のテキストを取得
            original_text = template_shape.text.strip()
            
            # content_mappingから対応するテキストを取得
            replacement_text = content_mapping.get(original_text, original_text)
            
            # 新しいテキストボックスを作成
            new_textbox = slide.shapes.add_textbox(left, top, width, height)
            new_textbox.text = replacement_text
        
        # その他のshape（画像など）の場合
        else:
            # プレースホルダーとして空のテキストボックスを作成
            placeholder_textbox = slide.shapes.add_textbox(left, top, width, height)
            placeholder_textbox.text = "[要素]"
    
    def generate_single_slide(self, processed_slide: ProcessedSlide, output_filename: str = None) -> str:
        """単一スライドのPowerPointファイルを生成（デバッグ用）"""
        
        if output_filename is None:
            output_filename = f"slide_{processed_slide.slide_number}_{processed_slide.template_name}.pptx"
        
        # テンプレートのPowerPointファイルパスを取得
        template_pptx_path = self.template_manager.get_template_pptx_path(processed_slide.template_name)
        
        if not template_pptx_path or not template_pptx_path.exists():
            raise FileNotFoundError(f"Template file not found: {processed_slide.template_name}")
        
        # テンプレートプレゼンテーションを読み込み
        template_prs = Presentation(str(template_pptx_path))
        
        if len(template_prs.slides) == 0:
            raise ValueError(f"No slides found in template: {processed_slide.template_name}")
        
        # テンプレートの最初のスライドを取得
        template_slide = template_prs.slides[0]
        
        # 各shapeのテキストを置換
        for shape in template_slide.shapes:
            if hasattr(shape, 'text'):
                original_text = shape.text.strip()
                if original_text in processed_slide.content_mapping:
                    shape.text = processed_slide.content_mapping[original_text]
        
        # 出力ファイルパス
        output_path = self.output_dir / output_filename
        
        # プレゼンテーションを保存
        template_prs.save(str(output_path))
        
        return str(output_path)
