from typing import List
from src.models import SlideContent

class MarkdownParser:
    """マークダウン解析クラス"""
    
    def __init__(self):
        self.slide_separator = "---"
    
    def parse_markdown(self, markdown_content: str) -> List[SlideContent]:
        """マークダウンをスライドコンテンツに解析"""
        # マークダウンを正規化（改行コードを統一）
        normalized_content = markdown_content.replace('\r\n', '\n').replace('\r', '\n')
        
        # スライド区切りで分割
        slide_texts = normalized_content.split(self.slide_separator)
        
        slides = []
        for i, slide_text in enumerate(slide_texts):
            slide_text = slide_text.strip()
            if not slide_text:
                continue
            
            slide = SlideContent(
                content=slide_text,
                slide_number=i + 1
            )
            slides.append(slide)
        
        return slides
    
    def validate_markdown_format(self, markdown_content: str) -> bool:
        """マークダウン形式の妥当性をチェック"""
        # 基本的な妥当性チェック
        if not markdown_content.strip():
            return False
        
        # スライド区切りが存在するかチェック
        slides = markdown_content.split(self.slide_separator)
        if len(slides) < 2:
            # 区切りがない場合は単一スライドとして扱う
            return True
        
        # 各スライドに内容があるかチェック
        valid_slides = 0
        for slide in slides:
            if slide.strip():
                valid_slides += 1
        
        return valid_slides > 0
