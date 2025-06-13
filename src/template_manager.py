import json
import os
from typing import List, Dict, Optional
from pathlib import Path
from src.models import Template, TemplateObject

class TemplateManager:
    """テンプレート管理クラス"""
    
    def __init__(self, templates_dir: str = "templates"):
        self.templates_dir = Path(templates_dir)
        self._templates_cache: Dict[str, Template] = {}
        self._load_templates()
    
    def _load_templates(self) -> None:
        """全テンプレートを読み込み"""
        if not self.templates_dir.exists():
            raise FileNotFoundError(f"Templates directory not found: {self.templates_dir}")
        
        for template_dir in self.templates_dir.iterdir():
            if template_dir.is_dir():
                template_json_path = template_dir / "template.json"
                if template_json_path.exists():
                    try:
                        with open(template_json_path, 'r', encoding='utf-8') as f:
                            template_data = json.load(f)
                        
                        template = Template(**template_data)
                        self._templates_cache[template.template_name] = template
                    except Exception as e:
                        print(f"Error loading template {template_dir.name}: {e}")
    
    def get_all_templates(self) -> List[Template]:
        """全テンプレートを取得"""
        return list(self._templates_cache.values())
    
    def get_template(self, template_name: str) -> Optional[Template]:
        """指定されたテンプレートを取得"""
        return self._templates_cache.get(template_name)
    
    def get_template_names(self) -> List[str]:
        """全テンプレート名を取得"""
        return list(self._templates_cache.keys())
    
    def get_templates_info_for_ai(self) -> str:
        """AI用のテンプレート情報を文字列で取得"""
        info_parts = []
        for template in self._templates_cache.values():
            objects_info = []
            for obj in template.objects:
                objects_info.append(f"  - {obj.name} ({obj.type}): {obj.role}")
            
            template_info = f"""
テンプレート名: {template.template_name}
説明: {template.description}
利用ケース例: {', '.join(template.use_case_examples)}
オブジェクト構成:
{chr(10).join(objects_info)}
"""
            info_parts.append(template_info)
        
        return "\n" + "="*50 + "\n".join(info_parts)
    
    def get_template_pptx_path(self, template_name: str) -> Optional[Path]:
        """テンプレートのPowerPointファイルパスを取得"""
        if template_name in self._templates_cache:
            template_dir = self.templates_dir / template_name
            pptx_path = template_dir / "template.pptx"
            if pptx_path.exists():
                return pptx_path
        return None
    
    def reload_templates(self) -> None:
        """テンプレートを再読み込み"""
        self._templates_cache.clear()
        self._load_templates()
