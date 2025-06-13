from pydantic import BaseModel
from typing import List, Dict, Any, Optional
from enum import Enum

class ObjectType(str, Enum):
    TEXT = "text"
    IMAGE = "image"

class TemplateObject(BaseModel):
    type: ObjectType
    name: str
    role: str

class Template(BaseModel):
    template_name: str
    description: str
    use_case_examples: List[str]
    objects: List[TemplateObject]

class SlideContent(BaseModel):
    title: Optional[str] = None
    content: str
    slide_number: int

class ProcessedSlide(BaseModel):
    slide_number: int
    template_name: str
    content_mapping: Dict[str, str]

class SlideGenerationRequest(BaseModel):
    markdown_content: str

class SlideGenerationResponse(BaseModel):
    success: bool
    output_file: Optional[str] = None
    error_message: Optional[str] = None
    intermediate_files: List[str] = []

class WorkflowState(BaseModel):
    """LangGraphワークフローの状態管理用"""
    markdown_content: str = ""
    slides: List[SlideContent] = []
    processed_slides: List[ProcessedSlide] = []
    templates: List[Template] = []
    output_file: Optional[str] = None
    intermediate_files: List[str] = []
    error_message: Optional[str] = None
