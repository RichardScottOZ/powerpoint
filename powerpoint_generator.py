"""
PowerPoint Generator Module

This module provides functions to create PowerPoint presentations programmatically.
Based on the Claude Skills pptx skill structure.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from typing import List, Dict, Any, Optional


class PowerPointGenerator:
    """A class to generate PowerPoint presentations with various slide layouts."""
    
    def __init__(self):
        """Initialize a new PowerPoint presentation."""
        self.prs = Presentation()
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(7.5)
    
    def add_title_slide(self, title: str, subtitle: str = "") -> None:
        """
        Add a title slide to the presentation.
        
        Args:
            title: Main title text
            subtitle: Subtitle text (optional)
        """
        slide_layout = self.prs.slide_layouts[0]  # Title slide layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        title_placeholder = slide.shapes.title
        title_placeholder.text = title
        
        if subtitle and len(slide.placeholders) > 1:
            subtitle_placeholder = slide.placeholders[1]
            subtitle_placeholder.text = subtitle
    
    def add_content_slide(self, title: str, content: List[str]) -> None:
        """
        Add a content slide with bullet points.
        
        Args:
            title: Slide title
            content: List of bullet points
        """
        slide_layout = self.prs.slide_layouts[1]  # Title and content layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        title_placeholder = slide.shapes.title
        title_placeholder.text = title
        
        body_placeholder = slide.placeholders[1]
        text_frame = body_placeholder.text_frame
        text_frame.clear()
        
        for i, item in enumerate(content):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            p.text = item
            p.level = 0
            p.font.size = Pt(18)
    
    def add_two_column_slide(self, title: str, left_content: List[str], 
                            right_content: List[str]) -> None:
        """
        Add a slide with two columns of content.
        
        Args:
            title: Slide title
            left_content: List of items for left column
            right_content: List of items for right column
        """
        slide_layout = self.prs.slide_layouts[5]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Add title
        left = Inches(0.5)
        top = Inches(0.5)
        width = Inches(9)
        height = Inches(1)
        
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = title
        p = title_frame.paragraphs[0]
        p.font.size = Pt(32)
        p.font.bold = True
        
        # Left column
        left_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), 
                                           Inches(4.5), Inches(5))
        left_frame = left_box.text_frame
        left_frame.word_wrap = True
        
        for i, item in enumerate(left_content):
            if i == 0:
                p = left_frame.paragraphs[0]
            else:
                p = left_frame.add_paragraph()
            p.text = item
            p.font.size = Pt(16)
        
        # Right column
        right_box = slide.shapes.add_textbox(Inches(5.5), Inches(2), 
                                            Inches(4), Inches(5))
        right_frame = right_box.text_frame
        right_frame.word_wrap = True
        
        for i, item in enumerate(right_content):
            if i == 0:
                p = right_frame.paragraphs[0]
            else:
                p = right_frame.add_paragraph()
            p.text = item
            p.font.size = Pt(16)
    
    def add_section_header(self, title: str, subtitle: str = "") -> None:
        """
        Add a section header slide.
        
        Args:
            title: Section title
            subtitle: Optional subtitle
        """
        slide_layout = self.prs.slide_layouts[2]  # Section header layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        title_placeholder = slide.shapes.title
        title_placeholder.text = title
        
        if subtitle and len(slide.placeholders) > 1:
            body_placeholder = slide.placeholders[1]
            body_placeholder.text = subtitle
    
    def add_blank_slide_with_text(self, content: str, 
                                  font_size: int = 24) -> None:
        """
        Add a blank slide with centered text.
        
        Args:
            content: Text content
            font_size: Font size in points
        """
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        left = Inches(1)
        top = Inches(2.5)
        width = Inches(8)
        height = Inches(2)
        
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.text = content
        
        p = text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(font_size)
        p.font.bold = True
    
    def save(self, filename: str) -> None:
        """
        Save the presentation to a file.
        
        Args:
            filename: Output filename (should end with .pptx)
        """
        self.prs.save(filename)
        print(f"Presentation saved as: {filename}")


def create_simple_presentation(title: str, slides_data: List[Dict[str, Any]], 
                               filename: str) -> None:
    """
    Create a simple presentation from structured data.
    
    Args:
        title: Presentation title
        slides_data: List of dictionaries with slide information
        filename: Output filename
    """
    ppt = PowerPointGenerator()
    ppt.add_title_slide(title)
    
    for slide_info in slides_data:
        slide_type = slide_info.get('type', 'content')
        
        if slide_type == 'content':
            ppt.add_content_slide(
                slide_info.get('title', ''),
                slide_info.get('content', [])
            )
        elif slide_type == 'section':
            ppt.add_section_header(
                slide_info.get('title', ''),
                slide_info.get('subtitle', '')
            )
        elif slide_type == 'two_column':
            ppt.add_two_column_slide(
                slide_info.get('title', ''),
                slide_info.get('left_content', []),
                slide_info.get('right_content', [])
            )
    
    ppt.save(filename)
