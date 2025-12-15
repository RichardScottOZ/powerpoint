"""
Example demonstrating the PowerPoint generator.

This shows how to create presentations using the working generator.
"""

from pptx_generator_working import PPTXGeneratorWorking as PPTXGenerator


def create_simple_example():
    """Create a simple example presentation."""
    ppt = PPTXGenerator()
    
    # Title slide
    ppt.add_title_slide(
        "From-Scratch PowerPoint Generator",
        "Built using XML and Python standard library"
    )
    
    # About slide
    ppt.add_content_slide(
        "About This Generator",
        [
            "Creates .pptx files by building XML structure directly",
            "No external libraries like python-pptx needed",
            "Uses Python's zipfile and xml.etree modules only",
            "Full control over PowerPoint Open XML format",
            "Educational and practical"
        ]
    )
    
    # How it works
    ppt.add_content_slide(
        "How It Works",
        [
            "PowerPoint files are ZIP archives containing XML",
            "Generates [Content_Types].xml for file definitions",
            "Creates presentation.xml for structure",
            "Builds individual slide XML files",
            "Includes themes, layouts, and relationships"
        ]
    )
    
    # XML structure
    ppt.add_content_slide(
        "PowerPoint XML Structure",
        [
            "[Content_Types].xml: MIME type definitions",
            "ppt/presentation.xml: Main presentation",
            "ppt/slides/slide*.xml: Individual slides",
            "ppt/slideLayouts/: Layout templates",
            "ppt/slideMasters/: Master templates",
            "ppt/theme/: Theme definitions"
        ]
    )
    
    # Features
    ppt.add_content_slide(
        "Current Features",
        [
            "Title slides with subtitles",
            "Content slides with bullet points",
            "Custom positioning and sizing",
            "Font size and color control",
            "Text alignment options",
            "Bold text support"
        ]
    )
    
    # Usage
    ppt.add_content_slide(
        "Usage Example",
        [
            "from pptx_generator import PPTXGenerator",
            "ppt = PPTXGenerator()",
            "ppt.add_title_slide('Title', 'Subtitle')",
            "ppt.add_content_slide('Topic', ['Point 1', 'Point 2'])",
            "ppt.save('output.pptx')"
        ]
    )
    
    ppt.save("example_presentation.pptx")


if __name__ == "__main__":
    print("Creating example presentation using from-scratch XML generator...")
    print()
    create_simple_example()
    print()
    print("Done! Open example_presentation.pptx to view the result.")
