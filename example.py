"""
Simple example demonstrating the PowerPoint Generator.

This script creates a basic presentation to show how to use the library.
"""

from powerpoint_generator import PowerPointGenerator


def create_example_presentation():
    """Create a simple example presentation."""
    ppt = PowerPointGenerator()
    
    # Title slide
    ppt.add_title_slide(
        "PowerPoint Generator Demo",
        "A Simple Example Presentation"
    )
    
    # Introduction
    ppt.add_content_slide(
        "Welcome",
        [
            "This is an example presentation",
            "Created using the PowerPoint Generator library",
            "It demonstrates basic functionality",
            "You can create presentations programmatically"
        ]
    )
    
    # Section header
    ppt.add_section_header("Features", "What you can do")
    
    # Feature list
    ppt.add_content_slide(
        "Available Slide Types",
        [
            "Title slides with subtitles",
            "Content slides with bullet points",
            "Section headers",
            "Two-column layouts",
            "Blank slides with custom text"
        ]
    )
    
    # Two-column example
    ppt.add_two_column_slide(
        "Two-Column Layout Example",
        [
            "Left Column:",
            "• Easy to use",
            "• Flexible design",
            "• Python-based",
            "• Open source"
        ],
        [
            "Right Column:",
            "• Multiple layouts",
            "• Customizable",
            "• Well-documented",
            "• Active development"
        ]
    )
    
    # Benefits
    ppt.add_content_slide(
        "Benefits",
        [
            "Automate presentation creation",
            "Consistent formatting",
            "Easy to maintain and update",
            "Version control friendly",
            "Integrate with data pipelines"
        ]
    )
    
    # Final slide
    ppt.add_content_slide(
        "Get Started",
        [
            "Install: pip install -r requirements.txt",
            "Import: from powerpoint_generator import PowerPointGenerator",
            "Create: ppt = PowerPointGenerator()",
            "Add slides: ppt.add_title_slide(...)",
            "Save: ppt.save('presentation.pptx')"
        ]
    )
    
    ppt.save("example_presentation.pptx")
    print("Example presentation created: example_presentation.pptx")


if __name__ == "__main__":
    create_example_presentation()
