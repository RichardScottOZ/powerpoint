"""
Comic Analysis Repository Presenter

Creates PowerPoint presentations about the Comic-Analysis repository
using the working PowerPoint generator.
"""

from pptx_generator_complete import CompletePPTXGenerator as PPTXGenerator


def create_overview_presentation():
    """Create high-level overview presentation."""
    ppt = PPTXGenerator()
    
    ppt.add_title_slide(
        "Comic Analysis Project",
        "Overview of Repository Structure and Purpose"
    )
    
    ppt.add_content_slide(
        "Project Overview",
        [
            "Comic Analysis: Comprehensive comic book analysis project",
            "Data extraction, analysis, and visualization focus",
            "Machine learning and computer vision techniques",
            "Open-source GitHub project",
            "Python and Jupyter notebooks"
        ]
    )
    
    ppt.add_content_slide(
        "Key Components",
        [
            "Data Processing Pipeline",
            "Machine Learning Models",
            "Visualization Tools",
            "Analysis Notebooks",
            "Documentation and Examples"
        ]
    )
    
    ppt.add_content_slide(
        "Technology Stack",
        [
            "Python: Primary language",
            "Jupyter Notebooks: Interactive analysis",
            "Computer Vision: Image analysis and OCR",
            "Machine Learning: Pattern recognition",
            "Data Processing: Pandas, NumPy"
        ]
    )
    
    ppt.add_content_slide(
        "Use Cases",
        [
            "Comic book digitization and cataloging",
            "Text extraction from comic panels",
            "Character and object detection",
            "Story structure analysis",
            "Statistical analysis of content"
        ]
    )
    
    ppt.save("comic_analysis_overview.pptx")


def create_technical_presentation():
    """Create detailed technical presentation."""
    ppt = PPTXGenerator()
    
    ppt.add_title_slide(
        "Comic Analysis: Technical Deep Dive",
        "Architecture, Implementation, and Methodology"
    )
    
    ppt.add_content_slide(
        "System Architecture",
        [
            "Modular design with separable components",
            "Pipeline-based data processing",
            "Notebook-driven analysis workflow",
            "Extensible plugin architecture",
            "Integration with external APIs"
        ]
    )
    
    ppt.add_content_slide(
        "Data Processing Pipeline",
        [
            "Input: Image ingestion and validation",
            "Processing: Feature extraction and OCR",
            "Analysis: Object detection and classification",
            "Output: Data structuring and export",
            "Storage: Results and metadata management"
        ]
    )
    
    ppt.add_content_slide(
        "Computer Vision Components",
        [
            "Panel Detection: Identifying panels in pages",
            "Text Recognition: OCR for speech bubbles",
            "Character Detection: Identifying characters",
            "Face Recognition: Tracking across panels",
            "Scene Classification: Categorizing content"
        ]
    )
    
    ppt.add_content_slide(
        "Machine Learning Models",
        [
            "CNNs for image analysis",
            "Object detection (YOLO, R-CNN)",
            "Text recognition with Tesseract OCR",
            "Transfer learning from pre-trained models",
            "Custom models on comic datasets"
        ]
    )
    
    ppt.add_content_slide(
        "Performance Considerations",
        [
            "Batch processing for large collections",
            "GPU acceleration for ML inference",
            "Caching of intermediate results",
            "Parallel processing where applicable",
            "Memory-efficient data handling"
        ]
    )
    
    ppt.save("comic_analysis_technical.pptx")


def create_data_insights_presentation():
    """Create data analysis insights presentation."""
    ppt = PPTXGenerator()
    
    ppt.add_title_slide(
        "Comic Analysis: Data Insights",
        "Extracting Knowledge from Comic Book Data"
    )
    
    ppt.add_content_slide(
        "Types of Data Analyzed",
        [
            "Visual: Images, panels, layouts",
            "Textual: Dialogue, captions, effects",
            "Metadata: Publication info, creators, genres",
            "Structural: Page layouts, panel sequences",
            "Relational: Character interactions, story arcs"
        ]
    )
    
    ppt.add_content_slide(
        "Quantitative Analysis",
        [
            "Panel count and size statistics",
            "Color palette analysis",
            "Text density measurements",
            "Character appearance frequency",
            "Page layout patterns"
        ]
    )
    
    ppt.add_content_slide(
        "Key Insights",
        [
            "Panel layout correlates with genre",
            "Color usage reflects emotional tone",
            "Text density varies by comic type",
            "Character patterns reveal protagonists",
            "Page structures follow conventions"
        ]
    )
    
    ppt.add_content_slide(
        "Applications",
        [
            "Comic book recommendation systems",
            "Automatic content summarization",
            "Accessibility improvements",
            "Digital preservation and archiving",
            "Educational and research tools"
        ]
    )
    
    ppt.save("comic_analysis_data_insights.pptx")


def create_getting_started_presentation():
    """Create getting started guide."""
    ppt = PPTXGenerator()
    
    ppt.add_title_slide(
        "Getting Started with Comic Analysis",
        "Guide for New Users"
    )
    
    ppt.add_content_slide(
        "Welcome",
        [
            "Open-source comic book analysis project",
            "Accessible to researchers and enthusiasts",
            "Well-documented with examples",
            "Active community support",
            "Continuous development"
        ]
    )
    
    ppt.add_content_slide(
        "Prerequisites",
        [
            "Python 3.7 or higher installed",
            "Basic Python programming knowledge",
            "Jupyter notebooks familiarity (helpful)",
            "Basic ML concepts (optional)",
            "Git for cloning repository"
        ]
    )
    
    ppt.add_content_slide(
        "Installation Steps",
        [
            "1. Clone repository from GitHub",
            "2. Create virtual environment",
            "3. Install required dependencies",
            "4. Download pre-trained models",
            "5. Verify with test scripts"
        ]
    )
    
    ppt.add_content_slide(
        "Quick Start",
        [
            "Open example notebooks in Jupyter",
            "Run basic analysis tutorial",
            "Process sample comic image",
            "Explore visualization examples",
            "Check documentation for details"
        ]
    )
    
    ppt.add_content_slide(
        "Next Steps",
        [
            "Explore example notebooks",
            "Analyze your own comic images",
            "Experiment with parameters",
            "Contribute to the project",
            "Share findings with community"
        ]
    )
    
    ppt.save("comic_analysis_getting_started.pptx")


def main():
    """Generate all presentations."""
    print("=" * 70)
    print("Comic Analysis PowerPoint Generator")
    print("Building presentations from scratch using XML")
    print("=" * 70)
    print()
    
    print("1. Creating overview presentation...")
    create_overview_presentation()
    print()
    
    print("2. Creating technical presentation...")
    create_technical_presentation()
    print()
    
    print("3. Creating data insights presentation...")
    create_data_insights_presentation()
    print()
    
    print("4. Creating getting started presentation...")
    create_getting_started_presentation()
    print()
    
    print("=" * 70)
    print("All presentations generated successfully!")
    print("=" * 70)


if __name__ == "__main__":
    main()
