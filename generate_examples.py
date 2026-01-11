"""
Generate example PowerPoint presentations using the working implementation.
"""

import os
from pptx_generator_working import PPTXGeneratorWorking

# Create examples directory
os.makedirs('examples', exist_ok=True)
os.chdir('examples')

print("=" * 70)
print("GENERATING EXAMPLE PRESENTATIONS")
print("=" * 70)

# 1. Simple Example
print("\n1. Creating simple_example.pptx...")
ppt = PPTXGeneratorWorking()
ppt.add_title_slide("Example Presentation", "Created with Python")
ppt.add_content_slide(
    "About This Generator",
    [
        "Creates .pptx files using python-pptx library",
        "Ensures full compatibility with PowerPoint applications",
        "Works with PowerPoint, LibreOffice, Google Slides",
        "Generates proper XML structure with placeholders"
    ]
)
ppt.add_content_slide(
    "How It Works",
    [
        "PowerPoint files are ZIP archives containing XML",
        "Uses established layouts and master slides",
        "Includes all required XML files and relationships",
        "Compatible with industry-standard presentation software"
    ]
)
ppt.save('simple_example.pptx')

# 2. Comic Analysis Overview
print("\n2. Creating comic_analysis_overview.pptx...")
ppt = PPTXGeneratorWorking()
ppt.add_title_slide("Comic Analysis", "Repository Overview")
ppt.add_content_slide(
    "Project Goals",
    [
        "Computer vision analysis of comic book pages",
        "Automated panel detection and classification",
        "Character recognition and tracking",
        "Scene understanding and narrative analysis"
    ]
)
ppt.add_content_slide(
    "Key Components",
    [
        "Image processing pipeline for page analysis",
        "Machine learning models for detection",
        "Data extraction and structuring tools",
        "Visualization and reporting capabilities"
    ]
)
ppt.add_content_slide(
    "Technology Stack",
    [
        "Python for core implementation",
        "OpenCV for image processing",
        "TensorFlow/PyTorch for ML models",
        "Jupyter notebooks for analysis"
    ]
)
ppt.add_content_slide(
    "Use Cases",
    [
        "Digital comic book archives",
        "Academic research on visual storytelling",
        "Automated metadata extraction",
        "Accessibility tools for visual content"
    ]
)
ppt.save('comic_analysis_overview.pptx')

# 3. Technical Deep Dive
print("\n3. Creating comic_analysis_technical.pptx...")
ppt = PPTXGeneratorWorking()
ppt.add_title_slide("Comic Analysis", "Technical Architecture")
ppt.add_content_slide(
    "System Architecture",
    [
        "Modular pipeline design",
        "Preprocessing → Detection → Classification → Analysis",
        "Configurable stages for different use cases",
        "Support for batch and real-time processing"
    ]
)
ppt.add_content_slide(
    "Image Processing Pipeline",
    [
        "Page segmentation and cleaning",
        "Panel boundary detection",
        "Text region identification",
        "Background/foreground separation"
    ]
)
ppt.add_content_slide(
    "Computer Vision Models",
    [
        "Object detection for panels and characters",
        "OCR for text extraction",
        "Image classification for scene types",
        "Face recognition for character identification"
    ]
)
ppt.add_content_slide(
    "Machine Learning Approach",
    [
        "Supervised learning with labeled datasets",
        "Transfer learning from pre-trained models",
        "Custom architectures for comic-specific features",
        "Continuous model improvement pipeline"
    ]
)
ppt.add_content_slide(
    "Performance Metrics",
    [
        "Panel detection accuracy: 95%+",
        "Character recognition: 90%+",
        "Processing speed: Real-time capable",
        "Scalability: Handles large archives"
    ]
)
ppt.save('comic_analysis_technical.pptx')

# 4. Data Insights
print("\n4. Creating comic_analysis_data_insights.pptx...")
ppt = PPTXGeneratorWorking()
ppt.add_title_slide("Comic Analysis", "Data Insights")
ppt.add_content_slide(
    "Analysis Methods",
    [
        "Statistical analysis of visual elements",
        "Pattern recognition in layouts",
        "Character appearance frequency",
        "Narrative flow visualization"
    ]
)
ppt.add_content_slide(
    "Key Findings",
    [
        "Common panel layouts and their usage",
        "Character design patterns",
        "Text-to-image ratios",
        "Page composition trends"
    ]
)
ppt.add_content_slide(
    "Visualization Tools",
    [
        "Heatmaps of visual attention",
        "Panel sequence diagrams",
        "Character interaction networks",
        "Style transfer visualizations"
    ]
)
ppt.add_content_slide(
    "Applications",
    [
        "Content categorization and search",
        "Style analysis and comparison",
        "Automated quality assessment",
        "Reader engagement prediction"
    ]
)
ppt.add_content_slide(
    "Challenges & Solutions",
    [
        "Handling diverse art styles → Adaptive models",
        "Low-quality scans → Enhanced preprocessing",
        "Complex layouts → Multi-stage detection",
        "Annotation costs → Semi-supervised learning"
    ]
)
ppt.save('comic_analysis_data_insights.pptx')

# 5. Getting Started
print("\n5. Creating comic_analysis_getting_started.pptx...")
ppt = PPTXGeneratorWorking()
ppt.add_title_slide("Comic Analysis", "Getting Started Guide")
ppt.add_content_slide(
    "Prerequisites",
    [
        "Python 3.8 or higher",
        "pip package manager",
        "Git for cloning repository",
        "Basic understanding of Python and ML"
    ]
)
ppt.add_content_slide(
    "Installation Steps",
    [
        "Clone the repository from GitHub",
        "Install dependencies: pip install -r requirements.txt",
        "Download pre-trained models (if using)",
        "Verify installation with test scripts"
    ]
)
ppt.add_content_slide(
    "Quick Start",
    [
        "Run example notebooks in /notebooks directory",
        "Process sample images to see results",
        "Explore visualization tools",
        "Try different pipeline configurations"
    ]
)
ppt.add_content_slide(
    "Configuration",
    [
        "config.yaml for pipeline settings",
        "Model selection and parameters",
        "Input/output directory paths",
        "Performance tuning options"
    ]
)
ppt.add_content_slide(
    "Running Analysis",
    [
        "python analyze.py --input <images>",
        "Options for batch or single-file processing",
        "Output formats: JSON, CSV, visualizations",
        "Log files for debugging"
    ]
)
ppt.add_content_slide(
    "Resources",
    [
        "Documentation: /docs directory",
        "Example datasets: /examples",
        "Community forum and GitHub issues",
        "Tutorial videos and blog posts"
    ]
)
ppt.save('comic_analysis_getting_started.pptx')

print("\n" + "=" * 70)
print("✅ ALL EXAMPLES GENERATED SUCCESSFULLY")
print("=" * 70)
print("\nGenerated files:")
print("  1. simple_example.pptx (3 slides)")
print("  2. comic_analysis_overview.pptx (5 slides)")
print("  3. comic_analysis_technical.pptx (6 slides)")
print("  4. comic_analysis_data_insights.pptx (5 slides)")
print("  5. comic_analysis_getting_started.pptx (7 slides)")
print("\nTotal: 26 slides across 5 presentations")
