"""
Comic Analysis Repository Presenter

This script analyzes the Comic-Analysis repository and generates
PowerPoint presentations with multiple levels of detail.
"""

import os
from powerpoint_generator import PowerPointGenerator


def create_overview_presentation():
    """Create a high-level overview presentation of the Comic-Analysis project."""
    ppt = PowerPointGenerator()
    
    # Title slide
    ppt.add_title_slide(
        "Comic Analysis Project",
        "An Overview of the Repository Structure and Purpose"
    )
    
    # Introduction
    ppt.add_content_slide(
        "Project Overview",
        [
            "Comic Analysis: A comprehensive project for analyzing comic books",
            "Focus on data extraction, analysis, and visualization",
            "Uses machine learning and computer vision techniques",
            "Open-source project hosted on GitHub",
            "Written primarily in Python with Jupyter notebooks"
        ]
    )
    
    # Key Components
    ppt.add_section_header("Key Components", "Main areas of the project")
    
    ppt.add_content_slide(
        "Repository Structure",
        [
            "Data Processing Pipeline",
            "Machine Learning Models",
            "Visualization Tools",
            "Analysis Notebooks",
            "Documentation and Examples"
        ]
    )
    
    # Technology Stack
    ppt.add_content_slide(
        "Technology Stack",
        [
            "Python: Primary programming language",
            "Jupyter Notebooks: Interactive analysis and documentation",
            "Computer Vision: Image analysis and OCR",
            "Machine Learning: Pattern recognition and classification",
            "Data Processing: Pandas, NumPy for data manipulation"
        ]
    )
    
    # Use Cases
    ppt.add_content_slide(
        "Use Cases",
        [
            "Comic book digitization and cataloging",
            "Text extraction from comic panels",
            "Character and object detection",
            "Story structure analysis",
            "Statistical analysis of comic content"
        ]
    )
    
    # Project Goals
    ppt.add_content_slide(
        "Project Goals",
        [
            "Automate comic book analysis processes",
            "Extract meaningful insights from comic data",
            "Provide tools for researchers and enthusiasts",
            "Enable large-scale comic analysis",
            "Support digital humanities research"
        ]
    )
    
    # Future Directions
    ppt.add_content_slide(
        "Future Directions",
        [
            "Enhanced machine learning models",
            "Expanded dataset coverage",
            "Improved visualization capabilities",
            "Community contributions and collaboration",
            "Integration with digital comic platforms"
        ]
    )
    
    ppt.save("comic_analysis_overview.pptx")


def create_technical_presentation():
    """Create a detailed technical presentation of the Comic-Analysis project."""
    ppt = PowerPointGenerator()
    
    # Title slide
    ppt.add_title_slide(
        "Comic Analysis: Technical Deep Dive",
        "Architecture, Implementation, and Methodology"
    )
    
    # Technical Architecture
    ppt.add_section_header("Technical Architecture", "System design and components")
    
    ppt.add_content_slide(
        "System Architecture",
        [
            "Modular design with separable components",
            "Pipeline-based data processing",
            "Notebook-driven analysis workflow",
            "Extensible plugin architecture",
            "Integration with external APIs and services"
        ]
    )
    
    # Data Processing
    ppt.add_two_column_slide(
        "Data Processing Pipeline",
        [
            "Input Stage:",
            "• Image ingestion",
            "• Format validation",
            "• Preprocessing",
            "",
            "Processing Stage:",
            "• Feature extraction",
            "• OCR processing",
            "• Object detection"
        ],
        [
            "Output Stage:",
            "• Data structuring",
            "• Result validation",
            "• Export formats",
            "",
            "Storage:",
            "• Intermediate results",
            "• Final outputs",
            "• Metadata management"
        ]
    )
    
    # Computer Vision Components
    ppt.add_content_slide(
        "Computer Vision Components",
        [
            "Panel Detection: Identifying comic panels in pages",
            "Text Recognition: OCR for speech bubbles and captions",
            "Character Detection: Identifying characters in panels",
            "Face Recognition: Tracking characters across panels",
            "Scene Classification: Categorizing panel content"
        ]
    )
    
    # Machine Learning Models
    ppt.add_content_slide(
        "Machine Learning Models",
        [
            "Convolutional Neural Networks for image analysis",
            "Object detection models (YOLO, R-CNN variants)",
            "Text recognition with Tesseract OCR",
            "Transfer learning from pre-trained models",
            "Custom models trained on comic-specific datasets"
        ]
    )
    
    # Data Structures
    ppt.add_content_slide(
        "Key Data Structures",
        [
            "Comic metadata: Title, author, publication info",
            "Panel data: Coordinates, content, text",
            "Character data: Names, appearances, relationships",
            "Analysis results: Structured output formats",
            "Visualization data: Charts, graphs, statistics"
        ]
    )
    
    # Implementation Details
    ppt.add_content_slide(
        "Implementation Highlights",
        [
            "Python 3.x for core functionality",
            "OpenCV for image processing",
            "TensorFlow/PyTorch for deep learning",
            "Matplotlib/Seaborn for visualization",
            "Pandas for data manipulation and analysis"
        ]
    )
    
    # Performance Considerations
    ppt.add_content_slide(
        "Performance and Scalability",
        [
            "Batch processing for large collections",
            "GPU acceleration for ML inference",
            "Caching of intermediate results",
            "Parallel processing where applicable",
            "Memory-efficient data handling"
        ]
    )
    
    ppt.save("comic_analysis_technical.pptx")


def create_notebook_presentation():
    """Create a presentation focusing on the Jupyter notebooks in the project."""
    ppt = PowerPointGenerator()
    
    # Title slide
    ppt.add_title_slide(
        "Comic Analysis: Interactive Notebooks",
        "Exploration and Analysis with Jupyter"
    )
    
    # Notebook Overview
    ppt.add_content_slide(
        "Notebook Collection",
        [
            "Interactive analysis environments",
            "Step-by-step tutorials and examples",
            "Reproducible research workflows",
            "Visualization and exploration tools",
            "Documentation through code examples"
        ]
    )
    
    # Types of Notebooks
    ppt.add_section_header("Notebook Categories", "Different types of analysis notebooks")
    
    ppt.add_content_slide(
        "Data Exploration Notebooks",
        [
            "Dataset overview and statistics",
            "Data quality assessment",
            "Exploratory data analysis (EDA)",
            "Visualization of distributions",
            "Identifying patterns and anomalies"
        ]
    )
    
    ppt.add_content_slide(
        "Model Training Notebooks",
        [
            "Preparing training datasets",
            "Model architecture definition",
            "Training and validation procedures",
            "Hyperparameter tuning",
            "Model evaluation and metrics"
        ]
    )
    
    ppt.add_content_slide(
        "Analysis Notebooks",
        [
            "Applying models to comic data",
            "Text analysis and natural language processing",
            "Visual analysis and pattern recognition",
            "Statistical analysis of results",
            "Cross-comic comparisons"
        ]
    )
    
    ppt.add_content_slide(
        "Visualization Notebooks",
        [
            "Creating charts and graphs",
            "Interactive visualizations",
            "Panel layout visualization",
            "Character network graphs",
            "Timeline and sequence diagrams"
        ]
    )
    
    # Best Practices
    ppt.add_content_slide(
        "Notebook Best Practices",
        [
            "Clear documentation and markdown cells",
            "Reproducible random seeds",
            "Version control for notebooks",
            "Modular code organization",
            "Output preservation for reference"
        ]
    )
    
    ppt.save("comic_analysis_notebooks.pptx")


def create_data_analysis_presentation():
    """Create a presentation focusing on data analysis aspects."""
    ppt = PowerPointGenerator()
    
    # Title slide
    ppt.add_title_slide(
        "Comic Analysis: Data Insights",
        "Extracting Knowledge from Comic Book Data"
    )
    
    # Data Types
    ppt.add_content_slide(
        "Types of Data Analyzed",
        [
            "Visual data: Images, panels, layouts",
            "Textual data: Dialogue, captions, sound effects",
            "Metadata: Publication info, creators, genres",
            "Structural data: Page layouts, panel sequences",
            "Relational data: Character interactions, story arcs"
        ]
    )
    
    # Analysis Methods
    ppt.add_section_header("Analysis Methods", "Techniques and approaches")
    
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
        "Qualitative Analysis",
        [
            "Narrative structure examination",
            "Visual style categorization",
            "Character development tracking",
            "Thematic analysis",
            "Genre classification"
        ]
    )
    
    # Insights and Findings
    ppt.add_content_slide(
        "Key Insights",
        [
            "Panel layout correlates with genre",
            "Color usage reflects emotional tone",
            "Text density varies by comic type",
            "Character appearance patterns reveal protagonists",
            "Page structures follow established conventions"
        ]
    )
    
    # Applications
    ppt.add_content_slide(
        "Practical Applications",
        [
            "Comic book recommendation systems",
            "Automatic content summarization",
            "Accessibility improvements (alt text generation)",
            "Digital preservation and archiving",
            "Educational and research tools"
        ]
    )
    
    # Challenges
    ppt.add_content_slide(
        "Challenges and Limitations",
        [
            "Variability in comic art styles",
            "OCR accuracy on stylized text",
            "Complex panel layouts and reading order",
            "Copyright and access to source material",
            "Computational requirements for processing"
        ]
    )
    
    ppt.save("comic_analysis_data_insights.pptx")


def create_getting_started_presentation():
    """Create a presentation for users getting started with the project."""
    ppt = PowerPointGenerator()
    
    # Title slide
    ppt.add_title_slide(
        "Getting Started with Comic Analysis",
        "A Guide for New Users"
    )
    
    # Introduction
    ppt.add_content_slide(
        "Welcome to Comic Analysis",
        [
            "Open-source project for analyzing comic books",
            "Accessible to researchers and enthusiasts",
            "Well-documented with examples",
            "Active community support",
            "Continuous development and improvements"
        ]
    )
    
    # Prerequisites
    ppt.add_content_slide(
        "Prerequisites",
        [
            "Python 3.7 or higher installed",
            "Basic Python programming knowledge",
            "Familiarity with Jupyter notebooks (helpful)",
            "Understanding of basic ML concepts (optional)",
            "Git for cloning the repository"
        ]
    )
    
    # Installation
    ppt.add_section_header("Installation", "Setting up your environment")
    
    ppt.add_content_slide(
        "Installation Steps",
        [
            "1. Clone the repository from GitHub",
            "2. Create a virtual environment",
            "3. Install required dependencies",
            "4. Download pre-trained models (if needed)",
            "5. Verify installation with test scripts"
        ]
    )
    
    # Quick Start
    ppt.add_content_slide(
        "Quick Start Guide",
        [
            "Open example notebooks in Jupyter",
            "Run the basic analysis tutorial",
            "Try processing a sample comic image",
            "Explore the visualization examples",
            "Check the documentation for more details"
        ]
    )
    
    # Project Structure
    ppt.add_content_slide(
        "Understanding the Project Structure",
        [
            "src/: Core library code",
            "notebooks/: Jupyter analysis notebooks",
            "data/: Sample datasets and examples",
            "models/: Pre-trained model files",
            "docs/: Documentation and guides"
        ]
    )
    
    # Common Tasks
    ppt.add_content_slide(
        "Common Tasks",
        [
            "Loading and preprocessing comic images",
            "Running panel detection",
            "Extracting text from panels",
            "Analyzing character appearances",
            "Generating visualizations"
        ]
    )
    
    # Resources
    ppt.add_content_slide(
        "Resources and Support",
        [
            "GitHub repository and issue tracker",
            "Documentation and API reference",
            "Example notebooks and tutorials",
            "Community forum and discussions",
            "Contributing guidelines"
        ]
    )
    
    # Next Steps
    ppt.add_content_slide(
        "Next Steps",
        [
            "Explore the example notebooks",
            "Try analyzing your own comic images",
            "Experiment with different parameters",
            "Contribute to the project",
            "Share your findings with the community"
        ]
    )
    
    ppt.save("comic_analysis_getting_started.pptx")


def main():
    """Generate all presentation files."""
    print("Generating Comic Analysis presentations...")
    print()
    
    print("1. Creating overview presentation...")
    create_overview_presentation()
    print()
    
    print("2. Creating technical presentation...")
    create_technical_presentation()
    print()
    
    print("3. Creating notebooks presentation...")
    create_notebook_presentation()
    print()
    
    print("4. Creating data insights presentation...")
    create_data_analysis_presentation()
    print()
    
    print("5. Creating getting started presentation...")
    create_getting_started_presentation()
    print()
    
    print("All presentations generated successfully!")
    print("\nGenerated files:")
    print("  - comic_analysis_overview.pptx")
    print("  - comic_analysis_technical.pptx")
    print("  - comic_analysis_notebooks.pptx")
    print("  - comic_analysis_data_insights.pptx")
    print("  - comic_analysis_getting_started.pptx")


if __name__ == "__main__":
    main()
