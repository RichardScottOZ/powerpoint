# PowerPoint Generator

A Python-based PowerPoint presentation generator built from scratch, designed to create comprehensive presentations about code repositories and projects.

## Features

- **Programmatic PowerPoint Generation**: Create presentations using Python code
- **Multiple Slide Types**: Title slides, content slides, two-column layouts, section headers
- **Repository Analysis**: Automatically generate presentations about GitHub repositories
- **Multi-Level Presentations**: Create overview, technical, and detailed presentations

## Installation

1. Clone this repository:
```bash
git clone https://github.com/RichardScottOZ/powerpoint.git
cd powerpoint
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Generate Comic Analysis Presentations

To generate all presentations about the Comic-Analysis repository:

```bash
python comic_analysis_presenter.py
```

This will create five PowerPoint presentations:
- `comic_analysis_overview.pptx` - High-level project overview
- `comic_analysis_technical.pptx` - Technical deep dive
- `comic_analysis_notebooks.pptx` - Jupyter notebook analysis
- `comic_analysis_data_insights.pptx` - Data analysis insights
- `comic_analysis_getting_started.pptx` - Getting started guide

### Using the PowerPoint Generator Library

You can also use the `PowerPointGenerator` class to create custom presentations:

```python
from powerpoint_generator import PowerPointGenerator

# Create a new presentation
ppt = PowerPointGenerator()

# Add a title slide
ppt.add_title_slide("My Presentation", "Subtitle Here")

# Add a content slide with bullet points
ppt.add_content_slide(
    "Key Points",
    [
        "First point",
        "Second point",
        "Third point"
    ]
)

# Add a section header
ppt.add_section_header("New Section", "Section description")

# Save the presentation
ppt.save("my_presentation.pptx")
```

## Project Structure

```
powerpoint/
├── powerpoint_generator.py      # Core PowerPoint generation library
├── comic_analysis_presenter.py  # Comic-Analysis repository presenter
├── requirements.txt             # Python dependencies
└── README.md                    # This file
```

## Requirements

- Python 3.7+
- python-pptx 0.6.23
- requests 2.31.0

## About

This project is based on the Claude Skills pptx skill and demonstrates how to:
- Create PowerPoint presentations programmatically
- Generate multi-level documentation about code repositories
- Provide overview, technical, and getting-started content
- Structure complex information into digestible slides

## Example: Comic-Analysis Repository

The included `comic_analysis_presenter.py` creates comprehensive presentations about the [Comic-Analysis repository](https://github.com/RichardScottOZ/Comic-Analysis), covering:

- Project overview and goals
- Technical architecture and implementation
- Jupyter notebook structure and usage
- Data analysis methods and insights
- Getting started guide for new users

## License

This project is open source and available for use and modification.
