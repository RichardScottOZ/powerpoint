# PowerPoint Generator for Comic-Analysis Repository

A **from-scratch PowerPoint generator** that creates multi-level presentations about the [Comic-Analysis repository](https://github.com/RichardScottOZ/Comic-Analysis). Built entirely using Python's standard library (`xml.etree.ElementTree` and `zipfile`) - **no external dependencies** for generation.

## 🎯 Built From Scratch

This implementation creates PowerPoint files by manually generating ALL required XML files including:
- Complete slide layouts with placeholders (11 layouts)
- Slide masters with proper structure
- Theme definitions
- Document properties
- Relationship files
- All required namespaces and attributes

**Zero external dependencies for generation** - uses only Python standard library!

## 📥 Example Presentations - WITH SCREENSHOTS

**Visual proof that presentations work!**

See [`SCREENSHOTS.md`](SCREENSHOTS.md) for visual previews of all presentations.

### Example: Simple Presentation
![Example Slide](https://github.com/user-attachments/assets/dad3164c-b844-4016-a551-d6e2abcbe5ac)

**Download presentations (generated from scratch):**
- [`simple_example.pptx`](examples/simple_example.pptx) - Basic 3-slide demonstration (17.5 KB, 40 files)
- [`comic_analysis_overview.pptx`](examples/comic_analysis_overview.pptx) - 5-slide overview (19.3 KB, 44 files)
- [`comic_analysis_technical.pptx`](examples/comic_analysis_technical.pptx) - 7-slide technical deep dive (21.3 KB, 48 files)
- [`comic_analysis_data_insights.pptx`](examples/comic_analysis_data_insights.pptx) - 7-slide data analysis (21.3 KB, 48 files)
- [`comic_analysis_getting_started.pptx`](examples/comic_analysis_getting_started.pptx) - 8-slide getting started guide (22.2 KB, 50 files)

✅ All files validated - open successfully in Microsoft PowerPoint, LibreOffice Impress, Google Slides, and Apple Keynote.

## Overview

This project demonstrates how PowerPoint files work at the XML level and provides a functional generator that creates valid .pptx files by:

1. **Building XML structure manually** using Python's `xml.etree.ElementTree`
2. **Creating ZIP archives** containing the XML files using `zipfile`
3. **Following PowerPoint Open XML specification** for file format compliance

## How PowerPoint Files Work

A `.pptx` file is actually a **ZIP archive** containing multiple XML files:

```
presentation.pptx (ZIP archive)
├── [Content_Types].xml          # Defines MIME types for all parts
├── _rels/
│   └── .rels                    # Package-level relationships
├── ppt/
│   ├── presentation.xml         # Main presentation structure
│   ├── _rels/
│   │   └── presentation.xml.rels # Presentation relationships
│   ├── slides/
│   │   ├── slide1.xml           # Individual slide content (XML)
│   │   ├── slide2.xml
│   │   └── _rels/
│   │       ├── slide1.xml.rels  # Slide relationships
│   │       └── slide2.xml.rels
│   ├── slideLayouts/
│   │   ├── slideLayout1.xml     # Layout templates
│   │   └── _rels/
│   │       └── slideLayout1.xml.rels
│   ├── slideMasters/
│   │   ├── slideMaster1.xml     # Master templates
│   │   └── _rels/
│   │       └── slideMaster1.xml.rels
│   └── theme/
│       └── theme1.xml            # Theme definitions (colors, fonts)
```

## XML Elements Created

### Content Types XML
Defines the MIME type for each file in the package:
```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" 
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ...
</Types>
```

### Presentation XML
Main presentation structure with slide references:
```xml
<p:presentation xmlns:a="..." xmlns:p="..." xmlns:r="...">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>
```

### Slide XML
Individual slide content with shapes and text:
```xml
<p:sld xmlns:a="..." xmlns:p="..." xmlns:r="...">
  <p:cSld>
    <p:spTree>
      <!-- Shapes, text boxes, etc. -->
      <p:sp>
        <p:nvSpPr>...</p:nvSpPr>
        <p:spPr>...</p:spPr>
        <p:txBody>
          <a:p>
            <a:r>
              <a:t>Hello World</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
```

## Installation

**No installation required!** Uses only Python standard library.

```bash
git clone https://github.com/RichardScottOZ/powerpoint.git
cd powerpoint
python example.py
```

## Usage

### Basic Example

```python
from pptx_generator_complete import CompletePPTXGenerator

# Create generator
ppt = CompletePPTXGenerator()

# Add title slide
ppt.add_title_slide(
    "My Presentation",
    "Created from Scratch"
)

# Add content slide with bullet points
ppt.add_content_slide(
    "Key Points",
    [
        "No external dependencies",
        "Pure XML generation",
        "Complete slide layouts",
        "All placeholders included"
    ]
)

# Add two-column slide
ppt.add_two_column_slide(
    "Comparison",
    ["From scratch", "Complete XML", "Standard library only"],
    ["Full compatibility", "Proper layouts", "Works everywhere"]
)

# Save (creates ZIP with 40+ XML files)
ppt.save("my_presentation.pptx", verbose=True)
```

### Advanced Usage with Custom Shapes

```python
from pptx_generator import PPTXGenerator

ppt = PPTXGenerator()

# Add custom positioned text box
ppt.add_slide([
    {
        'x': 914400,        # 1 inch in EMUs
        'y': 914400,        # 1 inch in EMUs
        'width': 4572000,   # 5 inches
        'height': 914400,   # 1 inch
        'text': ['Custom text box'],
        'font_size': 2400,  # 24pt
        'bold': True,
        'align': 'ctr',     # center
        'color': '0000FF',  # blue
    }
])

ppt.save("custom.pptx")
```

### EMU Units

PowerPoint uses **English Metric Units (EMUs)** for measurements:
- 1 inch = 914,400 EMUs
- 1 cm = 360,000 EMUs
- 1 point = 12,700 EMUs (for font sizes, multiply pt by 100)

## Comic Analysis Presentations

Generate comprehensive presentations about the Comic-Analysis repository:

```bash
python comic_analysis_presenter.py
```

This creates 4 presentations:
- `comic_analysis_overview.pptx` - High-level overview (5 slides)
- `comic_analysis_technical.pptx` - Technical deep dive (6 slides)
- `comic_analysis_data_insights.pptx` - Data analysis insights (5 slides)
- `comic_analysis_getting_started.pptx` - Getting started guide (6 slides)

## Testing

Run the comprehensive test suite to validate the Comic-Analysis presentations:

```bash
python run_tests.py
```

Or run tests directly:

```bash
python test_comic_analysis.py
```

The test suite includes:
- **23 tests** covering all Comic-Analysis presentations
- Validates PPTX file structure (ZIP archives with XML)
- Checks slide counts and content
- Verifies XML validity and namespace correctness
- Tests the core PPTXGenerator functionality

Test coverage:
- `TestComicAnalysisPresentations`: 20 tests for the 4 Comic-Analysis presentations
- `TestPPTXGenerator`: 3 tests for the core generator

## Project Structure

```
powerpoint/
├── pptx_generator_complete.py  # Complete from-scratch generator (900+ lines)
├── pptx_generator.py            # Original simplified version (reference)
├── comic_analysis_presenter.py  # Comic-Analysis presentations
├── example.py                   # Simple example
├── test_comic_analysis.py       # Comprehensive test suite (23 tests)
├── run_tests.py                 # Test runner script
├── examples/                    # Generated presentation files
│   ├── simple_example.pptx
│   ├── comic_analysis_overview.pptx
│   ├── comic_analysis_technical.pptx
│   ├── comic_analysis_data_insights.pptx
│   └── comic_analysis_getting_started.pptx
├── screenshots/                 # Preview screenshots
├── XML_STRUCTURE.md             # Detailed XML documentation
├── EXAMPLES.md                  # Example files documentation
└── README.md                    # This file
```

## Features

- ✅ **Zero dependencies** - uses only Python standard library
- ✅ **Direct XML generation** - full control over output
- ✅ **PowerPoint Open XML compliant** - works in MS Office, LibreOffice, etc.
- ✅ **Title slides** with subtitles
- ✅ **Content slides** with bullet points
- ✅ **Custom shapes** with positioning and styling
- ✅ **Font control** - size, color, bold, alignment
- ✅ **Complete documentation** of XML structure

## Technical Details

### XML Namespaces

```python
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}
```

### File Generation Process

1. **Create XML elements** using `xml.etree.ElementTree`
2. **Build document structure** (presentation, slides, themes)
3. **Define relationships** between parts
4. **Package into ZIP** using `zipfile` module
5. **Save as .pptx** file

## Complete Implementation

This from-scratch implementation includes:
- ✅ **11 slide layouts** (title, content, two-column, section header, blank, etc.)
- ✅ **Complete slide master** with proper structure
- ✅ **Placeholders** for title, content, and body text
- ✅ **Theme definitions** with colors and fonts
- ✅ **Document properties** (core and application properties)
- ✅ **View and presentation properties**
- ✅ **Table styles** (required by PowerPoint)
- ✅ **Default text styles** for proper rendering
- ✅ **All relationship files** properly configured

Current limitations (extendable):
- No images, charts, or tables yet
- No animations or transitions
- Basic theme (customizable by modifying theme XML)

The foundation is complete - these features can be added by extending the XML generation.

## Resources

- [PowerPoint Open XML Format](http://officeopenxml.com/anatomyofOOXML-pptx.php)
- [ECMA-376 Standard](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [Office Open XML Wikipedia](https://en.wikipedia.org/wiki/Office_Open_XML)

## License

Open source and available for use and modification.
