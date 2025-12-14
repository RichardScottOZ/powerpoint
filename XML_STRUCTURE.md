# PowerPoint XML Structure Documentation

This document explains the XML structure that makes up a .pptx file and how this implementation creates it.

## What is a .pptx File?

A `.pptx` file is a **ZIP archive** containing XML files that define the presentation. This can be verified:

```bash
unzip -l presentation.pptx
```

## File Structure

```
presentation.pptx (ZIP archive)
│
├── [Content_Types].xml          # Defines MIME types for all parts
│
├── _rels/
│   └── .rels                    # Package-level relationships
│
└── ppt/
    ├── presentation.xml         # Main presentation definition
    │
    ├── _rels/
    │   └── presentation.xml.rels # Relationships to slides, masters
    │
    ├── slides/
    │   ├── slide1.xml           # Slide 1 content
    │   ├── slide2.xml           # Slide 2 content
    │   └── _rels/
    │       ├── slide1.xml.rels  # Slide 1 relationships
    │       └── slide2.xml.rels  # Slide 2 relationships
    │
    ├── slideLayouts/
    │   ├── slideLayout1.xml     # Layout template
    │   └── _rels/
    │       └── slideLayout1.xml.rels
    │
    ├── slideMasters/
    │   ├── slideMaster1.xml     # Master slide template
    │   └── _rels/
    │       └── slideMaster1.xml.rels
    │
    └── theme/
        └── theme1.xml           # Color, font, style definitions
```

## XML Namespaces

PowerPoint Open XML uses multiple XML namespaces:

```python
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',      # Drawing ML
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', # Presentation ML
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', # Relationships
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships', # Package relationships
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types', # Content types
}
```

## Key XML Files

### 1. [Content_Types].xml

Defines MIME types for each part in the package:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- Default types for extensions -->
  <Default Extension="rels" 
           ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" 
           ContentType="application/xml"/>
  
  <!-- Specific part types -->
  <Override PartName="/ppt/presentation.xml" 
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" 
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <!-- ... more overrides ... -->
</Types>
```

**Purpose**: Tells Microsoft Office what type of content each file contains.

### 2. _rels/.rels

Package-level relationships:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" 
                Target="ppt/presentation.xml"/>
</Relationships>
```

**Purpose**: Points to the main presentation.xml file.

### 3. ppt/presentation.xml

Main presentation structure:

```xml
<?xml version="1.0" encoding="utf-8"?>
<p:presentation xmlns:a="..." xmlns:p="..." xmlns:r="...">
  <!-- Slide Master references -->
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  
  <!-- Slide ID list -->
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>
  
  <!-- Slide dimensions (in EMUs) -->
  <p:sldSz cx="9144000" cy="6858000"/>
  
  <!-- Notes dimensions -->
  <p:notesSz cx="9144000" cy="6858000"/>
</p:presentation>
```

**Purpose**: 
- References all slides
- Defines slide size
- Links to slide master

**EMUs (English Metric Units)**: 
- 1 inch = 914,400 EMUs
- 1 cm = 360,000 EMUs

### 4. ppt/slides/slide1.xml

Individual slide content:

```xml
<?xml version="1.0" encoding="utf-8"?>
<p:sld xmlns:a="..." xmlns:p="..." xmlns:r="...">
  <p:cSld>
    <p:spTree>
      <!-- Group shape properties (required) -->
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="9144000" cy="6858000"/>
        </a:xfrm>
      </p:grpSpPr>
      
      <!-- Shapes (text boxes, images, etc.) -->
      <p:sp>
        <!-- Shape properties -->
        <p:nvSpPr>
          <p:cNvPr id="2" name="TextBox 2"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        
        <!-- Shape position and size -->
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="457200"/>    <!-- Position -->
            <a:ext cx="8229600" cy="914400"/> <!-- Size -->
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
        
        <!-- Text content -->
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0">
            <a:spAutoFit/>
          </a:bodyPr>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="l"/>
            <a:r>
              <a:rPr lang="en-US" sz="3200" b="1">
                <a:solidFill>
                  <a:srgbClr val="000000"/>
                </a:solidFill>
              </a:rPr>
              <a:t>Hello World</a:t>
            </a:r>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  
  <!-- Color map override -->
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>
```

**Purpose**:
- Contains all shapes on the slide
- Defines text content
- Sets positioning, sizing, colors, fonts

**Key Elements**:
- `<p:sp>` - Shape (text box, etc.)
- `<p:nvSpPr>` - Non-visual shape properties (ID, name)
- `<p:spPr>` - Visual properties (position, size, geometry)
- `<p:txBody>` - Text content
- `<a:p>` - Paragraph
- `<a:r>` - Text run
- `<a:t>` - Text content
- `<a:rPr>` - Run properties (font size, color, bold)

### 5. ppt/_rels/presentation.xml.rels

Relationships from presentation to other parts:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" 
                Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" 
                Target="slides/slide1.xml"/>
  <Relationship Id="rId3" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" 
                Target="slides/slide2.xml"/>
</Relationships>
```

**Purpose**: Links presentation.xml to slides and master templates.

### 6. ppt/theme/theme1.xml

Theme definition (colors, fonts, effects):

```xml
<?xml version="1.0" encoding="utf-8"?>
<a:theme xmlns:a="..." name="Office Theme">
  <a:themeElements>
    <!-- Color scheme -->
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>  <!-- Dark 1 -->
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>  <!-- Light 1 -->
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>  <!-- Dark 2 -->
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>  <!-- Light 2 -->
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <!-- ... more colors ... -->
    </a:clrScheme>
    
    <!-- Font scheme -->
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
      </a:minorFont>
    </a:fontScheme>
    
    <!-- Format scheme (fills, lines, effects) -->
    <a:fmtScheme name="Office">
      <!-- ... format definitions ... -->
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
```

**Purpose**: Defines default colors, fonts, and styles for the presentation.

## How This Implementation Works

### 1. Create XML Elements

```python
import xml.etree.ElementTree as ET

# Create root element with namespace
root = ET.Element(f"{{{NS['p']}}}presentation")

# Add child elements
sld_id_lst = ET.SubElement(root, f"{{{NS['p']}}}sldIdLst")
ET.SubElement(sld_id_lst, f"{{{NS['p']}}}sldId",
             attrib={'id': '256', 'r:id': 'rId2'})

# Convert to XML string
xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=True)
```

### 2. Package into ZIP

```python
import zipfile

with zipfile.ZipFile('presentation.pptx', 'w', zipfile.ZIP_DEFLATED) as zf:
    # Add all XML files
    zf.writestr('[Content_Types].xml', content_types_xml)
    zf.writestr('_rels/.rels', rels_xml)
    zf.writestr('ppt/presentation.xml', presentation_xml)
    zf.writestr('ppt/slides/slide1.xml', slide1_xml)
    # ... more files ...
```

### 3. Complete Implementation

See `pptx_generator.py` for the full implementation:

- `_create_content_types_xml()` - [Content_Types].xml
- `_create_rels_xml()` - _rels/.rels
- `_create_presentation_xml()` - ppt/presentation.xml
- `_create_slide_xml()` - ppt/slides/slide*.xml
- `_create_shape_xml()` - Individual shapes
- `_create_theme_xml()` - ppt/theme/theme1.xml
- And more...

## Text Properties

### Font Size
Font sizes are in points × 100:
- 18pt = 1800
- 24pt = 2400
- 32pt = 3200
- 44pt = 4400

### Alignment
- `l` - Left
- `ctr` - Center
- `r` - Right
- `just` - Justified

### Colors
RGB hex values:
- `000000` - Black
- `FFFFFF` - White
- `FF0000` - Red
- `0000FF` - Blue

## Resources

- **ECMA-376**: Office Open XML specification
  - http://www.ecma-international.org/publications-and-standards/standards/ecma-376/
  
- **Office Open XML Anatomy**:
  - http://officeopenxml.com/anatomyofOOXML-pptx.php
  
- **ISO/IEC 29500**: International standard for Office Open XML

## Validation

To validate a generated .pptx file:

```python
import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile('presentation.pptx', 'r') as zf:
    for filename in zf.namelist():
        if filename.endswith('.xml'):
            content = zf.read(filename)
            ET.fromstring(content)  # Will raise exception if invalid XML
```

## Extending the Implementation

To add new features:

1. **Images**: Add to `ppt/media/`, reference in `<p:pic>` elements
2. **Charts**: Create `ppt/charts/chart*.xml`
3. **Tables**: Use `<a:tbl>` elements in slides
4. **Animations**: Add `ppt/slides/slide*.xml` with `<p:timing>` elements
5. **Transitions**: Add `<p:transition>` to slide XML

Each requires:
- XML structure definition
- Content type entry
- Relationship definitions
- Proper namespace usage
