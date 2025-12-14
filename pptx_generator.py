"""
PowerPoint Generator from Scratch

This module creates PowerPoint presentations by directly generating the XML
structure that comprises a .pptx file. A .pptx file is essentially a ZIP archive
containing XML files that define the presentation structure, slides, and content.

No external libraries are used - this is built entirely from scratch using only
Python's built-in zipfile and xml modules.
"""

import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from datetime import datetime
from typing import List, Dict, Any, Optional


class PPTXGenerator:
    """
    A from-scratch PowerPoint generator that creates .pptx files by building
    the underlying XML structure directly.
    
    PowerPoint Open XML Structure:
    - [Content_Types].xml: Defines content types for all parts
    - _rels/.rels: Package relationships
    - ppt/presentation.xml: Main presentation definition
    - ppt/_rels/presentation.xml.rels: Presentation relationships
    - ppt/slides/slide*.xml: Individual slide content
    - ppt/slides/_rels/slide*.xml.rels: Slide relationships
    - ppt/slideLayouts/slideLayout*.xml: Layout templates
    - ppt/slideMasters/slideMaster*.xml: Master templates
    - ppt/theme/theme*.xml: Theme definitions
    """
    
    # XML Namespaces used in PowerPoint Open XML
    NS = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    }
    
    def __init__(self, width_inches: float = 10, height_inches: float = 7.5):
        """
        Initialize a new PowerPoint presentation.
        
        Args:
            width_inches: Slide width in inches (default: 10)
            height_inches: Slide height in inches (default: 7.5)
        """
        # Convert inches to EMUs (English Metric Units)
        # 1 inch = 914400 EMUs
        self.width = int(width_inches * 914400)
        self.height = int(height_inches * 914400)
        
        self.slides = []
        self.slide_counter = 0
        
        # Register all namespaces for pretty XML output
        for prefix, uri in self.NS.items():
            ET.register_namespace(prefix, uri)
    
    def _create_content_types_xml(self) -> bytes:
        """
        Create [Content_Types].xml - defines MIME types for all parts.
        This file tells Office what type of content each file in the archive contains.
        """
        root = ET.Element(f"{{{self.NS['ct']}}}Types")
        
        # Default extensions
        defaults = [
            ('rels', 'application/vnd.openxmlformats-package.relationships+xml'),
            ('xml', 'application/xml'),
        ]
        for ext, content_type in defaults:
            ET.SubElement(root, f"{{{self.NS['ct']}}}Default",
                         Extension=ext, ContentType=content_type)
        
        # Override specific files
        overrides = [
            ('/ppt/presentation.xml', 
             'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'),
            ('/ppt/slideMasters/slideMaster1.xml',
             'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'),
            ('/ppt/slideLayouts/slideLayout1.xml',
             'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'),
            ('/ppt/theme/theme1.xml',
             'application/vnd.openxmlformats-officedocument.theme+xml'),
            ('/ppt/viewProps.xml',
             'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml'),
            ('/ppt/presProps.xml',
             'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'),
            ('/ppt/tableStyles.xml',
             'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'),
            ('/docProps/core.xml',
             'application/vnd.openxmlformats-package.core-properties+xml'),
            ('/docProps/app.xml',
             'application/vnd.openxmlformats-officedocument.extended-properties+xml'),
        ]
        
        # Add overrides for each slide
        for i in range(1, len(self.slides) + 1):
            overrides.append(
                (f'/ppt/slides/slide{i}.xml',
                 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
            )
        
        for part_name, content_type in overrides:
            ET.SubElement(root, f"{{{self.NS['ct']}}}Override",
                         PartName=part_name, ContentType=content_type)
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_rels_xml(self) -> bytes:
        """
        Create _rels/.rels - defines package-level relationships.
        Points to the main presentation.xml file and document properties.
        """
        root = ET.Element(f"{{{self.NS['rel']}}}Relationships")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId1",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                     Target="ppt/presentation.xml")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId2",
                     Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                     Target="docProps/core.xml")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId3",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                     Target="docProps/app.xml")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_presentation_xml(self) -> bytes:
        """
        Create ppt/presentation.xml - main presentation structure.
        Defines slide size, slide IDs, and references to slides.
        """
        root = ET.Element(f"{{{self.NS['p']}}}presentation",
                         attrib={
                             'saveSubsetFonts': '1',
                             'autoCompressPictures': '0'
                         })
        
        # Slide master ID list
        sld_master_id_lst = ET.SubElement(root, f"{{{self.NS['p']}}}sldMasterIdLst")
        ET.SubElement(sld_master_id_lst, f"{{{self.NS['p']}}}sldMasterId",
                     attrib={
                         'id': '2147483648',
                         f"{{{self.NS['r']}}}id": 'rId1'
                     })
        
        # Slide ID list - references all slides
        sld_id_lst = ET.SubElement(root, f"{{{self.NS['p']}}}sldIdLst")
        for i in range(1, len(self.slides) + 1):
            ET.SubElement(sld_id_lst, f"{{{self.NS['p']}}}sldId",
                         attrib={
                             'id': str(256 + i),
                             f"{{{self.NS['r']}}}id": f"rId{i + 1}"
                         })
        
        # Slide size with type attribute
        sld_sz = ET.SubElement(root, f"{{{self.NS['p']}}}sldSz",
                              attrib={
                                  'cx': str(self.width),
                                  'cy': str(self.height),
                                  'type': 'screen4x3'
                              })
        
        # Notes size
        notes_sz = ET.SubElement(root, f"{{{self.NS['p']}}}notesSz",
                                attrib={
                                    'cx': str(self.width),
                                    'cy': str(self.height)
                                })
        
        # Default text style (REQUIRED by PowerPoint)
        self._add_default_text_style(root)
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _add_default_text_style(self, parent):
        """
        Add defaultTextStyle element with multiple paragraph levels.
        This is required by PowerPoint for proper text rendering.
        """
        def_text_style = ET.SubElement(parent, f"{{{self.NS['p']}}}defaultTextStyle")
        
        # Default paragraph properties
        def_ppr = ET.SubElement(def_text_style, f"{{{self.NS['a']}}}defPPr")
        def_rpr = ET.SubElement(def_ppr, f"{{{self.NS['a']}}}defRPr", lang="en-US")
        
        # Define 9 paragraph levels
        margins = [0, 457200, 914400, 1371600, 1828800, 2286000, 2743200, 3200400, 3657600]
        for level_num, margin in enumerate(margins, 1):
            lvl_ppr = ET.SubElement(def_text_style, f"{{{self.NS['a']}}}lvl{level_num}pPr",
                                   attrib={
                                       'marL': str(margin),
                                       'algn': 'l',
                                       'defTabSz': '457200',
                                       'rtl': '0',
                                       'eaLnBrk': '1',
                                       'latinLnBrk': '0',
                                       'hangingPunct': '1'
                                   })
            
            lvl_def_rpr = ET.SubElement(lvl_ppr, f"{{{self.NS['a']}}}defRPr",
                                       attrib={'sz': '1800', 'kern': '1200'})
            
            # Solid fill with scheme color
            solid_fill = ET.SubElement(lvl_def_rpr, f"{{{self.NS['a']}}}solidFill")
            ET.SubElement(solid_fill, f"{{{self.NS['a']}}}schemeClr", val='tx1')
            
            # Font typefaces
            ET.SubElement(lvl_def_rpr, f"{{{self.NS['a']}}}latin", typeface='+mn-lt')
            ET.SubElement(lvl_def_rpr, f"{{{self.NS['a']}}}ea", typeface='+mn-ea')
            ET.SubElement(lvl_def_rpr, f"{{{self.NS['a']}}}cs", typeface='+mn-cs')
    
    def _create_presentation_rels_xml(self) -> bytes:
        """
        Create ppt/_rels/presentation.xml.rels
        Defines relationships from presentation to slides, master, and theme.
        """
        root = ET.Element(f"{{{self.NS['rel']}}}Relationships")
        
        # Relationship to slide master
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId1",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
                     Target="slideMasters/slideMaster1.xml")
        
        # Relationships to slides
        for i in range(1, len(self.slides) + 1):
            ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                         Id=f"rId{i + 1}",
                         Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                         Target=f"slides/slide{i}.xml")
        
        # Relationships to view and presentation properties
        next_id = len(self.slides) + 2
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id=f"rId{next_id}",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
                     Target="viewProps.xml")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id=f"rId{next_id + 1}",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
                     Target="presProps.xml")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id=f"rId{next_id + 2}",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
                     Target="tableStyles.xml")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_slide_xml(self, slide_data: Dict[str, Any]) -> bytes:
        """
        Create ppt/slides/slide*.xml - individual slide content.
        This is where the actual content (text, shapes, etc.) is defined.
        """
        root = ET.Element(f"{{{self.NS['p']}}}sld")
        
        # Common slide data
        c_sld_data = ET.SubElement(root, f"{{{self.NS['p']}}}cSld")
        sp_tree = ET.SubElement(c_sld_data, f"{{{self.NS['p']}}}spTree")
        
        # Group shape properties (required)
        nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}nvGrpSpPr")
        c_nv_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvPr",
                               id="1", name="")
        c_nv_grp_sp_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvGrpSpPr")
        nv_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}nvPr")
        
        grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}grpSpPr")
        xfrm = ET.SubElement(grp_sp_pr, f"{{{self.NS['a']}}}xfrm")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}off", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}ext", cx="0", cy="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chOff", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chExt",
                     cx=str(self.width), cy=str(self.height))
        
        # Add shapes based on slide type
        shape_id = 2
        for shape in slide_data.get('shapes', []):
            shape_elem = self._create_shape_xml(shape, shape_id)
            sp_tree.append(shape_elem)
            shape_id += 1
        
        # Color map (required)
        clr_map_ovr = ET.SubElement(root, f"{{{self.NS['p']}}}clrMapOvr")
        master_clr_mapping = ET.SubElement(clr_map_ovr, f"{{{self.NS['a']}}}masterClrMapping")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_shape_xml(self, shape: Dict[str, Any], shape_id: int) -> ET.Element:
        """
        Create a shape element (text box, etc.) with content.
        """
        sp = ET.Element(f"{{{self.NS['p']}}}sp")
        
        # Non-visual shape properties
        nv_sp_pr = ET.SubElement(sp, f"{{{self.NS['p']}}}nvSpPr")
        c_nv_pr = ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}cNvPr",
                               id=str(shape_id), name=f"TextBox {shape_id}")
        c_nv_sp_pr = ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}cNvSpPr",
                                  txBox="1")
        nv_pr = ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}nvPr")
        
        # Visual shape properties (position and size)
        sp_pr = ET.SubElement(sp, f"{{{self.NS['p']}}}spPr")
        xfrm = ET.SubElement(sp_pr, f"{{{self.NS['a']}}}xfrm")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}off",
                     x=str(shape.get('x', 0)),
                     y=str(shape.get('y', 0)))
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}ext",
                     cx=str(shape.get('width', 0)),
                     cy=str(shape.get('height', 0)))
        
        # Preset geometry (rectangle)
        prst_geom = ET.SubElement(sp_pr, f"{{{self.NS['a']}}}prstGeom",
                                 prst="rect")
        av_lst = ET.SubElement(prst_geom, f"{{{self.NS['a']}}}avLst")
        
        # Text body
        tx_body = ET.SubElement(sp, f"{{{self.NS['p']}}}txBody")
        body_pr = ET.SubElement(tx_body, f"{{{self.NS['a']}}}bodyPr",
                               wrap="square", rtlCol="0")
        ET.SubElement(body_pr, f"{{{self.NS['a']}}}spAutoFit")
        
        lst_style = ET.SubElement(tx_body, f"{{{self.NS['a']}}}lstStyle")
        
        # Add paragraphs
        for para_text in shape.get('text', []):
            p = ET.SubElement(tx_body, f"{{{self.NS['a']}}}p")
            
            # Paragraph properties
            p_pr = ET.SubElement(p, f"{{{self.NS['a']}}}pPr",
                                algn=shape.get('align', 'l'))
            
            # Text run
            r = ET.SubElement(p, f"{{{self.NS['a']}}}r")
            r_pr = ET.SubElement(r, f"{{{self.NS['a']}}}rPr",
                                lang="en-US",
                                sz=str(shape.get('font_size', 1800)),
                                b=str(int(shape.get('bold', False))))
            
            # Solid fill (text color)
            solid_fill = ET.SubElement(r_pr, f"{{{self.NS['a']}}}solidFill")
            srgb_clr = ET.SubElement(solid_fill, f"{{{self.NS['a']}}}srgbClr",
                                    val=shape.get('color', '000000'))
            
            # Text content
            t = ET.SubElement(r, f"{{{self.NS['a']}}}t")
            t.text = para_text
            
            # End paragraph properties
            end_para_r_pr = ET.SubElement(p, f"{{{self.NS['a']}}}endParaRPr",
                                         lang="en-US")
        
        return sp
    
    def _create_slide_rels_xml(self, slide_num: int) -> bytes:
        """
        Create ppt/slides/_rels/slide*.xml.rels
        Defines relationships from slide to slide layout.
        """
        root = ET.Element(f"{{{self.NS['rel']}}}Relationships")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId1",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                     Target="../slideLayouts/slideLayout1.xml")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_slide_layout_xml(self) -> bytes:
        """
        Create ppt/slideLayouts/slideLayout1.xml
        Basic blank slide layout.
        """
        root = ET.Element(f"{{{self.NS['p']}}}sldLayout",
                         attrib={
                             'type': 'blank',
                             'preserve': '1',
                         })
        
        c_sld_data = ET.SubElement(root, f"{{{self.NS['p']}}}cSld",
                                  name="Blank")
        sp_tree = ET.SubElement(c_sld_data, f"{{{self.NS['p']}}}spTree")
        
        # Group shape properties
        nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}nvGrpSpPr")
        c_nv_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvPr",
                               id="1", name="")
        c_nv_grp_sp_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvGrpSpPr")
        nv_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}nvPr")
        
        grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}grpSpPr")
        xfrm = ET.SubElement(grp_sp_pr, f"{{{self.NS['a']}}}xfrm")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}off", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}ext", cx="0", cy="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chOff", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chExt",
                     cx=str(self.width), cy=str(self.height))
        
        # Color map override
        clr_map_ovr = ET.SubElement(root, f"{{{self.NS['p']}}}clrMapOvr")
        master_clr_mapping = ET.SubElement(clr_map_ovr, f"{{{self.NS['a']}}}masterClrMapping")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_slide_layout_rels_xml(self) -> bytes:
        """
        Create ppt/slideLayouts/_rels/slideLayout1.xml.rels
        """
        root = ET.Element(f"{{{self.NS['rel']}}}Relationships")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId1",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
                     Target="../slideMasters/slideMaster1.xml")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_slide_master_xml(self) -> bytes:
        """
        Create ppt/slideMasters/slideMaster1.xml
        Master template for slides.
        """
        root = ET.Element(f"{{{self.NS['p']}}}sldMaster")
        
        c_sld_data = ET.SubElement(root, f"{{{self.NS['p']}}}cSld")
        sp_tree = ET.SubElement(c_sld_data, f"{{{self.NS['p']}}}spTree")
        
        # Group shape properties
        nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}nvGrpSpPr")
        c_nv_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvPr",
                               id="1", name="")
        c_nv_grp_sp_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvGrpSpPr")
        nv_pr = ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}nvPr")
        
        grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}grpSpPr")
        xfrm = ET.SubElement(grp_sp_pr, f"{{{self.NS['a']}}}xfrm")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}off", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}ext", cx="0", cy="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chOff", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chExt",
                     cx=str(self.width), cy=str(self.height))
        
        # Color map
        clr_map = ET.SubElement(root, f"{{{self.NS['p']}}}clrMap",
                               bg1="lt1", tx1="dk1", bg2="lt2", tx2="dk2",
                               accent1="accent1", accent2="accent2", accent3="accent3",
                               accent4="accent4", accent5="accent5", accent6="accent6",
                               hlink="hlink", folHlink="folHlink")
        
        # Slide layout ID list
        sld_layout_id_lst = ET.SubElement(root, f"{{{self.NS['p']}}}sldLayoutIdLst")
        ET.SubElement(sld_layout_id_lst, f"{{{self.NS['p']}}}sldLayoutId",
                     id="2147483649",
                     attrib={f"{{{self.NS['r']}}}id": "rId1"})
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_slide_master_rels_xml(self) -> bytes:
        """
        Create ppt/slideMasters/_rels/slideMaster1.xml.rels
        """
        root = ET.Element(f"{{{self.NS['rel']}}}Relationships")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId1",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                     Target="../slideLayouts/slideLayout1.xml")
        
        ET.SubElement(root, f"{{{self.NS['rel']}}}Relationship",
                     Id="rId2",
                     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                     Target="../theme/theme1.xml")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_theme_xml(self) -> bytes:
        """
        Create ppt/theme/theme1.xml
        Basic theme definition.
        """
        root = ET.Element(f"{{{self.NS['a']}}}theme",
                         attrib={
                             'name': 'Office Theme',
                         })
        
        theme_elements = ET.SubElement(root, f"{{{self.NS['a']}}}themeElements")
        
        # Color scheme
        clr_scheme = ET.SubElement(theme_elements, f"{{{self.NS['a']}}}clrScheme",
                                  name="Office")
        
        color_defs = [
            ('dk1', '000000'), ('lt1', 'FFFFFF'),
            ('dk2', '44546A'), ('lt2', 'E7E6E6'),
            ('accent1', '4472C4'), ('accent2', 'ED7D31'),
            ('accent3', 'A5A5A5'), ('accent4', 'FFC000'),
            ('accent5', '5B9BD5'), ('accent6', '70AD47'),
            ('hlink', '0563C1'), ('folHlink', '954F72'),
        ]
        
        for name, rgb in color_defs:
            color_elem = ET.SubElement(clr_scheme, f"{{{self.NS['a']}}}{name}")
            ET.SubElement(color_elem, f"{{{self.NS['a']}}}srgbClr", val=rgb)
        
        # Font scheme (simplified)
        font_scheme = ET.SubElement(theme_elements, f"{{{self.NS['a']}}}fontScheme",
                                   name="Office")
        major_font = ET.SubElement(font_scheme, f"{{{self.NS['a']}}}majorFont")
        major_latin = ET.SubElement(major_font, f"{{{self.NS['a']}}}latin",
                             typeface="Calibri Light")
        
        minor_font = ET.SubElement(font_scheme, f"{{{self.NS['a']}}}minorFont")
        minor_latin = ET.SubElement(minor_font, f"{{{self.NS['a']}}}latin",
                             typeface="Calibri")
        
        # Format scheme (simplified)
        fmt_scheme = ET.SubElement(theme_elements, f"{{{self.NS['a']}}}fmtScheme",
                                  name="Office")
        fill_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}fillStyleLst")
        solid_fill = ET.SubElement(fill_style_lst, f"{{{self.NS['a']}}}solidFill")
        scheme_clr = ET.SubElement(solid_fill, f"{{{self.NS['a']}}}schemeClr",
                                  val="phClr")
        
        ln_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}lnStyleLst")
        ln = ET.SubElement(ln_style_lst, f"{{{self.NS['a']}}}ln", w="6350", cap="flat",
                          cmpd="sng", algn="ctr")
        
        effect_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}effectStyleLst")
        effect_style = ET.SubElement(effect_style_lst, f"{{{self.NS['a']}}}effectStyle")
        effect_lst = ET.SubElement(effect_style, f"{{{self.NS['a']}}}effectLst")
        
        bg_fill_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}bgFillStyleLst")
        solid_fill = ET.SubElement(bg_fill_style_lst, f"{{{self.NS['a']}}}solidFill")
        scheme_clr = ET.SubElement(solid_fill, f"{{{self.NS['a']}}}schemeClr",
                                  val="phClr")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_app_properties_xml(self) -> bytes:
        """
        Create docProps/app.xml - application properties.
        Required by PowerPoint for proper compatibility.
        """
        root = ET.Element("Properties",
                         attrib={
                             'xmlns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                             'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
                         })
        
        ET.SubElement(root, "TotalTime").text = "0"
        ET.SubElement(root, "Words").text = "0"
        ET.SubElement(root, "Application").text = "Python PPTX Generator"
        ET.SubElement(root, "PresentationFormat").text = "On-screen Show (4:3)"
        ET.SubElement(root, "Paragraphs").text = "0"
        ET.SubElement(root, "Slides").text = str(len(self.slides))
        ET.SubElement(root, "Notes").text = "0"
        ET.SubElement(root, "HiddenSlides").text = "0"
        ET.SubElement(root, "MMClips").text = "0"
        ET.SubElement(root, "ScaleCrop").text = "false"
        ET.SubElement(root, "Company").text = ""
        ET.SubElement(root, "LinksUpToDate").text = "false"
        ET.SubElement(root, "SharedDoc").text = "false"
        ET.SubElement(root, "HyperlinksChanged").text = "false"
        ET.SubElement(root, "AppVersion").text = "16.0000"
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_core_properties_xml(self) -> bytes:
        """
        Create docProps/core.xml - core document properties.
        Required for proper document metadata.
        """
        from datetime import datetime
        
        root = ET.Element("cp:coreProperties",
                         attrib={
                             'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                             'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
                             'xmlns:dcterms': 'http://purl.org/dc/terms/',
                             'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
                             'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance'
                         })
        
        ET.SubElement(root, "dc:title").text = "Presentation"
        ET.SubElement(root, "dc:creator").text = "Python PPTX Generator"
        ET.SubElement(root, "cp:lastModifiedBy").text = "Python PPTX Generator"
        ET.SubElement(root, "cp:revision").text = "1"
        
        # Add timestamps
        now = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
        created = ET.SubElement(root, "dcterms:created",
                               attrib={'xsi:type': 'dcterms:W3CDTF'})
        created.text = now
        modified = ET.SubElement(root, "dcterms:modified",
                                attrib={'xsi:type': 'dcterms:W3CDTF'})
        modified.text = now
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_view_props_xml(self) -> bytes:
        """
        Create ppt/viewProps.xml - view properties.
        """
        root = ET.Element(f"{{{self.NS['p']}}}viewPr")
        
        normal_view_pr = ET.SubElement(root, f"{{{self.NS['p']}}}normalViewPr")
        restored_left = ET.SubElement(normal_view_pr, f"{{{self.NS['p']}}}restoredLeft",
                                     sz="15620")
        restored_top = ET.SubElement(normal_view_pr, f"{{{self.NS['p']}}}restoredTop",
                                    sz="94660")
        
        grid_spacing = ET.SubElement(root, f"{{{self.NS['p']}}}gridSpacing",
                                    cx="72008", cy="72008")
        
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_pres_props_xml(self) -> bytes:
        """
        Create ppt/presProps.xml - presentation properties.
        """
        root = ET.Element(f"{{{self.NS['p']}}}presentationPr")
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _create_table_styles_xml(self) -> bytes:
        """
        Create ppt/tableStyles.xml - table style definitions.
        Required by PowerPoint even if no tables are used.
        """
        root = ET.Element(f"{{{self.NS['a']}}}tblStyleLst",
                         attrib={
                             'def': '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'
                         })
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def add_slide(self, shapes: List[Dict[str, Any]]) -> None:
        """
        Add a slide with the specified shapes.
        
        Args:
            shapes: List of shape dictionaries with properties:
                - x, y: Position in EMUs
                - width, height: Size in EMUs
                - text: List of text strings (paragraphs)
                - font_size: Size in points * 100 (default: 1800 = 18pt)
                - bold: Boolean (default: False)
                - align: Alignment ('l', 'c', 'r')
                - color: RGB hex color (default: '000000')
        """
        self.slide_counter += 1
        self.slides.append({
            'number': self.slide_counter,
            'shapes': shapes
        })
    
    def add_title_slide(self, title: str, subtitle: str = "") -> None:
        """
        Add a title slide.
        
        Args:
            title: Main title text
            subtitle: Optional subtitle text
        """
        shapes = []
        
        # Title text box (centered, top area)
        shapes.append({
            'x': 457200,  # 0.5 inches
            'y': 1828800,  # 2 inches
            'width': 8229600,  # 9 inches
            'height': 1371600,  # 1.5 inches
            'text': [title],
            'font_size': 4400,  # 44pt
            'bold': True,
            'align': 'ctr',
            'color': '000000',
        })
        
        # Subtitle text box (if provided)
        if subtitle:
            shapes.append({
                'x': 457200,
                'y': 3657600,  # 4 inches
                'width': 8229600,
                'height': 914400,  # 1 inch
                'text': [subtitle],
                'font_size': 2400,  # 24pt
                'bold': False,
                'align': 'ctr',
                'color': '666666',
            })
        
        self.add_slide(shapes)
    
    def add_content_slide(self, title: str, bullet_points: List[str]) -> None:
        """
        Add a content slide with title and bullet points.
        
        Args:
            title: Slide title
            bullet_points: List of bullet point texts
        """
        shapes = []
        
        # Title
        shapes.append({
            'x': 457200,
            'y': 457200,  # 0.5 inches
            'width': 8229600,
            'height': 914400,
            'text': [title],
            'font_size': 3200,  # 32pt
            'bold': True,
            'align': 'l',
            'color': '000000',
        })
        
        # Content area with bullet points
        bullet_text = []
        for point in bullet_points:
            bullet_text.append(f"• {point}")
        
        shapes.append({
            'x': 457200,
            'y': 1828800,  # 2 inches
            'width': 8229600,
            'height': 4571000,  # 5 inches
            'text': bullet_text,
            'font_size': 1800,  # 18pt
            'bold': False,
            'align': 'l',
            'color': '000000',
        })
        
        self.add_slide(shapes)
    
    def save(self, filename: str) -> None:
        """
        Save the presentation to a .pptx file.
        
        This creates a ZIP archive containing all the XML files that make up
        the PowerPoint presentation.
        
        Args:
            filename: Output filename (should end with .pptx)
        """
        with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml
            zf.writestr('[Content_Types].xml', self._create_content_types_xml())
            
            # _rels/.rels
            zf.writestr('_rels/.rels', self._create_rels_xml())
            
            # ppt/presentation.xml
            zf.writestr('ppt/presentation.xml', self._create_presentation_xml())
            
            # ppt/_rels/presentation.xml.rels
            zf.writestr('ppt/_rels/presentation.xml.rels',
                       self._create_presentation_rels_xml())
            
            # ppt/slides/slide*.xml and rels
            for slide in self.slides:
                slide_num = slide['number']
                zf.writestr(f'ppt/slides/slide{slide_num}.xml',
                           self._create_slide_xml(slide))
                zf.writestr(f'ppt/slides/_rels/slide{slide_num}.xml.rels',
                           self._create_slide_rels_xml(slide_num))
            
            # ppt/slideLayouts/slideLayout1.xml
            zf.writestr('ppt/slideLayouts/slideLayout1.xml',
                       self._create_slide_layout_xml())
            
            # ppt/slideLayouts/_rels/slideLayout1.xml.rels
            zf.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels',
                       self._create_slide_layout_rels_xml())
            
            # ppt/slideMasters/slideMaster1.xml
            zf.writestr('ppt/slideMasters/slideMaster1.xml',
                       self._create_slide_master_xml())
            
            # ppt/slideMasters/_rels/slideMaster1.xml.rels
            zf.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels',
                       self._create_slide_master_rels_xml())
            
            # ppt/theme/theme1.xml
            zf.writestr('ppt/theme/theme1.xml', self._create_theme_xml())
            
            # ppt/viewProps.xml
            zf.writestr('ppt/viewProps.xml', self._create_view_props_xml())
            
            # ppt/presProps.xml
            zf.writestr('ppt/presProps.xml', self._create_pres_props_xml())
            
            # ppt/tableStyles.xml
            zf.writestr('ppt/tableStyles.xml', self._create_table_styles_xml())
            
            # docProps/core.xml
            zf.writestr('docProps/core.xml', self._create_core_properties_xml())
            
            # docProps/app.xml
            zf.writestr('docProps/app.xml', self._create_app_properties_xml())
        
        print(f"✓ PowerPoint created: {filename}")
        print(f"  - {len(self.slides)} slides")
        print(f"  - Built from scratch using XML")
