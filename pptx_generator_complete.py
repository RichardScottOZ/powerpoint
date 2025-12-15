"""
Complete From-Scratch PowerPoint Generator

This creates PowerPoint presentations by directly generating ALL required XML
files including proper slide layouts with placeholders. Built entirely from
scratch using only Python's built-in zipfile and xml modules.

Based on analysis of python-pptx generated files to understand the complete
structure required by PowerPoint/LibreOffice.
"""

import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List, Dict, Tuple


class CompletePPTXGenerator:
    """
    From-scratch PowerPoint generator with complete slide layouts and placeholders.
    """
    
    # XML Namespaces
    NS = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'dcterms': 'http://purl.org/dc/terms/',
        'dcmitype': 'http://purl.org/dc/dcmitype/',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
        'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    }
    
    def __init__(self):
        """Initialize the presentation."""
        # Slide dimensions in EMUs (English Metric Units)
        # Standard 4:3 aspect ratio: 10" x 7.5"
        self.width = 9144000  # 10 inches
        self.height = 6858000  # 7.5 inches
        
        self.slides = []
        self.slide_counter = 0
        
        # Register namespaces - but NOT the ones that should be default (ct and rel)
        for prefix, uri in self.NS.items():
            if prefix not in ['ct', 'rel']:  # Don't register these - they'll be default namespaces
                ET.register_namespace(prefix, uri)
    
    def _xml_to_bytes(self, root: ET.Element) -> bytes:
        """Convert XML element to bytes with proper declaration."""
        xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=False)
        declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        return (declaration + xml_bytes.decode('utf-8')).encode('utf-8')
    
    def _create_content_types(self) -> bytes:
        """Create [Content_Types].xml with default namespace (no prefix)"""
        # Manually construct XML to avoid namespace prefix issues
        declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        xml_parts = ['<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
        
        # Default types
        for ext, ctype in [
            ('rels', 'application/vnd.openxmlformats-package.relationships+xml'),
            ('xml', 'application/xml'),
            ('jpeg', 'image/jpeg'),
        ]:
            xml_parts.append(f'<Default Extension="{ext}" ContentType="{ctype}"/>')
        
        # Override types
        overrides = [
            ('/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'),
            ('/ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'),
            ('/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'),
            ('/ppt/presProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'),
            ('/ppt/viewProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml'),
            ('/ppt/tableStyles.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'),
            ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
            ('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'),
        ]
        
        # Add slide layouts
        for i in range(1, 12):
            overrides.append((f'/ppt/slideLayouts/slideLayout{i}.xml', 
                            'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'))
        
        # Add slides
        for i in range(1, len(self.slides) + 1):
            overrides.append((f'/ppt/slides/slide{i}.xml',
                            'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'))
        
        for part_name, ctype in overrides:
            xml_parts.append(f'<Override PartName="{part_name}" ContentType="{ctype}"/>')
        
        xml_parts.append('</Types>')
        
        return (declaration + ''.join(xml_parts)).encode('utf-8')
    
    def _create_package_rels(self) -> bytes:
        """Create _rels/.rels with default namespace (no prefix)"""
        # Manually construct XML to avoid namespace prefix issues
        declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        xml_parts = ['<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        
        rels = [
            ('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'ppt/presentation.xml'),
            ('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'),
            ('rId3', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'docProps/app.xml'),
        ]
        
        for rid, rtype, target in rels:
            xml_parts.append(f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>')
        
        xml_parts.append('</Relationships>')
        return (declaration + ''.join(xml_parts)).encode('utf-8')
    
    def _create_core_props(self) -> bytes:
        """Create docProps/core.xml"""
        root = ET.Element(f"{{{self.NS['cp']}}}coreProperties")
        root.set(f"{{{self.NS['xsi']}}}schemaLocation",
                "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
        
        ET.SubElement(root, f"{{{self.NS['dc']}}}title").text = "Presentation"
        ET.SubElement(root, f"{{{self.NS['dc']}}}creator").text = "Python Generator"
        ET.SubElement(root, f"{{{self.NS['cp']}}}lastModifiedBy").text = "Python Generator"
        ET.SubElement(root, f"{{{self.NS['cp']}}}revision").text = "1"
        
        now = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
        ET.SubElement(root, f"{{{self.NS['dcterms']}}}created",
                     {f"{{{self.NS['xsi']}}}type": "dcterms:W3CDTF"}).text = now
        ET.SubElement(root, f"{{{self.NS['dcterms']}}}modified",
                     {f"{{{self.NS['xsi']}}}type": "dcterms:W3CDTF"}).text = now
        
        return self._xml_to_bytes(root)
    
    def _create_app_props(self) -> bytes:
        """Create docProps/app.xml"""
        root = ET.Element(f"{{{self.NS['ep']}}}Properties")
        
        ET.SubElement(root, f"{{{self.NS['ep']}}}Application").text = "Python Generator"
        ET.SubElement(root, f"{{{self.NS['ep']}}}PresentationFormat").text = "On-screen Show (4:3)"
        ET.SubElement(root, f"{{{self.NS['ep']}}}Slides").text = str(len(self.slides))
        
        # Title vector
        titles_of_parts = ET.SubElement(root, f"{{{self.NS['ep']}}}TitlesOfParts")
        vt_vector = ET.SubElement(titles_of_parts, f"{{{self.NS['vt']}}}vector",
                                 size=str(len(self.slides)), baseType="lpstr")
        for i in range(len(self.slides)):
            ET.SubElement(vt_vector, f"{{{self.NS['vt']}}}lpstr").text = f"Slide {i+1}"
        
        return self._xml_to_bytes(root)
    
    def _create_presentation_xml(self) -> bytes:
        """Create ppt/presentation.xml"""
        root = ET.Element(f"{{{self.NS['p']}}}presentation", {
            'saveSubsetFonts': '1',
            'autoCompressPictures': '0'
        })
        
        # Slide master ID list
        sld_master_id_lst = ET.SubElement(root, f"{{{self.NS['p']}}}sldMasterIdLst")
        ET.SubElement(sld_master_id_lst, f"{{{self.NS['p']}}}sldMasterId",
                     {f"{{{self.NS['r']}}}id": "rId1"}, id="2147483648")
        
        # Slide ID list
        sld_id_lst = ET.SubElement(root, f"{{{self.NS['p']}}}sldIdLst")
        for i in range(len(self.slides)):
            ET.SubElement(sld_id_lst, f"{{{self.NS['p']}}}sldId",
                         {f"{{{self.NS['r']}}}id": f"rId{i+2}"}, id=str(256 + i))
        
        # Slide size
        sld_sz = ET.SubElement(root, f"{{{self.NS['p']}}}sldSz",
                              cx=str(self.width), cy=str(self.height), type="screen4x3")
        
        # Notes size
        notes_sz = ET.SubElement(root, f"{{{self.NS['p']}}}notesSz",
                                cx=str(self.height), cy=str(self.width))
        
        # Default text style
        default_text_style = ET.SubElement(root, f"{{{self.NS['p']}}}defaultTextStyle")
        for lvl in range(1, 10):
            lvl_ppr = ET.SubElement(default_text_style, f"{{{self.NS['a']}}}lvl{lvl}pPr",
                                   marL=str((lvl-1)*457200), algn="l", defTabSz="914400",
                                   rtl="0", eaLnBrk="1", latinLnBrk="0", hangingPunct="1")
            def_rpr = ET.SubElement(lvl_ppr, f"{{{self.NS['a']}}}defRPr",
                                   sz="1800" if lvl <= 2 else "1600", kern="1200")
            solid_fill = ET.SubElement(def_rpr, f"{{{self.NS['a']}}}solidFill")
            ET.SubElement(solid_fill, f"{{{self.NS['a']}}}schemeClr", val="tx1")
            latin = ET.SubElement(def_rpr, f"{{{self.NS['a']}}}latin",
                                 typeface="+mn-lt")
            ea = ET.SubElement(def_rpr, f"{{{self.NS['a']}}}ea",
                              typeface="+mn-ea")
            cs = ET.SubElement(def_rpr, f"{{{self.NS['a']}}}cs",
                              typeface="+mn-cs")
        
        return self._xml_to_bytes(root)
    
    def _create_presentation_rels(self) -> bytes:
        """Create ppt/_rels/presentation.xml.rels with default namespace (no prefix)"""
        # Manually construct XML to avoid namespace prefix issues
        declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        xml_parts = ['<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        
        # Relationships
        rels = [
            ('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', 'slideMasters/slideMaster1.xml'),
            ('rId4', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps', 'presProps.xml'),
            ('rId5', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps', 'viewProps.xml'),
            ('rId6', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles', 'tableStyles.xml'),
        ]
        
        # Add slide relationships
        for i in range(len(self.slides)):
            rels.append((f'rId{i+2}', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 
                        f'slides/slide{i+1}.xml'))
        
        for rid, rtype, target in rels:
            xml_parts.append(f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>')
        
        xml_parts.append('</Relationships>')
        return (declaration + ''.join(xml_parts)).encode('utf-8')
    
    def _create_pres_props(self) -> bytes:
        """Create ppt/presProps.xml"""
        root = ET.Element(f"{{{self.NS['p']}}}presentationPr")
        return self._xml_to_bytes(root)
    
    def _create_view_props(self) -> bytes:
        """Create ppt/viewProps.xml"""
        root = ET.Element(f"{{{self.NS['p']}}}viewPr")
        
        normal_view_pr = ET.SubElement(root, f"{{{self.NS['p']}}}normalViewPr")
        restored_left = ET.SubElement(normal_view_pr, f"{{{self.NS['p']}}}restoredLeft", sz="15620")
        restored_top = ET.SubElement(normal_view_pr, f"{{{self.NS['p']}}}restoredTop", sz="94660")
        
        grid_spacing = ET.SubElement(root, f"{{{self.NS['p']}}}gridSpacing", cx="72008", cy="72008")
        
        return self._xml_to_bytes(root)
    
    def _create_table_styles(self) -> bytes:
        """Create ppt/tableStyles.xml"""
        root = ET.Element(f"{{{self.NS['p']}}}tblStyleLst", def_="{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}")
        return self._xml_to_bytes(root)
    
    def _create_theme(self) -> bytes:
        """Create ppt/theme/theme1.xml"""
        root = ET.Element(f"{{{self.NS['a']}}}theme", name="Office Theme")
        
        # Theme elements
        theme_elements = ET.SubElement(root, f"{{{self.NS['a']}}}themeElements")
        
        # Color scheme
        clr_scheme = ET.SubElement(theme_elements, f"{{{self.NS['a']}}}clrScheme", name="Office")
        colors = [
            ('dk1', '000000'), ('lt1', 'FFFFFF'), ('dk2', '44546A'), ('lt2', 'E7E6E6'),
            ('accent1', '5B9BD5'), ('accent2', 'ED7D31'), ('accent3', 'A5A5A5'),
            ('accent4', 'FFC000'), ('accent5', '4472C4'), ('accent6', '70AD47'),
            ('hlink', '0563C1'), ('folHlink', '954F72')
        ]
        for name, rgb in colors:
            clr = ET.SubElement(clr_scheme, f"{{{self.NS['a']}}}{name}")
            ET.SubElement(clr, f"{{{self.NS['a']}}}srgbClr", val=rgb)
        
        # Font scheme
        font_scheme = ET.SubElement(theme_elements, f"{{{self.NS['a']}}}fontScheme", name="Office")
        major_font = ET.SubElement(font_scheme, f"{{{self.NS['a']}}}majorFont")
        ET.SubElement(major_font, f"{{{self.NS['a']}}}latin", typeface="Calibri Light", panose="020F0302020204030204")
        ET.SubElement(major_font, f"{{{self.NS['a']}}}ea", typeface="")
        ET.SubElement(major_font, f"{{{self.NS['a']}}}cs", typeface="")
        
        minor_font = ET.SubElement(font_scheme, f"{{{self.NS['a']}}}minorFont")
        ET.SubElement(minor_font, f"{{{self.NS['a']}}}latin", typeface="Calibri", panose="020F0502020204030204")
        ET.SubElement(minor_font, f"{{{self.NS['a']}}}ea", typeface="")
        ET.SubElement(minor_font, f"{{{self.NS['a']}}}cs", typeface="")
        
        # Format scheme
        fmt_scheme = ET.SubElement(theme_elements, f"{{{self.NS['a']}}}fmtScheme", name="Office")
        fill_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}fillStyleLst")
        solid_fill = ET.SubElement(fill_style_lst, f"{{{self.NS['a']}}}solidFill")
        ET.SubElement(solid_fill, f"{{{self.NS['a']}}}schemeClr", val="phClr")
        
        line_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}lineStyleLst")
        ln = ET.SubElement(line_style_lst, f"{{{self.NS['a']}}}ln", w="9525", cap="flat", cmpd="sng", algn="ctr")
        solid_fill2 = ET.SubElement(ln, f"{{{self.NS['a']}}}solidFill")
        ET.SubElement(solid_fill2, f"{{{self.NS['a']}}}schemeClr", val="phClr")
        
        effect_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}effectStyleLst")
        effect_style = ET.SubElement(effect_style_lst, f"{{{self.NS['a']}}}effectStyle")
        effect_lst = ET.SubElement(effect_style, f"{{{self.NS['a']}}}effectLst")
        
        bg_fill_style_lst = ET.SubElement(fmt_scheme, f"{{{self.NS['a']}}}bgFillStyleLst")
        solid_fill3 = ET.SubElement(bg_fill_style_lst, f"{{{self.NS['a']}}}solidFill")
        ET.SubElement(solid_fill3, f"{{{self.NS['a']}}}schemeClr", val="phClr")
        
        return self._xml_to_bytes(root)
    
    def _create_slide_master(self) -> bytes:
        """Create ppt/slideMasters/slideMaster1.xml - simplified but complete version"""
        root = ET.Element(f"{{{self.NS['p']}}}sldMaster")
        
        # Common slide data
        cSld = ET.SubElement(root, f"{{{self.NS['p']}}}cSld")
        bg = ET.SubElement(cSld, f"{{{self.NS['p']}}}bg")
        bg_ref = ET.SubElement(bg, f"{{{self.NS['p']}}}bgRef", idx="1001")
        ET.SubElement(bg_ref, f"{{{self.NS['a']}}}schemeClr", val="bg1")
        
        sp_tree = ET.SubElement(cSld, f"{{{self.NS['p']}}}spTree")
        
        # Group shape properties
        nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}nvGrpSpPr")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvPr", id="1", name="")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvGrpSpPr")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}nvPr")
        
        grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}grpSpPr")
        xfrm = ET.SubElement(grp_sp_pr, f"{{{self.NS['a']}}}xfrm")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}off", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}ext", cx="0", cy="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chOff", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chExt", cx="0", cy="0")
        
        # Color map
        clr_map = ET.SubElement(root, f"{{{self.NS['p']}}}clrMap",
                               bg1="lt1", tx1="dk1", bg2="lt2", tx2="dk2",
                               accent1="accent1", accent2="accent2", accent3="accent3",
                               accent4="accent4", accent5="accent5", accent6="accent6",
                               hlink="hlink", folHlink="folHlink")
        
        # Slide layout ID list
        sld_layout_id_lst = ET.SubElement(root, f"{{{self.NS['p']}}}sldLayoutIdLst")
        for i in range(1, 12):
            ET.SubElement(sld_layout_id_lst, f"{{{self.NS['p']}}}sldLayoutId",
                         {f"{{{self.NS['r']}}}id": f"rId{i}"}, id=str(2147483649 + i - 1))
        
        # Text styles
        txStyles = ET.SubElement(root, f"{{{self.NS['p']}}}txStyles")
        for style_type in ['titleStyle', 'bodyStyle', 'otherStyle']:
            style = ET.SubElement(txStyles, f"{{{self.NS['p']}}}{style_type}")
            lvl1pPr = ET.SubElement(style, f"{{{self.NS['a']}}}lvl1pPr")
            defRPr = ET.SubElement(lvl1pPr, f"{{{self.NS['a']}}}defRPr")
        
        return self._xml_to_bytes(root)
    
    def _create_slide_master_rels(self) -> bytes:
        """Create ppt/slideMasters/_rels/slideMaster1.xml.rels with default namespace (no prefix)"""
        # Manually construct XML to avoid namespace prefix issues
        declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        xml_parts = ['<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        
        # Layout relationships
        for i in range(1, 12):
            xml_parts.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{i}.xml"/>')
        
        # Theme relationship
        xml_parts.append('<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>')
        
        xml_parts.append('</Relationships>')
        return (declaration + ''.join(xml_parts)).encode('utf-8')
    
    def _create_slide_layout(self, layout_num: int, layout_type: str, layout_name: str,
                            placeholders: List[Tuple[str, str, int]]) -> bytes:
        """
        Create a slide layout XML.
        
        Args:
            layout_num: Layout number
            layout_type: Layout type (title, obj, etc.)
            layout_name: Human-readable name
            placeholders: List of (ph_type, ph_name, ph_id) tuples
        """
        root = ET.Element(f"{{{self.NS['p']}}}sldLayout", type=layout_type, preserve="1")
        
        # Common slide data
        cSld = ET.SubElement(root, f"{{{self.NS['p']}}}cSld", name=layout_name)
        sp_tree = ET.SubElement(cSld, f"{{{self.NS['p']}}}spTree")
        
        # Group shape properties
        nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}nvGrpSpPr")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvPr", id="1", name="")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvGrpSpPr")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}nvPr")
        
        grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}grpSpPr")
        xfrm = ET.SubElement(grp_sp_pr, f"{{{self.NS['a']}}}xfrm")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}off", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}ext", cx="0", cy="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chOff", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chExt", cx="0", cy="0")
        
        # Add placeholders
        for idx, (ph_type, ph_name, ph_id) in enumerate(placeholders):
            sp = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}sp")
            
            # Non-visual shape properties
            nv_sp_pr = ET.SubElement(sp, f"{{{self.NS['p']}}}nvSpPr")
            ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}cNvPr", id=str(ph_id), name=ph_name)
            cnv_sp_pr = ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}cNvSpPr")
            ET.SubElement(cnv_sp_pr, f"{{{self.NS['a']}}}spLocks", noGrp="1")
            nv_pr = ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}nvPr")
            ph_attrs = {'type': ph_type}
            if idx > 0:
                ph_attrs['idx'] = str(idx)
            ET.SubElement(nv_pr, f"{{{self.NS['p']}}}ph", **ph_attrs)
            
            # Shape properties
            ET.SubElement(sp, f"{{{self.NS['p']}}}spPr")
            
            # Text body
            tx_body = ET.SubElement(sp, f"{{{self.NS['p']}}}txBody")
            ET.SubElement(tx_body, f"{{{self.NS['a']}}}bodyPr")
            ET.SubElement(tx_body, f"{{{self.NS['a']}}}lstStyle")
            p = ET.SubElement(tx_body, f"{{{self.NS['a']}}}p")
            end_para_rpr = ET.SubElement(p, f"{{{self.NS['a']}}}endParaRPr", lang="en-US")
        
        return self._xml_to_bytes(root)
    
    def _create_slide_layout_rels(self, layout_num: int) -> bytes:
        """Create ppt/slideLayouts/_rels/slideLayout{n}.xml.rels with default namespace (no prefix)"""
        # Manually construct XML to avoid namespace prefix issues
        declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        xml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        xml += '</Relationships>'
        return (declaration + xml).encode('utf-8')
    
    def _create_all_slide_layouts(self) -> Dict[str, bytes]:
        """Create all 11 standard slide layouts"""
        layouts = {}
        
        # Layout definitions: (num, type, name, placeholders)
        layout_defs = [
            (1, 'title', 'Title Slide', [
                ('ctrTitle', 'Title 1', 2),
                ('subTitle', 'Subtitle 2', 3),
            ]),
            (2, 'obj', 'Title and Content', [
                ('title', 'Title 1', 2),
                ('body', 'Content Placeholder 2', 3),
            ]),
            (3, 'secHead', 'Section Header', [
                ('title', 'Title 1', 2),
                ('body', 'Text Placeholder 2', 3),
            ]),
            (4, 'twoObj', 'Two Content', [
                ('title', 'Title 1', 2),
                ('body', 'Content Placeholder 2', 3),
                ('body', 'Content Placeholder 3', 4),
            ]),
            (5, 'twoTxTwoObj', 'Comparison', [
                ('title', 'Title 1', 2),
                ('body', 'Text Placeholder 2', 3),
                ('body', 'Text Placeholder 3', 4),
            ]),
            (6, 'titleOnly', 'Title Only', [
                ('title', 'Title 1', 2),
            ]),
            (7, 'blank', 'Blank', []),
            (8, 'objTx', 'Content with Caption', [
                ('title', 'Title 1', 2),
                ('body', 'Content Placeholder 2', 3),
            ]),
            (9, 'picTx', 'Picture with Caption', [
                ('title', 'Title 1', 2),
                ('body', 'Content Placeholder 2', 3),
            ]),
            (10, 'vertTx', 'Title and Vertical Text', [
                ('title', 'Title 1', 2),
                ('body', 'Vertical Text Placeholder 2', 3),
            ]),
            (11, 'vertTitleAndTx', 'Vertical Title and Text', [
                ('title', 'Vertical Title 1', 2),
                ('body', 'Vertical Text Placeholder 2', 3),
            ]),
        ]
        
        for num, ltype, name, placeholders in layout_defs:
            layouts[f'ppt/slideLayouts/slideLayout{num}.xml'] = self._create_slide_layout(
                num, ltype, name, placeholders
            )
            layouts[f'ppt/slideLayouts/_rels/slideLayout{num}.xml.rels'] = self._create_slide_layout_rels(num)
        
        return layouts
    
    def _create_slide(self, slide_num: int, layout_num: int, shapes: List[Dict]) -> bytes:
        """
        Create a slide XML.
        
        Args:
            slide_num: Slide number
            layout_num: Layout to use (1-11)
            shapes: List of shape dictionaries with text content
        """
        root = ET.Element(f"{{{self.NS['p']}}}sld")
        
        # Common slide data
        cSld = ET.SubElement(root, f"{{{self.NS['p']}}}cSld")
        sp_tree = ET.SubElement(cSld, f"{{{self.NS['p']}}}spTree")
        
        # Group shape properties
        nv_grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}nvGrpSpPr")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvPr", id="1", name="")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}cNvGrpSpPr")
        ET.SubElement(nv_grp_sp_pr, f"{{{self.NS['p']}}}nvPr")
        
        grp_sp_pr = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}grpSpPr")
        xfrm = ET.SubElement(grp_sp_pr, f"{{{self.NS['a']}}}xfrm")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}off", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}ext", cx="0", cy="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chOff", x="0", y="0")
        ET.SubElement(xfrm, f"{{{self.NS['a']}}}chExt", cx="0", cy="0")
        
        # Add shapes with content
        for idx, shape in enumerate(shapes):
            sp = ET.SubElement(sp_tree, f"{{{self.NS['p']}}}sp")
            
            # Non-visual shape properties
            nv_sp_pr = ET.SubElement(sp, f"{{{self.NS['p']}}}nvSpPr")
            ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}cNvPr", id=str(idx + 2), name=shape.get('name', f"Shape {idx+2}"))
            cnv_sp_pr = ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}cNvSpPr")
            ET.SubElement(cnv_sp_pr, f"{{{self.NS['a']}}}spLocks", noGrp="1")
            nv_pr = ET.SubElement(nv_sp_pr, f"{{{self.NS['p']}}}nvPr")
            ph_attrs = {'type': shape.get('type', 'body')}
            if idx > 0:
                ph_attrs['idx'] = str(idx)
            ET.SubElement(nv_pr, f"{{{self.NS['p']}}}ph", **ph_attrs)
            
            # Shape properties
            ET.SubElement(sp, f"{{{self.NS['p']}}}spPr")
            
            # Text body
            tx_body = ET.SubElement(sp, f"{{{self.NS['p']}}}txBody")
            ET.SubElement(tx_body, f"{{{self.NS['a']}}}bodyPr")
            ET.SubElement(tx_body, f"{{{self.NS['a']}}}lstStyle")
            
            # Add text content
            text_content = shape.get('text', '')
            if isinstance(text_content, str):
                text_content = [text_content]
            
            for text_item in text_content:
                p = ET.SubElement(tx_body, f"{{{self.NS['a']}}}p")
                r = ET.SubElement(p, f"{{{self.NS['a']}}}r")
                rPr = ET.SubElement(r, f"{{{self.NS['a']}}}rPr", lang="en-US", dirty="0")
                ET.SubElement(r, f"{{{self.NS['a']}}}t").text = text_item
                ET.SubElement(p, f"{{{self.NS['a']}}}endParaRPr", lang="en-US")
        
        # Color map override
        clr_map_ovr = ET.SubElement(root, f"{{{self.NS['p']}}}clrMapOvr")
        ET.SubElement(clr_map_ovr, f"{{{self.NS['a']}}}masterClrMapping")
        
        return self._xml_to_bytes(root)
    
    def _create_slide_rels(self, slide_num: int, layout_num: int) -> bytes:
        """Create ppt/slides/_rels/slide{n}.xml.rels with default namespace (no prefix)"""
        # Manually construct XML to avoid namespace prefix issues
        declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        xml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        xml += f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{layout_num}.xml"/>'
        xml += '</Relationships>'
        return (declaration + xml).encode('utf-8')
    
    def add_title_slide(self, title: str, subtitle: str = ""):
        """Add a title slide"""
        shapes = [
            {'type': 'ctrTitle', 'name': 'Title 1', 'text': title},
            {'type': 'subTitle', 'name': 'Subtitle 2', 'text': subtitle},
        ]
        self.slides.append({'layout': 1, 'shapes': shapes})
    
    def add_content_slide(self, title: str, content: List[str]):
        """Add a content slide with bullet points"""
        shapes = [
            {'type': 'title', 'name': 'Title 1', 'text': title},
            {'type': 'body', 'name': 'Content Placeholder 2', 'text': content},
        ]
        self.slides.append({'layout': 2, 'shapes': shapes})
    
    def add_two_column_slide(self, title: str, left_content: List[str], right_content: List[str]):
        """Add a two-column slide"""
        shapes = [
            {'type': 'title', 'name': 'Title 1', 'text': title},
            {'type': 'body', 'name': 'Content Placeholder 2', 'text': left_content},
            {'type': 'body', 'name': 'Content Placeholder 3', 'text': right_content},
        ]
        self.slides.append({'layout': 4, 'shapes': shapes})
    
    def add_section_header(self, title: str, subtitle: str = ""):
        """Add a section header slide"""
        shapes = [
            {'type': 'title', 'name': 'Title 1', 'text': title},
            {'type': 'body', 'name': 'Text Placeholder 2', 'text': subtitle},
        ]
        self.slides.append({'layout': 3, 'shapes': shapes})
    
    def add_blank_slide(self, title: str, text: str):
        """Add a blank slide with title"""
        shapes = [
            {'type': 'title', 'name': 'Title 1', 'text': title},
        ]
        self.slides.append({'layout': 6, 'shapes': shapes})
    
    def save(self, filename: str, verbose: bool = False):
        """Save the presentation to a .pptx file"""
        if verbose:
            print(f"Generating presentation with {len(self.slides)} slides...")
        
        with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Package structure files
            zipf.writestr('[Content_Types].xml', self._create_content_types())
            zipf.writestr('_rels/.rels', self._create_package_rels())
            
            # Document properties
            zipf.writestr('docProps/core.xml', self._create_core_props())
            zipf.writestr('docProps/app.xml', self._create_app_props())
            
            # Presentation files
            zipf.writestr('ppt/presentation.xml', self._create_presentation_xml())
            zipf.writestr('ppt/_rels/presentation.xml.rels', self._create_presentation_rels())
            zipf.writestr('ppt/presProps.xml', self._create_pres_props())
            zipf.writestr('ppt/viewProps.xml', self._create_view_props())
            zipf.writestr('ppt/tableStyles.xml', self._create_table_styles())
            
            # Theme
            zipf.writestr('ppt/theme/theme1.xml', self._create_theme())
            
            # Slide master
            zipf.writestr('ppt/slideMasters/slideMaster1.xml', self._create_slide_master())
            zipf.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', self._create_slide_master_rels())
            
            # Slide layouts (all 11)
            layouts = self._create_all_slide_layouts()
            for path, content in layouts.items():
                zipf.writestr(path, content)
            
            # Slides
            for i, slide_data in enumerate(self.slides, 1):
                zipf.writestr(f'ppt/slides/slide{i}.xml',
                             self._create_slide(i, slide_data['layout'], slide_data['shapes']))
                zipf.writestr(f'ppt/slides/_rels/slide{i}.xml.rels',
                             self._create_slide_rels(i, slide_data['layout']))
            
            if verbose:
                print(f"Saved {filename} successfully!")
                print(f"  - {len(self.slides)} slides")
                print(f"  - {len(zipf.namelist())} files in archive")


if __name__ == "__main__":
    # Test the generator
    ppt = CompletePPTXGenerator()
    ppt.add_title_slide("Test Presentation", "Generated from Scratch")
    ppt.add_content_slide("Key Points", [
        "Built entirely from scratch",
        "Uses only Python standard library",
        "Creates complete slide layouts",
        "Includes all required placeholders"
    ])
    ppt.add_two_column_slide("Comparison", 
                            ["From scratch", "Complete XML", "No dependencies"],
                            ["Full compatibility", "Proper layouts", "Works everywhere"])
    ppt.save('/tmp/test_complete.pptx', verbose=True)
    print("\nTest presentation created at /tmp/test_complete.pptx")
