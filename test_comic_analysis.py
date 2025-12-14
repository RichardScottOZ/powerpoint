"""
Tests for Comic Analysis PowerPoint Presentations

Validates that the Comic-Analysis repository presentations are generated correctly
and contain the expected content and structure.
"""

import os
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pptx_generator import PPTXGenerator
import comic_analysis_presenter


class TestComicAnalysisPresentations(unittest.TestCase):
    """Test suite for Comic-Analysis presentations."""
    
    @classmethod
    def setUpClass(cls):
        """Generate all presentations before running tests."""
        # Clean up any existing presentations
        for filename in ['comic_analysis_overview.pptx',
                        'comic_analysis_technical.pptx',
                        'comic_analysis_data_insights.pptx',
                        'comic_analysis_getting_started.pptx']:
            if os.path.exists(filename):
                os.remove(filename)
        
        # Generate presentations
        comic_analysis_presenter.create_overview_presentation()
        comic_analysis_presenter.create_technical_presentation()
        comic_analysis_presenter.create_data_insights_presentation()
        comic_analysis_presenter.create_getting_started_presentation()
    
    @classmethod
    def tearDownClass(cls):
        """Clean up generated presentations after tests."""
        for filename in ['comic_analysis_overview.pptx',
                        'comic_analysis_technical.pptx',
                        'comic_analysis_data_insights.pptx',
                        'comic_analysis_getting_started.pptx']:
            if os.path.exists(filename):
                os.remove(filename)
    
    def _validate_pptx_structure(self, filename):
        """Helper method to validate basic PPTX structure."""
        # Check file exists
        self.assertTrue(os.path.exists(filename), 
                       f"{filename} should exist")
        
        # Check it's a valid ZIP
        self.assertTrue(zipfile.is_zipfile(filename),
                       f"{filename} should be a valid ZIP file")
        
        with zipfile.ZipFile(filename, 'r') as zf:
            files = zf.namelist()
            
            # Check required files exist
            required_files = [
                '[Content_Types].xml',
                '_rels/.rels',
                'ppt/presentation.xml',
            ]
            
            for req_file in required_files:
                self.assertIn(req_file, files,
                            f"{filename} should contain {req_file}")
            
            # Validate XML files
            for file in files:
                if file.endswith('.xml'):
                    content = zf.read(file)
                    try:
                        ET.fromstring(content)
                    except ET.ParseError as e:
                        self.fail(f"{file} in {filename} has invalid XML: {e}")
    
    def _count_slides(self, filename):
        """Helper method to count slides in a presentation."""
        with zipfile.ZipFile(filename, 'r') as zf:
            files = zf.namelist()
            slide_count = sum(1 for f in files if f.startswith('ppt/slides/slide') and f.endswith('.xml'))
            return slide_count
    
    def _extract_slide_text(self, filename, slide_num):
        """Helper method to extract text content from a slide."""
        with zipfile.ZipFile(filename, 'r') as zf:
            slide_xml = zf.read(f'ppt/slides/slide{slide_num}.xml')
            root = ET.fromstring(slide_xml)
            
            # Find all text elements
            text_elements = []
            for elem in root.iter():
                if elem.text and elem.text.strip():
                    text_elements.append(elem.text.strip())
            
            return text_elements
    
    # Overview Presentation Tests
    
    def test_overview_presentation_exists(self):
        """Test that overview presentation is created."""
        self.assertTrue(os.path.exists('comic_analysis_overview.pptx'))
    
    def test_overview_presentation_structure(self):
        """Test that overview presentation has valid structure."""
        self._validate_pptx_structure('comic_analysis_overview.pptx')
    
    def test_overview_presentation_slide_count(self):
        """Test that overview presentation has expected number of slides."""
        slide_count = self._count_slides('comic_analysis_overview.pptx')
        self.assertEqual(slide_count, 5, 
                        "Overview presentation should have 5 slides")
    
    def test_overview_presentation_content(self):
        """Test that overview presentation contains expected content."""
        # Check title slide
        text = self._extract_slide_text('comic_analysis_overview.pptx', 1)
        self.assertTrue(any('Comic Analysis' in t for t in text),
                       "Title slide should contain 'Comic Analysis'")
        
        # Check for key topics
        all_text = []
        for i in range(1, 6):
            all_text.extend(self._extract_slide_text('comic_analysis_overview.pptx', i))
        
        key_topics = ['Project Overview', 'Key Components', 'Technology Stack', 'Use Cases']
        for topic in key_topics:
            self.assertTrue(any(topic in t for t in all_text),
                          f"Overview should mention '{topic}'")
    
    # Technical Presentation Tests
    
    def test_technical_presentation_exists(self):
        """Test that technical presentation is created."""
        self.assertTrue(os.path.exists('comic_analysis_technical.pptx'))
    
    def test_technical_presentation_structure(self):
        """Test that technical presentation has valid structure."""
        self._validate_pptx_structure('comic_analysis_technical.pptx')
    
    def test_technical_presentation_slide_count(self):
        """Test that technical presentation has expected number of slides."""
        slide_count = self._count_slides('comic_analysis_technical.pptx')
        self.assertEqual(slide_count, 6,
                        "Technical presentation should have 6 slides")
    
    def test_technical_presentation_content(self):
        """Test that technical presentation contains technical content."""
        all_text = []
        for i in range(1, 7):
            all_text.extend(self._extract_slide_text('comic_analysis_technical.pptx', i))
        
        technical_topics = ['Architecture', 'Pipeline', 'Computer Vision', 
                           'Machine Learning', 'Performance']
        for topic in technical_topics:
            self.assertTrue(any(topic in t for t in all_text),
                          f"Technical presentation should mention '{topic}'")
    
    # Data Insights Presentation Tests
    
    def test_data_insights_presentation_exists(self):
        """Test that data insights presentation is created."""
        self.assertTrue(os.path.exists('comic_analysis_data_insights.pptx'))
    
    def test_data_insights_presentation_structure(self):
        """Test that data insights presentation has valid structure."""
        self._validate_pptx_structure('comic_analysis_data_insights.pptx')
    
    def test_data_insights_presentation_slide_count(self):
        """Test that data insights presentation has expected number of slides."""
        slide_count = self._count_slides('comic_analysis_data_insights.pptx')
        self.assertEqual(slide_count, 5,
                        "Data insights presentation should have 5 slides")
    
    def test_data_insights_presentation_content(self):
        """Test that data insights presentation contains data analysis content."""
        all_text = []
        for i in range(1, 6):
            all_text.extend(self._extract_slide_text('comic_analysis_data_insights.pptx', i))
        
        data_topics = ['Data Insights', 'Types of Data', 'Quantitative Analysis', 
                      'Applications']
        for topic in data_topics:
            self.assertTrue(any(topic in t for t in all_text),
                          f"Data insights should mention '{topic}'")
    
    # Getting Started Presentation Tests
    
    def test_getting_started_presentation_exists(self):
        """Test that getting started presentation is created."""
        self.assertTrue(os.path.exists('comic_analysis_getting_started.pptx'))
    
    def test_getting_started_presentation_structure(self):
        """Test that getting started presentation has valid structure."""
        self._validate_pptx_structure('comic_analysis_getting_started.pptx')
    
    def test_getting_started_presentation_slide_count(self):
        """Test that getting started presentation has expected number of slides."""
        slide_count = self._count_slides('comic_analysis_getting_started.pptx')
        self.assertEqual(slide_count, 6,
                        "Getting started presentation should have 6 slides")
    
    def test_getting_started_presentation_content(self):
        """Test that getting started presentation contains onboarding content."""
        all_text = []
        for i in range(1, 7):
            all_text.extend(self._extract_slide_text('comic_analysis_getting_started.pptx', i))
        
        onboarding_topics = ['Getting Started', 'Prerequisites', 'Installation', 
                            'Quick Start']
        for topic in onboarding_topics:
            self.assertTrue(any(topic in t for t in all_text),
                          f"Getting started should mention '{topic}'")
    
    # Cross-cutting Tests
    
    def test_all_presentations_created(self):
        """Test that all four presentations are created."""
        expected_files = [
            'comic_analysis_overview.pptx',
            'comic_analysis_technical.pptx',
            'comic_analysis_data_insights.pptx',
            'comic_analysis_getting_started.pptx'
        ]
        
        for filename in expected_files:
            self.assertTrue(os.path.exists(filename),
                          f"{filename} should be created")
    
    def test_total_slide_count(self):
        """Test total number of slides across all presentations."""
        total_slides = 0
        for filename in ['comic_analysis_overview.pptx',
                        'comic_analysis_technical.pptx',
                        'comic_analysis_data_insights.pptx',
                        'comic_analysis_getting_started.pptx']:
            total_slides += self._count_slides(filename)
        
        self.assertEqual(total_slides, 22,
                        "Total slides across all presentations should be 22")
    
    def test_xml_namespace_correctness(self):
        """Test that XML files use correct namespaces."""
        with zipfile.ZipFile('comic_analysis_overview.pptx', 'r') as zf:
            presentation_xml = zf.read('ppt/presentation.xml')
            content = presentation_xml.decode('utf-8')
            
            # Check for PresentationML namespace in presentation.xml
            self.assertIn('http://schemas.openxmlformats.org/presentationml/2006/main',
                         content, "presentation.xml should have PresentationML namespace")
            
            # Check for DrawingML namespace in slide XML
            slide_xml = zf.read('ppt/slides/slide1.xml')
            slide_content = slide_xml.decode('utf-8')
            self.assertIn('http://schemas.openxmlformats.org/drawingml/2006/main',
                         slide_content, "slide XML should have DrawingML namespace")
    
    def test_presentations_openable(self):
        """Test that presentations can be opened and parsed."""
        for filename in ['comic_analysis_overview.pptx',
                        'comic_analysis_technical.pptx',
                        'comic_analysis_data_insights.pptx',
                        'comic_analysis_getting_started.pptx']:
            with zipfile.ZipFile(filename, 'r') as zf:
                # Try to read and parse main presentation file
                presentation_xml = zf.read('ppt/presentation.xml')
                root = ET.fromstring(presentation_xml)
                self.assertIsNotNone(root, 
                                   f"{filename} presentation.xml should be parseable")


class TestPPTXGenerator(unittest.TestCase):
    """Test suite for the core PPTX generator."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.test_file = 'test_output.pptx'
    
    def tearDown(self):
        """Clean up test files."""
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def test_generator_creates_valid_pptx(self):
        """Test that generator creates valid PPTX files."""
        ppt = PPTXGenerator()
        ppt.add_title_slide("Test Title", "Test Subtitle")
        ppt.save(self.test_file)
        
        self.assertTrue(os.path.exists(self.test_file))
        self.assertTrue(zipfile.is_zipfile(self.test_file))
    
    def test_generator_xml_validation(self):
        """Test that generated XML is valid."""
        ppt = PPTXGenerator()
        ppt.add_title_slide("Test", "Test")
        ppt.add_content_slide("Content", ["Point 1", "Point 2"])
        ppt.save(self.test_file)
        
        with zipfile.ZipFile(self.test_file, 'r') as zf:
            for filename in zf.namelist():
                if filename.endswith('.xml'):
                    content = zf.read(filename)
                    # Should not raise exception
                    ET.fromstring(content)
    
    def test_generator_slide_count(self):
        """Test that generator creates correct number of slides."""
        ppt = PPTXGenerator()
        ppt.add_title_slide("Title", "Subtitle")
        ppt.add_content_slide("Slide 1", ["Content"])
        ppt.add_content_slide("Slide 2", ["Content"])
        ppt.save(self.test_file)
        
        with zipfile.ZipFile(self.test_file, 'r') as zf:
            files = zf.namelist()
            slide_count = sum(1 for f in files if f.startswith('ppt/slides/slide') and f.endswith('.xml'))
            self.assertEqual(slide_count, 3)


def run_tests():
    """Run all tests and return results."""
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add all test cases
    suite.addTests(loader.loadTestsFromTestCase(TestComicAnalysisPresentations))
    suite.addTests(loader.loadTestsFromTestCase(TestPPTXGenerator))
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result


if __name__ == '__main__':
    result = run_tests()
    
    # Exit with appropriate code
    exit(0 if result.wasSuccessful() else 1)
