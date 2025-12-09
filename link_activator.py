"""
link_activator.py

Post-processing module that converts plain text URLs in Word documents
into clickable hyperlinks.

Processes document.xml, endnotes.xml, and footnotes.xml to find URLs
and convert them to proper Word hyperlinks with blue underlined styling.

FIXED: 2025-12-09 - Replaced regex-based string replacement with proper
ElementTree XML parsing to prevent malformed XML that caused Word to
display "unreadable content" repair dialogs.

Now uses relationship-based hyperlinks (<w:hyperlink r:id="...">) which
is Word's native format, and properly manages .rels files.
"""

import os
import re
import html
import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET
from io import BytesIO
from typing import Dict, Optional


class RelsManager:
    """
    Manages the .rels file for a Word document part.
    
    Handles reading, modifying, and writing relationship files
    that map rIds to external hyperlinks.
    """
    
    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    HYPERLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    
    def __init__(self, rels_path: str):
        """
        Initialize the RelsManager.
        
        Args:
            rels_path: Path to the .rels file
        """
        self.rels_path = rels_path
        self.relationships: Dict[str, dict] = {}  # rId -> {type, target, mode}
        self.next_id = 1
        self.url_to_rid: Dict[str, str] = {}  # URL -> existing rId
        
        self._load()
    
    def _load(self):
        """Load existing relationships from file."""
        if not os.path.exists(self.rels_path):
            # Create directory if needed
            os.makedirs(os.path.dirname(self.rels_path), exist_ok=True)
            return
        
        try:
            ET.register_namespace('', self.RELS_NS)
            tree = ET.parse(self.rels_path)
            root = tree.getroot()
            
            for rel in root.findall(f'{{{self.RELS_NS}}}Relationship'):
                r_id = rel.get('Id', '')
                rel_type = rel.get('Type', '')
                target = rel.get('Target', '')
                target_mode = rel.get('TargetMode', '')
                
                self.relationships[r_id] = {
                    'type': rel_type,
                    'target': target,
                    'mode': target_mode
                }
                
                # Track existing hyperlink URLs
                if rel_type == self.HYPERLINK_TYPE:
                    self.url_to_rid[target] = r_id
                
                # Track highest rId number
                if r_id.startswith('rId'):
                    try:
                        num = int(r_id[3:])
                        if num >= self.next_id:
                            self.next_id = num + 1
                    except ValueError:
                        pass
                        
        except ET.ParseError as e:
            print(f"[RelsManager] Error parsing {self.rels_path}: {e}")
    
    def add_hyperlink(self, url: str) -> str:
        """
        Add a hyperlink relationship and return its rId.
        
        If the URL already exists, returns the existing rId.
        
        Args:
            url: The URL to link to
            
        Returns:
            The rId for this hyperlink
        """
        # Check if URL already has a relationship
        if url in self.url_to_rid:
            return self.url_to_rid[url]
        
        # Create new relationship
        r_id = f'rId{self.next_id}'
        self.next_id += 1
        
        self.relationships[r_id] = {
            'type': self.HYPERLINK_TYPE,
            'target': url,
            'mode': 'External'
        }
        
        self.url_to_rid[url] = r_id
        
        return r_id
    
    def save(self):
        """Save relationships to file."""
        ET.register_namespace('', self.RELS_NS)
        
        root = ET.Element(f'{{{self.RELS_NS}}}Relationships')
        
        # Sort by rId for consistent output
        for r_id in sorted(self.relationships.keys(), key=lambda x: (len(x), x)):
            rel_data = self.relationships[r_id]
            
            rel = ET.SubElement(root, f'{{{self.RELS_NS}}}Relationship')
            rel.set('Id', r_id)
            rel.set('Type', rel_data['type'])
            rel.set('Target', rel_data['target'])
            if rel_data.get('mode'):
                rel.set('TargetMode', rel_data['mode'])
        
        # Write to file
        tree = ET.ElementTree(root)
        tree.write(self.rels_path, encoding='UTF-8', xml_declaration=True)


class LinkActivator:
    """
    Post-processing module that converts plain text URLs in Word documents
    into clickable hyperlinks using relationship-based approach.
    
    This is the native Word hyperlink format that:
    - Creates <w:hyperlink r:id="rIdX"> elements
    - Manages word/_rels/*.xml.rels files
    - Produces valid XML that Word opens without repair
    """
    
    # Namespaces used in Word documents
    NS = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    
    # URL pattern
    URL_PATTERN = re.compile(r'https?://[^\s<>"\']+')
    
    @classmethod
    def process(cls, docx_bytes: bytes) -> bytes:
        """
        Process a .docx file to make all URLs clickable.
        
        Args:
            docx_bytes: Raw bytes of the input .docx file
            
        Returns:
            Bytes of the processed .docx file with clickable URLs
        """
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Extract docx to temp directory
            with zipfile.ZipFile(BytesIO(docx_bytes), 'r') as zf:
                zf.extractall(temp_dir)
            
            # Process each relevant XML file with its corresponding .rels file
            target_files = [
                ('word/document.xml', 'word/_rels/document.xml.rels'),
                ('word/endnotes.xml', 'word/_rels/endnotes.xml.rels'),
                ('word/footnotes.xml', 'word/_rels/footnotes.xml.rels'),
            ]
            
            for xml_file, rels_file in target_files:
                xml_path = os.path.join(temp_dir, xml_file)
                rels_path = os.path.join(temp_dir, rels_file)
                
                if os.path.exists(xml_path):
                    cls._process_xml_file(xml_path, rels_path)
            
            # Repackage as docx
            buffer = BytesIO()
            with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zf.write(file_path, arcname)
            
            buffer.seek(0)
            return buffer.read()
            
        except Exception as e:
            print(f"[LinkActivator] Error: {e}")
            import traceback
            traceback.print_exc()
            return docx_bytes  # Return original on error
            
        finally:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
    
    @classmethod
    def _process_xml_file(cls, xml_path: str, rels_path: str):
        """
        Process a single XML file to convert URLs to hyperlinks.
        
        Args:
            xml_path: Path to the XML file (document.xml, endnotes.xml, etc.)
            rels_path: Path to the corresponding .rels file
        """
        # Register namespaces to preserve them in output
        for prefix, uri in cls.NS.items():
            ET.register_namespace(prefix, uri)
        
        # Also register common namespaces found in Word docs
        ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
        ET.register_namespace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
        ET.register_namespace('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
        ET.register_namespace('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')
        ET.register_namespace('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas')
        ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
        ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        
        # Parse the XML file
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Load or create relationships
        rels_manager = RelsManager(rels_path)
        
        # Track URLs we've already processed to avoid duplicates
        processed_urls: Dict[str, str] = {}  # url -> rId
        
        # Find all text elements
        w_ns = cls.NS['w']
        
        # Process all paragraphs
        for para in root.iter(f'{{{w_ns}}}p'):
            cls._process_paragraph(para, rels_manager, processed_urls)
        
        # Save the modified XML
        tree.write(xml_path, encoding='UTF-8', xml_declaration=True)
        
        # Save the relationships file
        rels_manager.save()
    
    @classmethod
    def _process_paragraph(cls, para: ET.Element, rels_manager: 'RelsManager', 
                          processed_urls: Dict[str, str]):
        """
        Process a paragraph to convert URLs to hyperlinks.
        
        Args:
            para: The paragraph element
            rels_manager: Manager for the .rels file
            processed_urls: Cache of URL -> rId mappings
        """
        w_ns = cls.NS['w']
        r_ns = cls.NS['r']
        
        # Get list of direct children - we'll iterate through them
        # We need to work on a copy because we'll be modifying the paragraph
        children = list(para)
        
        for child in children:
            # Only process runs that are direct children (not inside hyperlinks)
            if child.tag != f'{{{w_ns}}}r':
                continue
            
            # Skip if this run is inside a hyperlink
            if cls._is_inside_hyperlink(child, para):
                continue
            
            # Find text element
            t_elem = child.find(f'{{{w_ns}}}t')
            if t_elem is None or not t_elem.text:
                continue
            
            text = t_elem.text
            
            # Find URLs in the text
            matches = list(cls.URL_PATTERN.finditer(text))
            if not matches:
                continue
            
            # Get run properties to preserve formatting
            rPr = child.find(f'{{{w_ns}}}rPr')
            rPr_copy = None
            if rPr is not None:
                rPr_copy = cls._copy_element(rPr)
            
            # Find position of this run in paragraph
            try:
                run_index = list(para).index(child)
            except ValueError:
                continue
            
            # Build new elements to replace this run
            new_elements = []
            last_end = 0
            
            for match in matches:
                url = match.group(0)
                start, end = match.start(), match.end()
                
                # Clean URL (remove trailing punctuation)
                clean_url = url.rstrip('.,;:)]\'"')
                trailing_punct = url[len(clean_url):]
                
                # Text before this URL
                if start > last_end:
                    text_before = text[last_end:start]
                    before_run = cls._create_run(text_before, rPr_copy, w_ns)
                    new_elements.append(before_run)
                
                # Get or create relationship ID for this URL
                if clean_url in processed_urls:
                    r_id = processed_urls[clean_url]
                else:
                    r_id = rels_manager.add_hyperlink(clean_url)
                    processed_urls[clean_url] = r_id
                
                # Create hyperlink element
                hyperlink = ET.Element(f'{{{w_ns}}}hyperlink')
                hyperlink.set(f'{{{r_ns}}}id', r_id)
                hyperlink.set(f'{{{w_ns}}}history', '1')
                
                # Create run inside hyperlink with URL text
                link_run = cls._create_hyperlink_run(clean_url, rPr_copy, w_ns)
                hyperlink.append(link_run)
                new_elements.append(hyperlink)
                
                # Include trailing punctuation in the "after" text
                last_end = end - len(trailing_punct)
            
            # Text after last URL
            if last_end < len(text):
                text_after = text[last_end:]
                after_run = cls._create_run(text_after, rPr_copy, w_ns)
                new_elements.append(after_run)
            
            # Replace original run with new elements
            para.remove(child)
            for i, elem in enumerate(new_elements):
                para.insert(run_index + i, elem)
    
    @classmethod
    def _create_run(cls, text: str, rPr_template: Optional[ET.Element], w_ns: str) -> ET.Element:
        """Create a run element with text."""
        run = ET.Element(f'{{{w_ns}}}r')
        
        if rPr_template is not None:
            run.append(cls._copy_element(rPr_template))
        
        t = ET.SubElement(run, f'{{{w_ns}}}t')
        t.text = text
        # Preserve whitespace
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        
        return run
    
    @classmethod
    def _create_hyperlink_run(cls, url: str, rPr_template: Optional[ET.Element], w_ns: str) -> ET.Element:
        """Create a run element for inside a hyperlink (with blue/underline styling)."""
        run = ET.Element(f'{{{w_ns}}}r')
        
        # Create run properties
        rPr = ET.SubElement(run, f'{{{w_ns}}}rPr')
        
        # Copy existing properties if present
        if rPr_template is not None:
            for child in rPr_template:
                # Skip color and underline - we'll add our own
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag not in ('color', 'u'):
                    rPr.append(cls._copy_element(child))
        
        # Add hyperlink styling (blue, underlined)
        color = ET.SubElement(rPr, f'{{{w_ns}}}color')
        color.set(f'{{{w_ns}}}val', '0000FF')
        
        underline = ET.SubElement(rPr, f'{{{w_ns}}}u')
        underline.set(f'{{{w_ns}}}val', 'single')
        
        # Add the URL text
        t = ET.SubElement(run, f'{{{w_ns}}}t')
        t.text = url
        
        return run
    
    @classmethod
    def _is_inside_hyperlink(cls, run: ET.Element, para: ET.Element) -> bool:
        """Check if a run is inside an existing hyperlink element."""
        w_ns = cls.NS['w']
        
        # Check if any hyperlink in paragraph contains this run
        for elem in para:
            if elem.tag == f'{{{w_ns}}}hyperlink':
                for child in elem.iter():
                    if child is run:
                        return True
        
        return False
    
    @classmethod
    def _copy_element(cls, elem: ET.Element) -> ET.Element:
        """Deep copy an element."""
        new_elem = ET.Element(elem.tag, elem.attrib)
        new_elem.text = elem.text
        new_elem.tail = elem.tail
        for child in elem:
            new_elem.append(cls._copy_element(child))
        return new_elem


def activate_links(docx_bytes: bytes) -> bytes:
    """
    Convenience function to activate links in a document.
    
    Args:
        docx_bytes: Raw bytes of the input .docx file
        
    Returns:
        Bytes of the processed .docx file with clickable URLs
    """
    return LinkActivator.process(docx_bytes)
