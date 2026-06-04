"""
Citation Rebuild Module
Detects broken citations in Word documents and rebuilds CWYW field codes
using metadata from EndNote libraries or prior documents.
"""

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass


@dataclass
class Citation:
    """Represents a broken citation in text."""
    author: str
    year: str
    recnum: str
    position: Tuple[int, int]  # (start, end) in document
    full_text: str


@dataclass
class ReferenceRecord:
    """Represents a reference from the library."""
    recnum: str
    author: str
    year: str
    title: str
    ref_type: str
    metadata_xml: str  # Full XML for embedding in field code


class CitationDetector:
    """Detects broken citations in Word documents."""
    
    CITATION_PATTERN = r'\{([^}#]+),\s*(\d{4})\s*#(\d+)\}'
    
    @staticmethod
    def detect_broken_citations(word_doc_path: Path) -> List[Citation]:
        """
        Detect all broken citations in a Word document.
        
        Returns list of Citation objects with positions in document XML.
        """
        # Unpack the Word document
        citations = []
        
        try:
            with zipfile.ZipFile(word_doc_path, 'r') as docx:
                xml_content = docx.read('word/document.xml').decode('utf-8')
        except Exception as e:
            raise ValueError(f"Failed to read Word document: {e}")
        
        # Find all broken citations
        for match in re.finditer(CitationDetector.CITATION_PATTERN, xml_content):
            author, year, recnum = match.groups()
            citation = Citation(
                author=author.strip(),
                year=year,
                recnum=recnum,
                position=(match.start(), match.end()),
                full_text=match.group(0)
            )
            citations.append(citation)
        
        return citations


class LibraryExtractor:
    """Extracts reference metadata from EndNote .enlx library files."""
    
    @staticmethod
    def extract_from_enlx(enlx_path: Path, target_recnums: List[str]) -> Dict[str, ReferenceRecord]:
        """
        Extract reference records from .enlx file.
        
        Returns dict mapping RecNum -> ReferenceRecord
        """
        records = {}
        
        try:
            # .enlx is a ZIP file
            with zipfile.ZipFile(enlx_path, 'r') as lib_zip:
                # Try to find XML or library data
                file_list = lib_zip.namelist()
                
                # Look for XML export or data files
                xml_files = [f for f in file_list if f.endswith('.xml')]
                
                if xml_files:
                    # If there's an XML file, try to parse it
                    xml_data = lib_zip.read(xml_files[0]).decode('utf-8')
                    records = LibraryExtractor._parse_library_xml(xml_data, target_recnums)
        except Exception as e:
            raise ValueError(f"Failed to extract from .enlx file: {e}")
        
        if not records:
            raise ValueError("Could not extract references from library file")
        
        return records
    
    @staticmethod
    def _parse_library_xml(xml_text: str, target_recnums: List[str]) -> Dict[str, ReferenceRecord]:
        """Parse EndNote XML export format."""
        records = {}
        
        # Split by <record> tags
        record_blocks = re.split(r'(?=<record>)', xml_text)
        
        for block in record_blocks:
            if '<rec-number>' not in block:
                continue
            
            # Extract RecNum
            rec_match = re.search(r'<rec-number>(\d+)</rec-number>', block)
            if not rec_match:
                continue
            
            rec_num = rec_match.group(1)
            if rec_num not in target_recnums:
                continue
            
            # Extract author
            auth_match = re.search(r'<author[^>]*>(.*?)</author>', block)
            author = auth_match.group(1) if auth_match else "Unknown"
            author = re.sub(r'<[^>]+>', '', author).strip()
            
            # Extract year
            year_match = re.search(r'<year>(.*?)</year>', block, re.DOTALL)
            year = year_match.group(1) if year_match else "0000"
            year = re.sub(r'<[^>]+>', '', year).strip()
            
            # Extract title
            title_match = re.search(r'<title[^>]*>(.*?)</title>', block)
            title = title_match.group(1) if title_match else "No title"
            title = re.sub(r'<[^>]+>', '', title).strip()
            
            # Extract ref type
            ref_type_match = re.search(r'<ref-type[^>]*>(\d+)</ref-type>', block)
            ref_type = ref_type_match.group(1) if ref_type_match else "17"
            
            # Store full block for metadata embedding
            metadata_xml = block
            
            records[rec_num] = ReferenceRecord(
                recnum=rec_num,
                author=author,
                year=year,
                title=title,
                ref_type=ref_type,
                metadata_xml=metadata_xml
            )
        
        return records


class FieldCodeBuilder:
    """Builds proper CWYW field codes from reference metadata."""
    
    DB_ID = "assewx0rnp2et8efwrpv9een9wreav5x9w99"  # Default Tachdjian's library DB ID
    
    @staticmethod
    def build_field_code_xml(recnum: str, metadata: ReferenceRecord, display_number: str) -> str:
        """
        Build complete field code XML structure for a citation.
        
        Structure:
        <w:r> (begin marker)
        <w:r> (instrText with metadata)
        <w:r> (separate marker)
        <w:r> (display number - superscript)
        <w:r> (end marker)
        """
        
        # Create minimal ADDIN EN.CITE structure
        instr_content = FieldCodeBuilder._create_addin_content(recnum, metadata)
        
        # Escape for XML
        escaped_instr = instr_content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        # Build complete field code XML
        field_code = f"""<w:r>
        <w:rPr>
          <w:color w:val="000000"/>
        </w:rPr>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:color w:val="000000"/>
        </w:rPr>
        <w:instrText xml:space="preserve">{escaped_instr}</w:instrText>
      </w:r>
      <w:r>
        <w:rPr>
          <w:color w:val="000000"/>
        </w:rPr>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:rPr>
          <w:noProof/>
          <w:color w:val="000000"/>
          <w:vertAlign w:val="superscript"/>
        </w:rPr>
        <w:t>{display_number}</w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:color w:val="000000"/>
        </w:rPr>
        <w:fldChar w:fldCharType="end"/>
      </w:r>"""
        
        return field_code
    
    @staticmethod
    def _create_addin_content(recnum: str, metadata: ReferenceRecord) -> str:
        """Create ADDIN EN.CITE content string."""
        
        # Minimal ADDIN structure with metadata
        addin = f""" ADDIN EN.CITE <EndNote><Cite><Author>{metadata.author}</Author><Year>{metadata.year}</Year><RecNum>{recnum}</RecNum><DisplayText><style face="superscript">{recnum}</style></DisplayText><record><rec-number>{recnum}</rec-number><foreign-keys><key app="EN" db-id="{FieldCodeBuilder.DB_ID}" timestamp="0">{recnum}</key></foreign-keys><ref-type name="Journal Article">{metadata.ref_type}</ref-type><contributors><authors><author>{metadata.author}</author></authors></contributors><titles><title>{metadata.title}</title></titles><dates><year>{metadata.year}</year></dates><urls></urls></record></Cite></EndNote>"""
        
        return addin


class CitationRebuilder:
    """Main orchestrator for citation rebuild workflow."""
    
    @staticmethod
    def rebuild_document(
        broken_doc_path: Path,
        library_path: Optional[Path] = None,
        prior_doc_path: Optional[Path] = None,
        output_path: Optional[Path] = None
    ) -> Tuple[Path, Dict]:
        """
        Rebuild broken citations in a Word document.
        
        Args:
            broken_doc_path: Path to document with broken citations
            library_path: Path to .enlx library (optional if prior_doc provided)
            prior_doc_path: Path to prior working document (optional if library provided)
            output_path: Where to save fixed document (default: broken_doc with _FIXED suffix)
        
        Returns:
            (output_path, report_dict) with statistics and results
        """
        
        report = {
            'citations_found': 0,
            'citations_rebuilt': 0,
            'recnums_processed': {},
            'errors': []
        }
        
        try:
            # Step 1: Detect broken citations
            citations = CitationDetector.detect_broken_citations(broken_doc_path)
            report['citations_found'] = len(citations)
            
            if not citations:
                report['errors'].append("No broken citations detected in document")
                return broken_doc_path, report
            
            # Step 2: Extract metadata
            target_recnums = list(set(c.recnum for c in citations))
            
            if library_path:
                metadata = LibraryExtractor.extract_from_enlx(library_path, target_recnums)
            elif prior_doc_path:
                metadata = CitationRebuilder._extract_from_prior_doc(prior_doc_path, target_recnums)
            else:
                raise ValueError("Must provide either library_path or prior_doc_path")
            
            # Step 3: Read the broken document
            with zipfile.ZipFile(broken_doc_path, 'r') as docx:
                doc_xml = docx.read('word/document.xml').decode('utf-8')
            
            # Step 4: Replace citations (in reverse order to preserve positions)
            sorted_citations = sorted(citations, key=lambda c: c.position[0], reverse=True)
            
            for citation in sorted_citations:
                if citation.recnum not in metadata:
                    report['errors'].append(f"RecNum {citation.recnum} not found in metadata")
                    continue
                
                ref = metadata[citation.recnum]
                field_code_xml = FieldCodeBuilder.build_field_code_xml(
                    citation.recnum,
                    ref,
                    citation.recnum  # Use RecNum as display number for now
                )
                
                # Replace in document
                start, end = citation.position
                doc_xml = doc_xml[:start] + field_code_xml + doc_xml[end:]
                
                report['citations_rebuilt'] += 1
                report['recnums_processed'][citation.recnum] = ref.author
            
            # Step 5: Write output document
            if output_path is None:
                output_path = broken_doc_path.parent / f"{broken_doc_path.stem}_FIXED.docx"
            
            with zipfile.ZipFile(broken_doc_path, 'r') as docx_in:
                with zipfile.ZipFile(output_path, 'w') as docx_out:
                    for item in docx_in.infolist():
                        if item.filename == 'word/document.xml':
                            docx_out.writestr(item, doc_xml.encode('utf-8'))
                        else:
                            docx_out.writestr(item, docx_in.read(item.filename))
            
            return output_path, report
            
        except Exception as e:
            report['errors'].append(str(e))
            return None, report
    
    @staticmethod
    def _extract_from_prior_doc(prior_doc_path: Path, target_recnums: List[str]) -> Dict[str, ReferenceRecord]:
        """Extract reference metadata from prior working document's field codes."""
        # TODO: Implement extraction from prior doc's CWYW field codes
        raise NotImplementedError("Prior document extraction coming soon")


