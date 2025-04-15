import os
import pandas as pd
from docx import Document
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32
import json
from datetime import datetime

class EnhancedConverter:
    def __init__(self):
        self.ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        self.docx_analysis = None
        self.debug_info = {
            'docx_analysis': [],
            'xml_analysis': [],
            'format_matches': [],
            'conversion_issues': []
        }
        
    def analyze_docx(self, file_path):
        """Analyze DOCX file to extract formatting information"""
        doc = Document(file_path)
        content = []
        
        for para in doc.paragraphs:
            # Skip completely empty paragraphs
            if not para.text.strip():
                continue
                
            # Get indentation info
            indent_info = para.paragraph_format.left_indent
            left_indent = indent_info.pt if indent_info else 0
            
            first_line_indent = para.paragraph_format.first_line_indent
            first_line = first_line_indent.pt if first_line_indent else 0
            
            # Get alignment
            alignment = para.alignment
            alignment_map = {
                0: 'LEFT',
                1: 'CENTER',
                2: 'RIGHT',
                3: 'JUSTIFY'
            }
            alignment_str = alignment_map.get(alignment, 'LEFT')
            
            # Analyze runs for character formatting
            run_formats = []
            current_position = 0
            
            for run in para.runs:
                run_text = run.text
                run_format = {
                    'text': run_text,
                    'bold': run.bold if run.bold is not None else False,
                    'italic': run.italic if run.italic is not None else False,
                    'underline': run.underline if run.underline is not None else False,
                    'font_name': run.font.name,
                    'font_size': run.font.size.pt if run.font.size else None,
                    'position_start': current_position,
                    'position_end': current_position + len(run_text)
                }
                current_position += len(run_text)
                run_formats.append(run_format)
            
            # Check for numbering
            is_numbered = bool(para._p.find('.//w:numPr', self.ns))
            numbering_level = None
            if is_numbered:
                num_pr = para._p.find('.//w:numPr', self.ns)
                ilvl = num_pr.find('.//w:ilvl', self.ns)
                if ilvl is not None:
                    numbering_level = int(ilvl.get(f'{{{self.ns["w"]}}}val', 0))
            
            para_info = {
                'text': para.text,
                'left_indent': left_indent,
                'first_line_indent': first_line,
                'alignment': alignment_str,
                'run_formats': run_formats,
                'is_numbered': is_numbered,
                'numbering_level': numbering_level
            }
            
            content.append(para_info)
            self.debug_info['docx_analysis'].append(para_info)
        
        self.docx_analysis = pd.DataFrame(content)
        return self.docx_analysis
    
    def _get_run_properties(self, rPr):
        """Extract run properties from XML"""
        props = {}
        if rPr is None:
            return props
            
        # Font family
        rFonts = rPr.find('.//w:rFonts', self.ns)
        if rFonts is not None:
            for attr in ['w:ascii', 'w:hAnsi']:
                font = rFonts.get(f'{{{self.ns["w"]}}}{attr[2:]}')
                if font:
                    props['font-family'] = font
                    break
        
        # Bold
        bold = rPr.find('.//w:b', self.ns)
        if bold is not None and bold.get(f'{{{self.ns["w"]}}}val', 'true') != 'false':
            props['font-weight'] = 'bold'
        
        # Italic
        italic = rPr.find('.//w:i', self.ns)
        if italic is not None and italic.get(f'{{{self.ns["w"]}}}val', 'true') != 'false':
            props['font-style'] = 'italic'
        
        # Underline
        underline = rPr.find('.//w:u', self.ns)
        if underline is not None:
            u_val = underline.get(f'{{{self.ns["w"]}}}val')
            if u_val and u_val != 'none':
                props['text-decoration'] = 'underline'
        
        return props
    
    def find_matching_format(self, text):
        """Find matching format from analysis for given text"""
        if self.docx_analysis is None or not text.strip():
            return None
            
        # Find the most similar text in analysis
        best_match = None
        best_score = 0
        
        text_words = set(text.lower().split())
        
        for _, row in self.docx_analysis.iterrows():
            # Skip empty rows
            if not row['text'].strip():
                continue
                
            # Calculate similarity score based on common words and length difference
            row_words = set(row['text'].lower().split())
            common_words = text_words & row_words
            score = len(common_words)
            
            # Penalize large length differences
            len_diff = abs(len(text) - len(row['text']))
            score = score - (len_diff / 100)  # Small penalty for length differences
            
            if score > best_score:
                best_score = score
                best_match = row.to_dict()
                
        if best_match:
            self.debug_info['format_matches'].append({
                'text': text,
                'matched_text': best_match['text'],
                'score': best_score,
                'formats': best_match
            })
            
        return best_match if best_score > 0 else None
    
    def enhance_fo_block(self, block, paragraph, text):
        """Enhance FO block with formatting from analysis"""
        format_info = self.find_matching_format(text)
        if format_info is not None:
            # Skip indentation handling as requested
            
            # Apply alignment
            if format_info['alignment'] != 'LEFT':
                block.set('text-align', format_info['alignment'].lower())
            
            # Handle numbering - improved to detect and preserve number prefixes
            if format_info['is_numbered'] or any(text.strip().startswith(str(i)) for i in range(1, 100)):
                # Extract number prefix if it exists
                import re
                number_match = re.match(r'^(\d+\.?\d*)\s+(.+)$', text.strip())
                if number_match:
                    number, content = number_match.groups()
                    # Create a marker for the number
                    marker = ET.SubElement(block, 'fo:inline')
                    marker.text = number + " "
                    marker.set('font-weight', 'bold')
                    
                    # Add the remaining content
                    if len(block) > 0:
                        if block[-1].tail is None:
                            block[-1].tail = content
                        else:
                            block[-1].tail += content
                    else:
                        block.text = content
                    return True
                
            # Process character formatting
            text_with_formats = ''
            current_pos = 0
            
            # Get run properties from XML
            for run in paragraph.findall('.//w:r', self.ns):
                run_text = ''.join(t.text or '' for t in run.findall('.//w:t', self.ns))
                if not run_text:
                    continue
                    
                rPr = run.find('.//w:rPr', self.ns)
                props = self._get_run_properties(rPr)
                
                if props:
                    inline = ET.SubElement(block, 'fo:inline')
                    for prop, value in props.items():
                        inline.set(prop, value)
                    inline.text = run_text
                else:
                    if not block.text:
                        block.text = run_text
                    else:
                        if len(block) > 0:
                            if block[-1].tail is None:
                                block[-1].tail = run_text
                            else:
                                block[-1].tail += run_text
                        else:
                            block.text += run_text
            
            return True
        return False
    
    def convert(self, xml_file, docx_file, output_file):
        """Convert XML to XSL-FO with enhanced formatting"""
        # First analyze the DOCX file
        self.analyze_docx(docx_file)
        
        # Parse input XML
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        # Create XSL-FO document structure
        fo_root = ET.Element('fo:root', {
            'xmlns:fo': 'http://www.w3.org/1999/XSL/Format',
            'xmlns:xfd': 'http://www.ecrion.com/xfd/1.0',
            'xmlns:xf': 'http://www.ecrion.com/xf/1.0'
        })
        
        # Add layout master set
        layout_master_set = ET.SubElement(fo_root, 'fo:layout-master-set')
        page_master = ET.SubElement(layout_master_set, 'fo:simple-page-master', {
            'master-name': 'main',
            'page-height': '11in',
            'page-width': '8.5in',
            'margin-top': '0.5in',
            'margin-bottom': '0.5in',
            'margin-left': '1in',
            'margin-right': '1in'
        })
        
        region_body = ET.SubElement(page_master, 'fo:region-body')
        
        # Create page sequence
        page_sequence = ET.SubElement(fo_root, 'fo:page-sequence', {
            'master-reference': 'main'
        })
        
        flow = ET.SubElement(page_sequence, 'fo:flow', {
            'flow-name': 'xsl-region-body'
        })

        # Convert paragraphs with enhanced formatting
        for paragraph in root.findall('.//w:p', self.ns):
            block = ET.SubElement(flow, 'fo:block')
            text = ''.join(t.text or '' for t in paragraph.findall('.//w:t', self.ns))
            
            # Apply enhanced formatting
            success = self.enhance_fo_block(block, paragraph, text)
            if not success:
                self.debug_info['conversion_issues'].append({
                    'text': text,
                    'issue': 'Failed to find matching format'
                })
            
        # Write output
        tree = ET.ElementTree(fo_root)
        tree.write(output_file, encoding='utf-8', xml_declaration=True)
        
        # Write debug information to same directory as output file
        output_dir = os.path.dirname(output_file)
        debug_filename = f'conversion_debug_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
        debug_file = os.path.join(output_dir, debug_filename)
        
        with open(debug_file, 'w', encoding='utf-8') as f:
            json.dump(self.debug_info, f, indent=2, default=str)
        return debug_file

def main():
    root = tk.Tk()
    root.withdraw()
    
    # Select XML file
    xml_file = filedialog.askopenfilename(
        title="Select XML file",
        filetypes=[("XML files", "*.xml")]
    )
    
    if not xml_file:
        messagebox.showinfo("No file selected", "No XML file was selected. Exiting.")
        return
    
    # Select corresponding DOCX file
    docx_file = filedialog.askopenfilename(
        title="Select corresponding DOCX file",
        filetypes=[("Word Documents", "*.docx")]
    )
    
    if not docx_file:
        messagebox.showinfo("No file selected", "No DOCX file was selected. Exiting.")
        return
    
    # Select output file location
    output_file = filedialog.asksaveasfilename(
        title="Save XSL-FO file as",
        defaultextension=".fo",
        filetypes=[("XSL-FO files", "*.fo")]
    )
    
    if not output_file:
        messagebox.showinfo("No file selected", "No output location selected. Exiting.")
        return
    
    try:
        converter = EnhancedConverter()
        debug_file = converter.convert(xml_file, docx_file, output_file)
        messagebox.showinfo("Success", 
            f"Successfully converted files to:\n{output_file}\n\n"
            f"Debug information has been saved to:\n{debug_file}\n\n"
            f"Please check both files to verify the conversion.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == '__main__':
    main() 