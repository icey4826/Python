import xml.etree.ElementTree as ET
import re
import argparse
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document

def select_files():
    """Show file selection dialogs and return input/output paths"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Select input file
    input_file = filedialog.askopenfilename(
        title="Select Word XML file",
        filetypes=[
            ("XML files", "*.xml"),
            ("All files", "*.*")
        ]
    )
    
    if not input_file:
        return None, None
        
    # Select output file
    output_file = filedialog.asksaveasfilename(
        title="Save XSL-FO file as",
        defaultextension=".fo",
        filetypes=[
            ("XSL-FO files", "*.fo"),
            ("All files", "*.*")
        ],
        initialfile="output.fo"
    )
    
    return input_file, output_file

class WordToFOConverter:
    def __init__(self):
        self.ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'fo': 'http://www.w3.org/1999/XSL/Format'
        }
        self.numbering_defs = {}  # Store numbering definitions
        self.list_counters = {}  # Track list numbering
        self.current_format = {}
        self.debug_log = []  # Store debug information

    def _log_debug(self, category, details):
        """Add debug information to the log"""
        self.debug_log.append({
            'category': category,
            'details': details
        })

    def _dump_debug_log(self, output_file):
        """Write debug log to a file"""
        import json
        debug_file = output_file + '.debug.json'
        with open(debug_file, 'w', encoding='utf-8') as f:
            json.dump(self.debug_log, f, indent=2, ensure_ascii=False)

    def _analyze_paragraph(self, paragraph, is_docx=False):
        """Analyze paragraph structure and properties"""
        analysis = {
            'type': 'docx' if is_docx else 'xml',
            'properties': {},
            'runs': [],
            'text': '',
            'list_info': None,
            'indentation': {},
            'alignment': None
        }

        # Get paragraph properties
        if is_docx:
            pPr = paragraph._p.find('.//w:pPr', self.ns)
        else:
            pPr = paragraph.find('.//w:pPr', self.ns)

        if pPr is not None:
            # Analyze numbering
            numPr = pPr.find('.//w:numPr', self.ns)
            if numPr is not None:
                ilvl = numPr.find('.//w:ilvl', self.ns)
                numId = numPr.find('.//w:numId', self.ns)
                if ilvl is not None and numId is not None:
                    analysis['list_info'] = {
                        'level': ilvl.get(f'{{{self.ns["w"]}}}val'),
                        'num_id': numId.get(f'{{{self.ns["w"]}}}val')
                    }

            # Analyze indentation
            ind = pPr.find('.//w:ind', self.ns)
            if ind is not None:
                for attr in ['left', 'right', 'firstLine', 'hanging']:
                    val = ind.get(f'{{{self.ns["w"]}}}{attr}')
                    if val:
                        analysis['indentation'][attr] = float(int(val)/1440)  # Convert twips to inches

            # Analyze alignment
            jc = pPr.find('.//w:jc', self.ns)
            if jc is not None:
                analysis['alignment'] = jc.get(f'{{{self.ns["w"]}}}val')

        # Analyze runs
        runs = paragraph._p.findall('.//w:r', self.ns) if is_docx else paragraph.findall('.//w:r', self.ns)
        for run in runs:
            run_info = {
                'text': '',
                'formatting': {}
            }

            # Get run properties
            rPr = run.find('.//w:rPr', self.ns)
            if rPr is not None:
                run_info['formatting'] = self._get_run_properties(rPr)

            # Get text
            text_elements = run.findall('.//w:t', self.ns)
            run_info['text'] = ''.join(elem.text or '' for elem in text_elements)
            analysis['text'] += run_info['text']
            analysis['runs'].append(run_info)

        return analysis

    def _parse_numbering_definitions(self, doc):
        """Parse numbering definitions from the Word document"""
        numbering_part = None
        for rel in doc.part.rels.values():
            if rel.reltype == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering':
                numbering_part = rel.target_part
                break
        
        if not numbering_part:
            return
            
        root = ET.fromstring(numbering_part.blob)
        
        # Parse abstract numbering definitions
        abstract_nums = {}
        for abstract_num in root.findall('.//w:abstractNum', self.ns):
            num_id = abstract_num.get(f'{{{self.ns["w"]}}}abstractNumId')
            levels = {}
            
            for level in abstract_num.findall('.//w:lvl', self.ns):
                ilvl = level.get(f'{{{self.ns["w"]}}}ilvl')
                num_fmt = level.find('.//w:numFmt', self.ns)
                format_type = num_fmt.get(f'{{{self.ns["w"]}}}val') if num_fmt is not None else 'decimal'
                
                # Get level text (e.g., "%1.", "%1.%2.")
                lvl_text = level.find('.//w:lvlText', self.ns)
                text_format = lvl_text.get(f'{{{self.ns["w"]}}}val') if lvl_text is not None else "%1."
                
                # Get indentation
                ind = level.find('.//w:ind', self.ns)
                left = ind.get(f'{{{self.ns["w"]}}}left') if ind is not None else None
                hanging = ind.get(f'{{{self.ns["w"]}}}hanging') if ind is not None else None
                
                levels[ilvl] = {
                    'format': format_type,
                    'text': text_format,
                    'left': left,
                    'hanging': hanging
                }
            
            abstract_nums[num_id] = levels
        
        # Link concrete numbering instances to abstract definitions
        for num in root.findall('.//w:num', self.ns):
            num_id = num.get(f'{{{self.ns["w"]}}}numId')
            abstract_num_id = num.find('.//w:abstractNumId', self.ns).get(f'{{{self.ns["w"]}}}val')
            self.numbering_defs[num_id] = abstract_nums[abstract_num_id]

    def _get_list_properties(self, pPr):
        """Get list properties from paragraph properties"""
        num_pr = pPr.find('.//w:numPr', self.ns)
        if num_pr is None:
            return None
            
        ilvl_elem = num_pr.find('.//w:ilvl', self.ns)
        num_id_elem = num_pr.find('.//w:numId', self.ns)
        
        if ilvl_elem is None or num_id_elem is None:
            return None
            
        ilvl = ilvl_elem.get(f'{{{self.ns["w"]}}}val')
        num_id = num_id_elem.get(f'{{{self.ns["w"]}}}val')
        
        if num_id not in self.numbering_defs:
            return None
            
        level_def = self.numbering_defs[num_id].get(ilvl)
        if level_def is None:
            return None
            
        return {
            'level': ilvl,
            'num_id': num_id,
            'format': level_def['format'],
            'text': level_def['text'],
            'left': level_def['left'],
            'hanging': level_def['hanging']
        }

    def _get_font_family(self, rFonts):
        """Extract font family from rFonts element"""
        if rFonts is None:
            return None
            
        # Try different font attributes in priority order
        for attr in ['w:ascii', 'w:hAnsi']:
            font = rFonts.get(f'{{{self.ns["w"]}}}{attr[2:]}')
            if font:
                return font
                
        return None

    def _get_color(self, color_elem):
        """Extract color value from color element"""
        if color_elem is None:
            return None
            
        val = color_elem.get(f'{{{self.ns["w"]}}}val')
        if val == 'auto' or val == '000000':
            return None
        
        # Handle theme colors
        theme_color = color_elem.get(f'{{{self.ns["w"]}}}themeColor')
        if theme_color:
            # Map theme colors to actual colors
            theme_colors = {
                'accent1': '#4472C4',
                'accent2': '#ED7D31',
                'accent3': '#A5A5A5',
                'accent4': '#FFC000',
                'accent5': '#5B9BD5',
                'accent6': '#70AD47'
            }
            return theme_colors.get(theme_color)
            
        return f'#{val}' if val else None

    def _get_run_properties(self, rPr):
        """Get formatting properties without creating element"""
        props = {}
        
        if rPr is None:
            return props
            
        # Font family
        rFonts = rPr.find('.//w:rFonts', self.ns)
        if rFonts is not None:
            font = self._get_font_family(rFonts)
            if font:
                props['font-family'] = font
            
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
            if u_val == 'single':
                props['text-underline-style'] = 'solid'
            elif u_val == 'double':
                props['text-underline-style'] = 'double'
            elif u_val == 'wave':
                props['text-underline-style'] = 'wave'
                
        # Strike through
        strike = rPr.find('.//w:strike', self.ns)
        if strike is not None:
            props['text-decoration'] = 'line-through'
            
        # Font size
        sz = rPr.find('.//w:sz', self.ns)
        if sz is not None:
            size = sz.get(f'{{{self.ns["w"]}}}val')
            if size:
                props['font-size'] = f'{int(size)/2}pt'
                
        # Text color
        color = rPr.find('.//w:color', self.ns)
        text_color = self._get_color(color)
        if text_color:
            props['color'] = text_color
                
        return props

    def _format_list_number(self, format_type, number):
        """Format a list number according to the specified format"""
        if format_type == 'decimal':
            return str(number)
        elif format_type == 'upperRoman':
            return self._to_roman(number).upper()
        elif format_type == 'lowerRoman':
            return self._to_roman(number).lower()
        elif format_type == 'upperLetter':
            if 0 < number <= 26:
                return chr(64 + number)
            return f'({number})'  # Fallback for numbers > 26
        elif format_type == 'lowerLetter':
            if 0 < number <= 26:
                return chr(96 + number)
            return f'({number})'  # Fallback for numbers > 26
        elif format_type == 'bullet':
            return '•'
        elif format_type == 'none':
            return ''
        return str(number)

    def _to_roman(self, num):
        """Convert number to Roman numerals"""
        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syb[i]
                num -= val[i]
            i += 1
        return roman_num

    def _format_list_label(self, list_props, numbers):
        """Format the list label according to the level text format"""
        text = list_props['text']
        format_type = list_props['format']
        
        # Handle special case for bullet lists
        if format_type == 'bullet':
            return '•'
            
        # Handle special case for no numbering
        if format_type == 'none':
            return ''
            
        # For letter formats, convert number to letter
        if format_type in ['lowerLetter', 'upperLetter']:
            if 0 < numbers[0] <= 26:
                letter = chr(96 + numbers[0]) if format_type == 'lowerLetter' else chr(64 + numbers[0])
                return f"{letter}."
            return f"({numbers[0]})"  # Fallback for numbers > 26
        
        # Format each number according to the format type
        formatted_numbers = [self._format_list_number(format_type, num) for num in numbers]
        
        # Replace placeholders with formatted numbers
        result = text
        for i, num in enumerate(formatted_numbers, 1):
            result = result.replace(f'%{i}', str(num))
            
        # If there are any remaining %n placeholders, replace them with empty string
        result = re.sub(r'%\d+', '', result)
        
        return result.rstrip()

    def _apply_paragraph_properties(self, block, pPr):
        """Apply paragraph formatting properties"""
        if pPr is None:
            return
            
        # Check for list properties first
        list_props = self._get_list_properties(pPr)
        if list_props:
            # Set list-specific indentation
            if list_props['left']:
                left_indent = float(int(list_props['left'])/1440)  # Convert twips to inches
                block.set('start-indent', f'{left_indent}in')
            
            if list_props['hanging']:
                hanging_indent = float(int(list_props['hanging'])/1440)  # Convert twips to inches
                block.set('text-indent', f'-{hanging_indent}in')
            
            # Add provisional-distance-between-starts for list items
            block.set('provisional-distance-between-starts', '0.25in')
            
            # Get the list numbers for this level
            level = int(list_props['level'])
            numbers = self._get_list_numbers(list_props['num_id'], level)
            
            # Format the list label
            label_text = self._format_list_label(list_props, numbers)
            
            # Add list label
            label_span = ET.SubElement(block, 'fo:inline')
            label_span.text = label_text
            
            # Add tab after label
            tab = ET.SubElement(block, 'fo:leader')
            tab.set('leader-pattern', 'space')
            tab.set('leader-length', '0.25in')
            
            return

        # Alignment
        jc = pPr.find('.//w:jc', self.ns)
        if jc is not None:
            align = jc.get(f'{{{self.ns["w"]}}}val')
            if align == 'center':
                block.set('text-align', 'center')
            elif align == 'right':
                block.set('text-align', 'right')
            elif align == 'justify':
                block.set('text-align', 'justify')
                
        # Indentation
        ind = pPr.find('.//w:ind', self.ns)
        if ind is not None:
            def round_indent(value, is_start_indent=False):
                """Helper function to round indents to common fractions"""
                if value is None:
                    return None
                    
                indent = float(int(value)/1440)  # Convert twips to inches
                
                # For start-indent, if greater than 1.0, return 0
                if is_start_indent:
                    if indent > 1.0:
                        return 0
                    # Force start-indent to closest quarter inch
                    quarters = [0, 0.25, 0.5, 0.75, 1.0]
                    return min(quarters, key=lambda x: abs(x - indent))
                
                # For other indents, use common fractions
                common_fractions = [0, 0.25, 0.5, 0.75, 1]
                
                # Find closest common fraction
                closest = min(common_fractions, key=lambda x: abs(x - indent))
                
                # If very close to a common fraction, use that value
                if abs(indent - closest) < 0.1:
                    return closest
                    
                # Otherwise round to 2 decimal places
                return round(indent, 2)
            
            # Left indent
            left = ind.get(f'{{{self.ns["w"]}}}left')
            if left:
                indent = round_indent(left, is_start_indent=True)
                if indent != 0:  # Only set if not zero
                    block.set('start-indent', f'{indent}in')
            
            # Right indent - skip setting end-indent entirely
            # right = ind.get(f'{{{self.ns["w"]}}}right')
            # if right:
            #     indent = round_indent(right)
            #     if indent != 0:  # Only set if not zero
            #         block.set('end-indent', f'{indent}in')
            
            # First line indent
            first = ind.get(f'{{{self.ns["w"]}}}firstLine')
            if first:
                indent = round_indent(first)
                if indent != 0:  # Only set if not zero
                    block.set('text-indent', f'{indent}in')
            
            # Hanging indent
            hanging = ind.get(f'{{{self.ns["w"]}}}hanging')
            if hanging:
                indent = round_indent(hanging)
                if indent != 0:  # Only set if not zero
                    block.set('text-indent', f'-{indent}in')
                
        # Removed the spacing/line-height section

    def _convert_paragraph(self, paragraph):
        """Convert Word paragraph to FO block with proper styling"""
        # Create block without default properties
        block = ET.Element('fo:block')
        
        # Get paragraph's default formatting from paragraph properties
        pPr = paragraph.find('.//w:pPr', self.ns)
        if pPr is not None:
            # Get paragraph-level font settings
            rPr = pPr.find('.//w:rPr', self.ns)
            if rPr is not None:
                format_props = self._get_run_properties(rPr)
                for prop, value in format_props.items():
                    block.set(prop, value)
            
            self._apply_paragraph_properties(block, pPr)
            
            # Modified tab handling:
            if block.get('text-indent') and float(block.get('text-indent').replace('in','')) < 0:
                text_indent = float(block.get('text-indent').replace('in',''))
                current_start_indent = float(block.get('start-indent', '0in').replace('in','')) if block.get('start-indent') else 0
                new_start_indent = current_start_indent + abs(text_indent)
                block.set('start-indent', f'{new_start_indent}in')
        
        # Reset current format tracking
        self.current_format = {}
        
        # Handle text runs
        runs = paragraph.findall('.//w:r', self.ns)
        text_content = ''.join(self._get_text_from_run(run) for run in runs).strip()
        
        if not text_content:  # If paragraph is empty or contains only whitespace
            block.text = '\u00A0'  # Add non-breaking space
        else:
            for run in runs:
                text = self._get_text_from_run(run)
                if text:
                    # Get run properties
                    rPr = run.find('.//w:rPr', self.ns)
                    format_props = self._get_run_properties(rPr)
                    
                    # Only create new inline if formatting differs from current
                    if format_props != self.current_format:
                        inline = ET.SubElement(block, 'fo:inline')
                        for prop, value in format_props.items():
                            inline.set(prop, value)
                        inline.text = text
                        self.current_format = format_props
                    else:
                        # Append to existing inline or block
                        last = block[-1] if len(block) > 0 else block
                        if last.text is None:
                            last.text = text
                        else:
                            last.text += text
        
        return block

    def _get_text_from_run(self, run):
        """Extract text from a run element"""
        text_elements = run.findall('.//w:t', self.ns)
        return ''.join(elem.text or '' for elem in text_elements)

    def convert(self, input_file, output_file):
        """Convert with debug logging"""
        self.debug_log = []  # Reset debug log
        
        # Log input file details
        self._log_debug('input_file', {
            'path': input_file,
            'size': os.path.getsize(input_file),
            'type': 'docx' if input_file.lower().endswith('.docx') else 'xml'
        })

        # Check if input is .docx or .xml
        is_docx = input_file.lower().endswith('.docx')
        
        if is_docx:
            # Parse input Word DOCX
            doc = Document(input_file)
            # Parse numbering definitions
            self._parse_numbering_definitions(doc)
            paragraphs = doc.paragraphs
        else:
            # Parse input Word XML
            tree = ET.parse(input_file)
            root = tree.getroot()
            
            # Try to find and parse numbering definitions from XML
            numbering_part = root.find('.//w:numbering', self.ns)
            if numbering_part is not None:
                # Parse abstract numbering definitions
                abstract_nums = {}
                for abstract_num in numbering_part.findall('.//w:abstractNum', self.ns):
                    num_id = abstract_num.get(f'{{{self.ns["w"]}}}abstractNumId')
                    levels = {}
                    
                    for level in abstract_num.findall('.//w:lvl', self.ns):
                        ilvl = level.get(f'{{{self.ns["w"]}}}ilvl')
                        num_fmt = level.find('.//w:numFmt', self.ns)
                        format_type = num_fmt.get(f'{{{self.ns["w"]}}}val') if num_fmt is not None else 'decimal'
                        
                        # Get level text (e.g., "%1.", "%1.%2.")
                        lvl_text = level.find('.//w:lvlText', self.ns)
                        text_format = lvl_text.get(f'{{{self.ns["w"]}}}val') if lvl_text is not None else "%1."
                        
                        # Get indentation
                        ind = level.find('.//w:ind', self.ns)
                        left = ind.get(f'{{{self.ns["w"]}}}left') if ind is not None else None
                        hanging = ind.get(f'{{{self.ns["w"]}}}hanging') if ind is not None else None
                        
                        levels[ilvl] = {
                            'format': format_type,
                            'text': text_format,
                            'left': left,
                            'hanging': hanging
                        }
                    
                    abstract_nums[num_id] = levels
                
                # Link concrete numbering instances to abstract definitions
                for num in numbering_part.findall('.//w:num', self.ns):
                    num_id = num.get(f'{{{self.ns["w"]}}}numId')
                    abstract_num_id = num.find('.//w:abstractNumId', self.ns).get(f'{{{self.ns["w"]}}}val')
                    self.numbering_defs[num_id] = abstract_nums[abstract_num_id]
            
            paragraphs = root.findall('.//w:p', self.ns)
        
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

        # Convert paragraphs
        for paragraph in paragraphs:
            if is_docx:
                block = self._convert_paragraph(paragraph)
            else:
                block = self._convert_xml_paragraph(paragraph)
            if block is not None:  # Only append if block is not None
                flow.append(block)
                # Add spacing block after each paragraph except the last one
                if paragraph != paragraphs[-1]:
                    spacing_block = ET.SubElement(flow, 'fo:block')
                    spacing_block.text = '\u00A0'
        
        # Write debug log
        self._dump_debug_log(output_file)

        # Write output
        tree = ET.ElementTree(fo_root)
        tree.write(output_file, encoding='utf-8', xml_declaration=True)

        # Log output file details
        self._log_debug('output_file', {
            'path': output_file,
            'size': os.path.getsize(output_file)
        })

    def _convert_xml_paragraph(self, paragraph):
        """Convert Word XML paragraph to FO block with debug logging"""
        # Analyze paragraph before conversion
        analysis = self._analyze_paragraph(paragraph)
        self._log_debug('paragraph_analysis', analysis)

        # Skip empty paragraphs
        if not analysis['text'].strip() and not analysis['runs']:
            return None

        # Create block without space-after
        block = ET.Element('fo:block')

        # Reset format tracking
        self.current_format = {}
        
        # Handle alignment first
        if analysis['alignment']:
            align = analysis['alignment']
            if align == 'center':
                block.set('text-align', 'center')
            elif align == 'right':
                block.set('text-align', 'right')
            elif align in ['both', 'justify']:
                block.set('text-align', 'justify')

        # Handle list properties
        if analysis['list_info']:
            level = int(analysis['list_info']['level'])
            num_id = analysis['list_info']['num_id']
            
            # Set consistent list indentation using start-indent
            indent = (level + 1) * 0.5  # 0.5 inch per level
            block.set('start-indent', f'{indent + 0.25}in')  # Add extra 0.25in for label
            block.set('text-indent', '-0.5in')  # Consistent hanging indent for label
            
            # Add list marker with proper numbering
            marker = ET.SubElement(block, 'fo:inline')
            
            if num_id in self.numbering_defs:
                level_def = self.numbering_defs[num_id].get(str(level))
                if level_def:
                    numbers = self._get_list_numbers(num_id, level)
                    marker.text = self._format_list_label(level_def, numbers)
                else:
                    marker.text = '• '
            else:
                marker.text = '• '

            # Add xf:tab after marker for proper spacing
            if marker.text and marker.text.strip():  # Only add tab if there's a marker
                tab = ET.SubElement(block, 'xf:tab')
                if self.current_format.get('text-underline-style'):
                    tab.set('fo:text-underline-style', 'none')

        # Handle regular paragraph indentation
        elif analysis['indentation']:
            start_indent = None
            text_indent = None
            
            for attr, value in analysis['indentation'].items():
                if attr == 'left' and value:
                    start_indent = value
                    block.set('start-indent', f'{value:.2f}in')
                elif attr == 'firstLine' and value:
                    text_indent = value
                    block.set('text-indent', f'{value:.2f}in')
                elif attr == 'hanging' and value:
                    text_indent = -value
                    block.set('text-indent', f'-{value:.2f}in')
            
            # Add xf:tab if we have both positive start-indent and negative text-indent
            if start_indent and text_indent and text_indent < 0:
                tab = ET.SubElement(block, 'xf:tab')
                if self.current_format.get('text-underline-style'):
                    tab.set('fo:text-underline-style', 'none')

        # Process runs and collect text with formatting
        accumulated_text = []
        
        for run in analysis['runs']:
            text = run['text']
            if not text.strip():  # Skip empty runs
                continue
                
            format_props = run['formatting']
            
            # Check if format changed
            if format_props != self.current_format:
                # Add accumulated text with current formatting
                if accumulated_text:
                    inline = ET.SubElement(block, 'fo:inline')
                    for prop, value in self.current_format.items():
                        inline.set(prop, value)
                    inline.text = ''.join(accumulated_text)
                    accumulated_text = []
                self.current_format = format_props
            
            accumulated_text.append(text)
        
        # Add any remaining text
        if accumulated_text:
            inline = ET.SubElement(block, 'fo:inline')
            for prop, value in self.current_format.items():
                inline.set(prop, value)
            inline.text = ''.join(accumulated_text)

        # Special handling for footers and page numbers
        if analysis['text'].strip().startswith('Page'):
            block.set('font-size', '9pt')
            block.set('text-align', 'center')
            
        # Special handling for initial/signature lines
        if 'Initial' in analysis['text']:
            block.set('text-align', 'right')
            block.set('space-before', '24pt')
            
        # Log the resulting FO block
        self._log_debug('fo_block', {
            'attributes': dict(block.attrib),
            'children': [self._element_to_dict(child) for child in block],
            'text': block.text
        })

        return block

    def _element_to_dict(self, element):
        """Convert XML element to dictionary for debugging"""
        result = {
            'tag': element.tag,
            'attributes': dict(element.attrib),
            'text': element.text,
            'tail': element.tail,
            'children': []
        }
        for child in element:
            result['children'].append(self._element_to_dict(child))
        return result

    def _get_list_numbers(self, num_id, level):
        """Get the list numbers for the current level"""
        # Get the numbering definition for this list
        level_def = None
        if num_id in self.numbering_defs:
            level_def = self.numbering_defs[num_id].get(str(level))

        # Check if this is a letter-based list
        is_letter_list = level_def and level_def.get('format') in ['lowerLetter', 'upperLetter']
        
        # For letter lists, use a special counter
        if is_letter_list:
            if 'letter_counter' not in self.list_counters:
                self.list_counters['letter_counter'] = 1
            number = self.list_counters['letter_counter']
            self.list_counters['letter_counter'] += 1
            return [number]  # Return single number for letter formatting
            
        # For regular numbering, use section-based tracking
        if level == 0:
            num_id = 'top_level'
        else:
            # For sub-levels, always use the current section number
            current_section = self.list_counters.get('top_level', {}).get('counters', [1])[0]
            num_id = f'section_{current_section}'
            
        if num_id not in self.list_counters:
            self.list_counters[num_id] = {
                'counters': [1] * 9,  # Initialize with 1s
                'last_level': -1,
                'parent_section': None if level == 0 else current_section
            }
        
        list_info = self.list_counters[num_id]
        counters = list_info['counters']
        
        # If this is a sub-level, ensure we're using the correct section
        if level > 0:
            current_section = self.list_counters['top_level']['counters'][0]
            if list_info['parent_section'] != current_section:
                # Section changed, reset this counter
                list_info['parent_section'] = current_section
                counters = [current_section] + [1] * 8
                list_info['counters'] = counters
                list_info['last_level'] = level - 1  # Set to previous level to trigger increment
        
        # Handle level transitions
        if level > list_info['last_level']:
            # Moving to deeper level, keep parent numbers
            for i in range(level + 1, len(counters)):
                counters[i] = 1
        elif level < list_info['last_level']:
            # Moving to higher level, increment and reset deeper
            counters[level] += 1
            for i in range(level + 1, len(counters)):
                counters[i] = 1
        else:
            # Same level, just increment
            counters[level] += 1
            
        list_info['last_level'] = level
        
        # Always ensure first number matches current section for sub-levels
        if level > 0:
            counters[0] = self.list_counters['top_level']['counters'][0]
        
        return counters[:level + 1]

def main():
    # Check if command line arguments are provided
    if len(sys.argv) > 1:
        # Use command line arguments
        parser = argparse.ArgumentParser(description='Convert Word XML to XSL-FO format')
        parser.add_argument('input', help='Input Word XML file path')
        parser.add_argument('output', help='Output XSL-FO file path')
        args = parser.parse_args()
        input_file = args.input
        output_file = args.output
    else:
        # Use file selection dialog
        input_file, output_file = select_files()
        if not input_file or not output_file:
            print("File selection cancelled", file=sys.stderr)
            sys.exit(1)
    
    # Validate input file
    if not os.path.exists(input_file):
        error_msg = f"Error: Input file '{input_file}' does not exist"
        if len(sys.argv) > 1:
            print(error_msg, file=sys.stderr)
        else:
            messagebox.showerror("Error", error_msg)
        sys.exit(1)
        
    # Validate input file is XML
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read(100)  # Read first 100 chars to check XML declaration
            if not content.strip().startswith('<?xml'):
                error_msg = f"Error: Input file '{input_file}' does not appear to be an XML file"
                if len(sys.argv) > 1:
                    print(error_msg, file=sys.stderr)
                else:
                    messagebox.showerror("Error", error_msg)
                sys.exit(1)
    except UnicodeDecodeError:
        error_msg = f"Error: Input file '{input_file}' is not a valid UTF-8 encoded text file"
        if len(sys.argv) > 1:
            print(error_msg, file=sys.stderr)
        else:
            messagebox.showerror("Error", error_msg)
        sys.exit(1)
    except Exception as e:
        error_msg = f"Error reading input file: {str(e)}"
        if len(sys.argv) > 1:
            print(error_msg, file=sys.stderr)
        else:
            messagebox.showerror("Error", error_msg)
        sys.exit(1)
        
    # Validate output directory exists
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except Exception as e:
            error_msg = f"Error creating output directory: {str(e)}"
            if len(sys.argv) > 1:
                print(error_msg, file=sys.stderr)
            else:
                messagebox.showerror("Error", error_msg)
            sys.exit(1)
            
    # Perform conversion
    try:
        converter = WordToFOConverter()
        converter.convert(input_file, output_file)
        success_msg = f"Successfully converted '{input_file}' to '{output_file}'"
        if len(sys.argv) > 1:
            print(success_msg)
        else:
            messagebox.showinfo("Success", success_msg)
    except ET.ParseError as e:
        error_msg = f"Error parsing XML file: {str(e)}"
        if len(sys.argv) > 1:
            print(error_msg, file=sys.stderr)
        else:
            messagebox.showerror("Error", error_msg)
        sys.exit(1)
    except Exception as e:
        error_msg = f"Error during conversion: {str(e)}"
        if len(sys.argv) > 1:
            print(error_msg, file=sys.stderr)
        else:
            messagebox.showerror("Error", error_msg)
        sys.exit(1)

if __name__ == '__main__':
    main() 