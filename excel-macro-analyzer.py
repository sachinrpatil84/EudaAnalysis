import os
import re
import sys
import pandas as pd
import win32com.client
from win32com.client import constants
import pythoncom
import logging
from pathlib import Path
import argparse
import time
import numpy as np
from collections import defaultdict

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("macro_analyzer.log"),
        logging.StreamHandler()
    ]
)

class ExcelMacroAnalyzer:
    """
    A class to analyze Excel macros, extract their code, and determine their purpose and complexity.
    """
    
    def __init__(self):
        self.excel = None
        self.macro_info = []
        self.keywords = {
            "data_manipulation": ["range", "cells", "select", "selection", "copy", "paste", "cut", "clearcontents", "sort", "filter", "autofilter", "removeduplicates"],
            "formatting": ["font", "interior", "color", "borders", "bold", "italic", "underline", "style", "format", "numberformat"],
            "file_operations": ["workbooks.open", "save", "saveas", "close", "workbooks.add", "workbooks.close"],
            "ui_interaction": ["msgbox", "inputbox", "userform", "dialog", "prompt", "alert"],
            "reporting": ["print", "printout", "preview", "report", "header", "footer"],
            "automation": ["application.run", "ontime", "scheduled", "timer", "auto", "automatic"],
            "calculation": ["calculate", "sum", "average", "count", "max", "min", "if", "vlookup", "hlookup", "match", "index"]
        }
    
    def initialize_excel(self):
        """Initialize Excel application object"""
        try:
            pythoncom.CoInitialize()
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            logging.info("Excel application initialized successfully")
        except Exception as e:
            logging.error(f"Failed to initialize Excel: {str(e)}")
            sys.exit(1)
    
    def close_excel(self):
        """Close Excel application"""
        if self.excel:
            try:
                self.excel.Quit()
                pythoncom.CoUninitialize()
                logging.info("Excel application closed successfully")
            except Exception as e:
                logging.error(f"Error closing Excel: {str(e)}")
    
    def extract_macros(self, file_path):
        """
        Extract macros from Excel file
        
        Args:
            file_path (str): Path to the Excel file
        """
        try:
            # Reset macro_info for new file
            self.macro_info = []
            
            file_path = os.path.abspath(file_path)
            logging.info(f"Opening Excel file: {file_path}")
            
            # Check if file exists
            if not os.path.exists(file_path):
                logging.error(f"File not found: {file_path}")
                return
            
            # Check if file is an Excel file
            if not file_path.lower().endswith(('.xls', '.xlsm', '.xlsb')):
                logging.error(f"Not an Excel file: {file_path}")
                return
            
            # Open workbook
            wb = self.excel.Workbooks.Open(file_path)
            
            try:
                # Try to access VBA project
                vbp = wb.VBProject
            except Exception as e:
                logging.error("Cannot access VBA project. Make sure 'Trust access to the VBA project object model' is enabled in Excel Trust Center settings.")
                wb.Close(False)
                return
            
            # Process each component (module, class module, etc.)
            for comp in vbp.VBComponents:
                try:
                    component_type = self._get_component_type(comp.Type)
                    module_name = comp.Name
                    
                    # Get code module
                    code_module = comp.CodeModule
                    
                    if code_module.CountOfLines == 0:
                        continue  # Skip empty modules
                    
                    # Extract the entire module code
                    full_code = code_module.Lines(1, code_module.CountOfLines)
                    
                    # Extract individual procedures (macros/functions)
                    self._extract_procedures(full_code, module_name, component_type)
                    
                except Exception as e:
                    logging.error(f"Error processing component {comp.Name}: {str(e)}")
            
            # Close workbook without saving
            wb.Close(False)
            
        except Exception as e:
            logging.error(f"Error extracting macros: {str(e)}")
            if 'wb' in locals():
                wb.Close(False)
    
    def _get_component_type(self, type_num):
        """Convert VBA component type number to string"""
        types = {
            1: "Standard Module",
            2: "Class Module",
            3: "UserForm",
            100: "Document Module"
        }
        return types.get(type_num, "Unknown")
    
    def _extract_procedures(self, code, module_name, component_type):
        """
        Extract individual procedures from module code
        
        Args:
            code (str): Full module code
            module_name (str): Name of the module
            component_type (str): Type of the component
        """
        # Pattern to match Sub, Function, or Property declarations
        proc_pattern = r'(?:Public |Private |Friend )?(?:Sub|Function|Property Get|Property Let|Property Set)\s+([^\(]+)'
        
        # Pattern to match end of procedures
        end_pattern = r'End (?:Sub|Function|Property)'
        
        lines = code.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            proc_match = re.search(proc_pattern, line)
            
            if proc_match:
                # Found procedure start
                proc_name = proc_match.group(1).strip()
                proc_type = re.search(r'(Sub|Function|Property\s+(?:Get|Let|Set))', line).group(1)
                
                # Extract parameters
                params_match = re.search(r'\((.*?)\)', line)
                params = params_match.group(1) if params_match else ""
                
                # Procedure starts at current line
                start_line = i + 1
                proc_code = [line]
                i += 1
                
                # Find the end of the procedure
                while i < len(lines) and not re.search(end_pattern, lines[i].strip()):
                    proc_code.append(lines[i].strip())
                    i += 1
                
                # Add the End Sub/Function line
                if i < len(lines):
                    proc_code.append(lines[i].strip())
                
                # Calculate procedure info
                proc_code_str = '\n'.join(proc_code)
                line_count = len(proc_code)
                complexity_score = self._calculate_complexity(proc_code_str)
                purpose = self._determine_purpose(proc_code_str)
                
                # Add to macro info list
                self.macro_info.append({
                    'Name': proc_name,
                    'Module': module_name,
                    'Line Count': line_count,
                    'Complexity Score': complexity_score,
                    'Purpose Description': purpose['description'],
                    'Type': purpose['type'],
                    'Component Type': component_type,
                    'Code': proc_code_str
                })
            
            i += 1
    
    def _calculate_complexity(self, code):
        """
        Calculate cyclomatic complexity score for the macro
        
        Args:
            code (str): Code of the macro
        
        Returns:
            int: Complexity score
        """
        # Base complexity is 1
        complexity = 1
        
        # Decision points that increase complexity
        decision_patterns = [
            r'\bIf\b.*\bThen\b',  # If statements
            r'\bElseIf\b',         # ElseIf branches
            r'\bCase\b',           # Case statements
            r'\bFor\b',            # For loops
            r'\bDo\b.*\bWhile\b',  # Do While loops
            r'\bWhile\b',          # While loops
            r'\bSelect\b',         # Select Case statements
            r'\bAnd\b',            # Logical AND operations
            r'\bOr\b'              # Logical OR operations
        ]
        
        for pattern in decision_patterns:
            complexity += len(re.findall(pattern, code, re.IGNORECASE))
        
        # Nested structures add more complexity
        nest_level = 0
        max_nest = 0
        
        for line in code.split('\n'):
            line = line.strip()
            
            # Increase nesting level
            if any(re.search(r'\b{}\b'.format(p), line, re.IGNORECASE) for p in ['If', 'For', 'Do', 'While', 'Select']):
                nest_level += 1
            
            # Decrease nesting level
            if any(re.search(r'\b{}\b'.format(p), line, re.IGNORECASE) for p in ['End If', 'Next', 'Loop', 'Wend', 'End Select']):
                nest_level = max(0, nest_level - 1)
            
            max_nest = max(max_nest, nest_level)
        
        # Add nesting factor to complexity
        complexity += max_nest * 0.5
        
        # Add complexity for error handling
        if re.search(r'\bOn\s+Error\b', code, re.IGNORECASE):
            complexity += 1
        
        # Add complexity for API calls
        if re.search(r'\bDeclare\b', code, re.IGNORECASE):
            complexity += 2
        
        return round(complexity, 1)
    
    def _determine_purpose(self, code):
        """
        Determine the purpose of the macro based on its code
        
        Args:
            code (str): Code of the macro
        
        Returns:
            dict: Purpose description and type
        """
        # Convert code to lowercase for easier matching
        code_lower = code.lower()
        
        # Count keyword occurrences
        keyword_counts = {}
        for category, keywords in self.keywords.items():
            count = sum(code_lower.count(kw.lower()) for kw in keywords)
            keyword_counts[category] = count
        
        # Determine primary and secondary categories
        sorted_categories = sorted(keyword_counts.items(), key=lambda x: x[1], reverse=True)
        primary_category = sorted_categories[0][0] if sorted_categories[0][1] > 0 else "unknown"
        
        # Generate description based on code analysis
        descriptions = {
            "data_manipulation": "Manipulates worksheet data through operations like selecting, copying, filtering, or sorting",
            "formatting": "Applies formatting to cells or ranges including colors, fonts, styles, or number formats",
            "file_operations": "Manages Excel files through opening, saving, or closing workbooks",
            "ui_interaction": "Interacts with users through message boxes, input forms, or dialogs",
            "reporting": "Generates or prepares reports for printing or distribution",
            "automation": "Automates processes or runs scheduled/timed operations",
            "calculation": "Performs calculations, lookups, or data analysis",
            "unknown": "Purpose could not be clearly determined from the code"
        }
        
        # Determine specific actions in the code
        specific_actions = []
        
        # Check for sheet manipulation
        if re.search(r'\b(sheet|worksheet|activesheet)\b', code_lower):
            specific_actions.append("worksheet manipulation")
        
        # Check for data validation
        if re.search(r'\b(validation|valid|check|verify)\b', code_lower):
            specific_actions.append("data validation")
        
        # Check for external data connections
        if re.search(r'\b(connection|query|sql|odbc|oledb|adodb)\b', code_lower):
            specific_actions.append("external data connections")
        
        # Check for chart manipulation
        if re.search(r'\b(chart|plot|graph|series)\b', code_lower):
            specific_actions.append("chart manipulation")
        
        # Check for pivot table operations
        if re.search(r'\b(pivot|pivotcache|pivottable|pivotfield)\b', code_lower):
            specific_actions.append("pivot table operations")
        
        # Additional specific purpose markers
        specific_purpose = ""
        if re.search(r'\bexport\b', code_lower):
            specific_purpose = "Exports data "
            if re.search(r'\bpdf\b', code_lower):
                specific_purpose += "to PDF"
            elif re.search(r'\bcsv\b', code_lower):
                specific_purpose += "to CSV"
            elif re.search(r'\btext\b', code_lower) or re.search(r'\btxt\b', code_lower):
                specific_purpose += "to text file"
            else:
                specific_purpose += "to another format"
        elif re.search(r'\bimport\b', code_lower):
            specific_purpose = "Imports data from external sources"
        
        # Create purpose description
        base_description = descriptions[primary_category]
        
        if specific_purpose:
            purpose_description = f"{specific_purpose}. Additionally, {base_description.lower()}"
        else:
            purpose_description = base_description
            if specific_actions:
                purpose_description += f", specifically involving {', '.join(specific_actions)}"
        
        # Create type label
        type_label = primary_category.replace('_', ' ').title()
        
        return {
            "description": purpose_description,
            "type": type_label
        }
    
    def generate_report(self, output_format='csv', output_file=None):
        """
        Generate a report of macro information
        
        Args:
            output_format (str): Format of the output ('csv', 'excel', or 'console')
            output_file (str): Path to the output file
        """
        if not self.macro_info:
            logging.warning("No macros found to report")
            return
        
        # Create DataFrame
        df = pd.DataFrame(self.macro_info)
        
        # Select only the requested columns
        columns = ['Name', 'Module', 'Line Count', 'Complexity Score', 'Purpose Description', 'Type']
        df = df[columns]
        
        if output_format == 'console':
            pd.set_option('display.max_colwidth', None)
            print("\n" + "="*100)
            print("EXCEL MACRO ANALYSIS REPORT")
            print("="*100)
            print(df.to_string(index=False))
            print("="*100)
        
        elif output_format == 'csv':
            if not output_file:
                output_file = f"macro_analysis_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            df.to_csv(output_file, index=False)
            logging.info(f"Report saved to {output_file}")
        
        elif output_format == 'excel':
            if not output_file:
                output_file = f"macro_analysis_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            # Create Excel writer
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Macro Analysis', index=False)
                
                # Get workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Macro Analysis']
                
                # Add formats
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'bg_color': '#D7E4BC',
                    'border': 1
                })
                
                # Apply formatting
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Set column widths
                worksheet.set_column('A:A', 20)  # Name
                worksheet.set_column('B:B', 15)  # Module
                worksheet.set_column('C:C', 10)  # Line Count
                worksheet.set_column('D:D', 15)  # Complexity Score
                worksheet.set_column('E:E', 50)  # Purpose Description
                worksheet.set_column('F:F', 15)  # Type
            
            logging.info(f"Report saved to {output_file}")
    
    def analyze_file(self, file_path, output_format='console', output_file=None):
        """
        Analyze a single Excel file
        
        Args:
            file_path (str): Path to the Excel file
            output_format (str): Format of the output ('csv', 'excel', or 'console')
            output_file (str): Path to the output file
        """
        try:
            self.initialize_excel()
            self.extract_macros(file_path)
            self.generate_report(output_format, output_file)
        finally:
            self.close_excel()
    
    def analyze_directory(self, directory_path, output_format='excel', output_file=None):
        """
        Analyze all Excel files in a directory
        
        Args:
            directory_path (str): Path to the directory
            output_format (str): Format of the output ('csv', 'excel', or 'console')
            output_file (str): Path to the output file
        """
        try:
            self.initialize_excel()
            
            directory = Path(directory_path)
            all_macro_info = []
            
            # Process each Excel file
            excel_files = list(directory.glob('*.xls*'))
            if not excel_files:
                logging.warning(f"No Excel files found in {directory_path}")
                return
            
            for file_path in excel_files:
                if file_path.suffix.lower() in ['.xls', '.xlsm', '.xlsb']:
                    logging.info(f"Processing file: {file_path}")
                    self.extract_macros(str(file_path))
                    all_macro_info.extend(self.macro_info)
                    # Reset for next file
                    self.macro_info = []
            
            # Store all results and generate report
            self.macro_info = all_macro_info
            self.generate_report(output_format, output_file)
            
        finally:
            self.close_excel()


def main():
    parser = argparse.ArgumentParser(description='Excel Macro Analyzer')
    parser.add_argument('path', help='Path to Excel file or directory containing Excel files')
    parser.add_argument('-o', '--output', choices=['console', 'csv', 'excel'], default='console',
                      help='Output format (default: console)')
    parser.add_argument('-f', '--file', help='Output file path')
    parser.add_argument('-r', '--recursive', action='store_true', help='Recursively process directories')
    
    args = parser.parse_args()
    
    analyzer = ExcelMacroAnalyzer()
    
    path = args.path
    if os.path.isfile(path):
        # Analyze single file
        analyzer.analyze_file(path, args.output, args.file)
    elif os.path.isdir(path):
        # Analyze directory
        analyzer.analyze_directory(path, args.output, args.file)
    else:
        logging.error(f"Path not found: {path}")


if __name__ == "__main__":
    main()
