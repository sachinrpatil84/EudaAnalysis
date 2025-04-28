#!/usr/bin/env python3
"""
Excel Database Connection Analyzer
=================================
This script analyzes Excel files to extract database connections, connection strings,
queries, and other metadata. It works with both modern Excel files (.xlsx) and
legacy formats (.xls).

The analyzer extracts:
- Connection names
- Connection types (ODBC, OLEDB, etc.)
- Connection strings
- Target data sources
- SQL queries
- Worksheet references
- Complexity metrics
- Purpose descriptions (inferred)

Usage:
    python excel_connection_analyzer.py path/to/excel_file.xlsx
"""

import os
import re
import sys
import argparse
import csv
import json
import pandas as pd
from typing import Dict, List, Tuple, Any, Optional
import xml.etree.ElementTree as ET
from pathlib import Path
import zipfile
import olefile
import uuid
import logging
from collections import defaultdict

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('excel_connection_analyzer')

# Define connection types and their patterns
CONNECTION_PATTERNS = {
    'ODBC': [r'ODBC', r'DSN='],
    'OLEDB': [r'Provider=', r'OLEDB', r'Microsoft\.ACE\.OLEDB', r'Microsoft\.Jet\.OLEDB'],
    'Power Query': [r'Power\s*Query', r'M Formula'],
    'CSV': [r'\.csv', r'text/csv', r'TextCSV'],
    'Web': [r'http://', r'https://', r'www\.'],
    'SharePoint': [r'sharepoint', r'SharePoint'],
    'REST': [r'REST', r'API', r'JSON'],
    'Oracle': [r'Oracle', r'OraOLEDB'],
    'SQL Server': [r'SQL Server', r'SQLServer', r'SQLOLEDB', r'MSOLEDBSQL'],
    'MySQL': [r'MySQL'],
    'PostgreSQL': [r'PostgreSQL', r'PgOLEDB'],
    'Access': [r'Access', r'\.mdb', r'\.accdb'],
    'Text': [r'\.txt', r'text/plain'],
    'Excel': [r'\.xls', r'\.xlsx', r'Excel Files'],
}

# SQL keywords for complexity analysis
SQL_KEYWORDS = [
    'SELECT', 'FROM', 'WHERE', 'JOIN', 'LEFT JOIN', 'RIGHT JOIN', 'INNER JOIN', 
    'FULL JOIN', 'CROSS JOIN', 'GROUP BY', 'HAVING', 'ORDER BY', 'DISTINCT',
    'UNION', 'INTERSECT', 'EXCEPT', 'WITH', 'CTE', 'CASE', 'WHEN', 'THEN',
    'ELSE', 'END', 'OVER', 'PARTITION BY', 'ROW_NUMBER', 'RANK', 'DENSE_RANK',
    'NTILE', 'LEAD', 'LAG', 'FIRST_VALUE', 'LAST_VALUE', 'SUBQUERY', 'EXISTS',
    'NOT EXISTS', 'IN', 'NOT IN', 'ALL', 'ANY', 'SOME', 'MERGE', 'PIVOT', 'UNPIVOT'
]

class ExcelConnectionAnalyzer:
    def __init__(self, file_path: str):
        """Initialize the analyzer with the path to an Excel file."""
        self.file_path = file_path
        self.connections = []
        self.xlsx_format = file_path.lower().endswith('.xlsx')
        self.temp_dir = None
        self.workbook_details = None
        
    def analyze(self) -> List[Dict[str, Any]]:
        """Analyze the Excel file and return connection details."""
        try:
            if self.xlsx_format:
                self._analyze_xlsx()
            else:
                self._analyze_xls()
                
            # Post-process connections to infer purposes and calculate complexity
            for conn in self.connections:
                self._calculate_complexity(conn)
                self._infer_purpose(conn)
                
            return self.connections
        except Exception as e:
            logger.error(f"Error analyzing file {self.file_path}: {str(e)}")
            raise
    
    def _analyze_xlsx(self):
        """Extract connection information from .xlsx files."""
        try:
            # Extract connection information from the workbook's internal XML files
            with zipfile.ZipFile(self.file_path) as zip_ref:
                # Extract workbook.xml to get sheets info
                try:
                    workbook_xml = zip_ref.read('xl/workbook.xml').decode('utf-8')
                    workbook_root = ET.fromstring(workbook_xml)
                    sheet_names = {}
                    
                    # Extract namespace information from the root element
                    namespaces = dict([node for _, node in ET.iterparse(zip_ref.open('xl/workbook.xml'), events=['start-ns'])])
                    
                    # Use the main namespace for queries
                    main_ns = '{' + namespaces.get('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') + '}'
                    
                    # Get sheet information
                    sheets = workbook_root.findall(f'.//{main_ns}sheet')
                    for sheet in sheets:
                        sheet_id = sheet.attrib.get('sheetId', '')
                        sheet_name = sheet.attrib.get('name', '')
                        sheet_names[sheet_id] = sheet_name
                    
                    self.workbook_details = {'sheet_names': sheet_names}
                except Exception as e:
                    logger.warning(f"Could not extract sheet information: {str(e)}")
                    self.workbook_details = {'sheet_names': {}}
                
                # Process connections.xml if it exists
                if 'xl/connections.xml' in zip_ref.namelist():
                    connections_xml = zip_ref.read('xl/connections.xml').decode('utf-8')
                    self._process_connections_xml(connections_xml)
                
                # Process queries if they exist
                if 'xl/queries/query1.xml' in zip_ref.namelist():
                    for file_name in zip_ref.namelist():
                        if file_name.startswith('xl/queries/query'):
                            query_xml = zip_ref.read(file_name).decode('utf-8')
                            self._process_query_xml(query_xml)
                
                # Check for data connections in custom XML parts
                for file_name in zip_ref.namelist():
                    if 'customXml/' in file_name and file_name.endswith('.xml'):
                        try:
                            custom_xml = zip_ref.read(file_name).decode('utf-8')
                            self._process_custom_xml(custom_xml)
                        except Exception as e:
                            logger.warning(f"Error processing custom XML {file_name}: {str(e)}")
                
                # Extract any connection information from VBA if present
                if 'xl/vbaProject.bin' in zip_ref.namelist():
                    vba_data = zip_ref.read('xl/vbaProject.bin')
                    self._extract_connections_from_vba(vba_data)
                
                # Look for PowerQuery/M formulas
                if any(name.startswith('xl/queries/') for name in zip_ref.namelist()):
                    for file_name in zip_ref.namelist():
                        if file_name.startswith('xl/queries/'):
                            try:
                                query_data = zip_ref.read(file_name).decode('utf-8')
                                self._process_power_query(query_data, file_name)
                            except Exception as e:
                                logger.warning(f"Error processing Power Query {file_name}: {str(e)}")
                
                # Check for potential connections in worksheets' data connections
                for file_name in zip_ref.namelist():
                    if file_name.startswith('xl/worksheets/sheet'):
                        try:
                            sheet_xml = zip_ref.read(file_name).decode('utf-8')
                            sheet_number = re.search(r'sheet(\d+)\.xml', file_name)
                            sheet_id = sheet_number.group(1) if sheet_number else "unknown"
                            sheet_name = self.workbook_details['sheet_names'].get(sheet_id, f"Sheet{sheet_id}")
                            self._extract_connections_from_sheet(sheet_xml, sheet_name)
                        except Exception as e:
                            logger.warning(f"Error processing sheet {file_name}: {str(e)}")
        
        except zipfile.BadZipFile:
            logger.error(f"The file {self.file_path} is not a valid XLSX file")
            raise ValueError(f"The file {self.file_path} is not a valid XLSX file")
    
    def _analyze_xls(self):
        """Extract connection information from .xls (legacy format) files."""
        try:
            if not olefile.isOleFile(self.file_path):
                logger.error(f"The file {self.file_path} is not a valid XLS file")
                raise ValueError(f"The file {self.file_path} is not a valid XLS file")
            
            # Process the OLE file
            with olefile.OleFile(self.file_path) as ole:
                # Check for DBC (Database Connectivity) streams
                for stream_name in ole.listdir():
                    stream_path = '/'.join(stream_name)
                    
                    # Look for known database connection streams
                    if 'Workbook' in stream_path or 'Book' in stream_path:
                        try:
                            workbook_data = ole.openstream(stream_path).read()
                            self._extract_connections_from_workbook_binary(workbook_data)
                        except Exception as e:
                            logger.warning(f"Error extracting from workbook stream {stream_path}: {str(e)}")
                    
                    # Extract VBA code that might contain connections
                    if 'VBA' in stream_path or 'vba' in stream_path:
                        try:
                            vba_data = ole.openstream(stream_path).read()
                            self._extract_connections_from_vba(vba_data)
                        except Exception as e:
                            logger.warning(f"Error extracting from VBA stream {stream_path}: {str(e)}")
                    
                    # Look for ODBC or OLEDB related streams
                    if any(keyword in stream_path.lower() for keyword in ['odbc', 'oledb', 'dbc', 'connect', 'query']):
                        try:
                            connection_data = ole.openstream(stream_path).read()
                            self._process_ole_connection_stream(connection_data, stream_path)
                        except Exception as e:
                            logger.warning(f"Error processing connection stream {stream_path}: {str(e)}")
        
        except Exception as e:
            logger.error(f"Error analyzing XLS file {self.file_path}: {str(e)}")
            raise
    
    def _process_connections_xml(self, xml_content: str):
        """Process the connections.xml file from an XLSX package."""
        try:
            root = ET.fromstring(xml_content)
            namespace = root.tag.split('}')[0].strip('{') if '}' in root.tag else ''
            ns_prefix = '{' + namespace + '}' if namespace else ''
            
            connections = root.findall(f'.//{ns_prefix}connection') or root.findall('.//connection')
            
            for conn_elem in connections:
                connection = {
                    'Name': conn_elem.get('name', 'Unnamed Connection'),
                    'Type': 'Unknown',
                    'ConnectionString': '',
                    'TargetDataSources': [],
                    'QueryText': '',
                    'WorksheetName': '',
                    'ComplexityScore': 0,
                    'PurposeDescription': ''
                }
                
                # Extract connection string
                conn_string_elem = conn_elem.find(f'.//{ns_prefix}dbPr') or conn_elem.find('.//dbPr')
                if conn_string_elem is not None:
                    connection['ConnectionString'] = conn_string_elem.get('connection', '')
                    connection['Type'] = self._determine_connection_type(connection['ConnectionString'])
                
                # Extract query text
                command_text_elem = conn_elem.find(f'.//{ns_prefix}commandText') or conn_elem.find('.//commandText')
                if command_text_elem is not None and command_text_elem.text:
                    connection['QueryText'] = command_text_elem.text
                
                # Extract target worksheets
                tables_elem = conn_elem.find(f'.//{ns_prefix}tables') or conn_elem.find('.//tables')
                if tables_elem is not None:
                    table_elems = tables_elem.findall(f'.//{ns_prefix}table') or tables_elem.findall('.//table')
                    for table in table_elems:
                        table_name = table.get('name', '')
                        if table_name:
                            connection['WorksheetName'] = table_name
                
                # Extract data source
                if connection['ConnectionString']:
                    data_sources = self._extract_data_sources(connection['ConnectionString'])
                    if data_sources:
                        connection['TargetDataSources'] = data_sources
                
                self.connections.append(connection)
        
        except Exception as e:
            logger.warning(f"Error processing connections XML: {str(e)}")
    
    def _process_query_xml(self, xml_content: str):
        """Process query XML files that contain connection information."""
        try:
            root = ET.fromstring(xml_content)
            namespace = root.tag.split('}')[0].strip('{') if '}' in root.tag else ''
            ns_prefix = '{' + namespace + '}' if namespace else ''
            
            # Look for connection details
            connection = {
                'Name': 'Query Connection',
                'Type': 'Unknown',
                'ConnectionString': '',
                'TargetDataSources': [],
                'QueryText': '',
                'WorksheetName': '',
                'ComplexityScore': 0,
                'PurposeDescription': ''
            }
            
            # Extract query name
            query_elem = root.find(f'.//{ns_prefix}queryName') or root.find('.//queryName')
            if query_elem is not None and query_elem.text:
                connection['Name'] = query_elem.text
            
            # Extract connection details
            connection_elem = root.find(f'.//{ns_prefix}connection') or root.find('.//connection')
            if connection_elem is not None:
                connection_string = connection_elem.get('connectionString', '')
                if connection_string:
                    connection['ConnectionString'] = connection_string
                    connection['Type'] = self._determine_connection_type(connection_string)
                    data_sources = self._extract_data_sources(connection_string)
                    if data_sources:
                        connection['TargetDataSources'] = data_sources
            
            # Extract query text
            command_text = root.find(f'.//{ns_prefix}commandText') or root.find('.//commandText')
            if command_text is not None and command_text.text:
                connection['QueryText'] = command_text.text
            
            if connection['ConnectionString'] or connection['QueryText']:
                self.connections.append(connection)
        
        except Exception as e:
            logger.warning(f"Error processing query XML: {str(e)}")
    
    def _process_custom_xml(self, xml_content: str):
        """Process custom XML parts that might contain connection information."""
        try:
            # Look for connection strings or queries in the custom XML
            root = ET.fromstring(xml_content)
            
            # Check for elements that might contain connection information
            for elem in root.iter():
                # Look for connection attributes
                for attr_name, attr_value in elem.attrib.items():
                    if any(keyword in attr_name.lower() for keyword in ['connect', 'source', 'query', 'db', 'data']):
                        if self._looks_like_connection_string(attr_value):
                            connection = {
                                'Name': f"Custom XML Connection {len(self.connections) + 1}",
                                'Type': self._determine_connection_type(attr_value),
                                'ConnectionString': attr_value,
                                'TargetDataSources': self._extract_data_sources(attr_value),
                                'QueryText': '',
                                'WorksheetName': '',
                                'ComplexityScore': 0,
                                'PurposeDescription': ''
                            }
                            self.connections.append(connection)
                
                # Check for SQL-like text content
                if elem.text and self._looks_like_sql_query(elem.text):
                    connection = {
                        'Name': f"Custom XML Query {len(self.connections) + 1}",
                        'Type': 'SQL',
                        'ConnectionString': '',
                        'TargetDataSources': [],
                        'QueryText': elem.text,
                        'WorksheetName': '',
                        'ComplexityScore': 0,
                        'PurposeDescription': ''
                    }
                    self.connections.append(connection)
        
        except Exception as e:
            logger.warning(f"Error processing custom XML: {str(e)}")
    
    def _extract_connections_from_vba(self, vba_data: bytes):
        """Extract connection information from VBA code."""
        try:
            # Convert binary data to text, ignoring errors
            vba_text = vba_data.decode('latin-1', errors='ignore')
            
            # Look for connection strings
            self._find_connection_strings_in_text(vba_text, 'VBA Code')
            
            # Look for SQL queries
            self._find_sql_queries_in_text(vba_text, 'VBA Code')
            
            # Look for specific connection patterns
            adodb_pattern = re.compile(r'(?:CreateObject|New)\s*\(\s*["\']ADODB\.Connection["\']\s*\)', re.IGNORECASE)
            adodb_matches = adodb_pattern.finditer(vba_text)
            
            for match in adodb_matches:
                # Get surrounding context
                start_pos = max(0, match.start() - 200)
                end_pos = min(len(vba_text), match.end() + 200)
                context = vba_text[start_pos:end_pos]
                
                # Look for connection strings and queries in this context
                self._find_connection_strings_in_text(context, 'VBA ADODB Connection')
        
        except Exception as e:
            logger.warning(f"Error extracting from VBA: {str(e)}")
    
    def _process_power_query(self, query_data: str, file_name: str):
        """Process Power Query / M formula information."""
        try:
            root = ET.fromstring(query_data)
            
            # Extract formula text
            m_formula = None
            for elem in root.iter():
                if 'formula' in elem.tag.lower() and elem.text:
                    m_formula = elem.text
                    break
            
            if not m_formula:
                return
            
            # Create a connection entry
            query_name = re.search(r'query(\d+)\.xml', file_name)
            connection = {
                'Name': f"Power Query {query_name.group(1) if query_name else len(self.connections) + 1}",
                'Type': 'Power Query',
                'ConnectionString': '',
                'TargetDataSources': [],
                'QueryText': m_formula,
                'WorksheetName': '',
                'ComplexityScore': 0,
                'PurposeDescription': ''
            }
            
            # Extract possible data sources from M formula
            data_source_patterns = [
                (r'Source\s*=\s*(?:Csv|Excel|Sql|Oracle|Odbc|OleDb|Web|SharePoint)\.(?:Database|Document|File|Folder|Contents)\s*\(\s*"([^"]+)"', 1),
                (r'(?:Csv|Excel|Sql|Oracle|Odbc|OleDb|Web|SharePoint)\.(?:Database|Document|File|Folder|Contents)\s*\(\s*"([^"]+)"', 1),
                (r'(?:Server|Database|Source)\s*=\s*"([^"]+)"', 1)
            ]
            
            for pattern, group in data_source_patterns:
                matches = re.finditer(pattern, m_formula, re.IGNORECASE)
                for match in matches:
                    if match.group(group) and match.group(group) not in connection['TargetDataSources']:
                        connection['TargetDataSources'].append(match.group(group))
            
            # Determine connection type more specifically if possible
            if not connection['TargetDataSources']:
                if 'Csv.Document' in m_formula:
                    connection['Type'] = 'CSV'
                    connection['TargetDataSources'] = ['CSV File']
                elif 'Excel.Workbook' in m_formula:
                    connection['Type'] = 'Excel'
                    connection['TargetDataSources'] = ['Excel File']
                elif 'Sql.Database' in m_formula:
                    connection['Type'] = 'SQL Server'
                    connection['TargetDataSources'] = ['SQL Database']
                elif 'Web.Contents' in m_formula:
                    connection['Type'] = 'Web'
                    connection['TargetDataSources'] = ['Web Source']
            
            self.connections.append(connection)
        
        except Exception as e:
            logger.warning(f"Error processing Power Query {file_name}: {str(e)}")
    
    def _extract_connections_from_sheet(self, sheet_xml: str, sheet_name: str):
        """Extract potential connection information from a worksheet."""
        try:
            root = ET.fromstring(sheet_xml)
            namespace = root.tag.split('}')[0].strip('{') if '}' in root.tag else ''
            ns_prefix = '{' + namespace + '}' if namespace else ''
            
            # Look for data connections in the sheet
            query_table_elems = root.findall(f'.//{ns_prefix}queryTable') or root.findall('.//queryTable')
            
            for query_table in query_table_elems:
                connection = {
                    'Name': f"Sheet Connection {len(self.connections) + 1}",
                    'Type': 'Unknown',
                    'ConnectionString': '',
                    'TargetDataSources': [],
                    'QueryText': '',
                    'WorksheetName': sheet_name,
                    'ComplexityScore': 0,
                    'PurposeDescription': ''
                }
                
                # Get connection ID
                connection_id = query_table.get('connectionId', '')
                
                # Find query text
                query_elem = query_table.find(f'.//{ns_prefix}queryPr') or query_table.find('.//queryPr')
                if query_elem is not None:
                    connection_string = query_elem.get('connectionString', '')
                    if connection_string:
                        connection['ConnectionString'] = connection_string
                        connection['Type'] = self._determine_connection_type(connection_string)
                        data_sources = self._extract_data_sources(connection_string)
                        if data_sources:
                            connection['TargetDataSources'] = data_sources
                
                if connection['ConnectionString']:
                    self.connections.append(connection)
        
        except Exception as e:
            logger.warning(f"Error extracting from sheet {sheet_name}: {str(e)}")
    
    def _process_ole_connection_stream(self, data: bytes, stream_path: str):
        """Process an OLE stream that might contain connection information."""
        try:
            # Try to decode the binary data
            text_data = data.decode('latin-1', errors='ignore')
            
            # Look for connection strings
            self._find_connection_strings_in_text(text_data, stream_path)
            
            # Look for SQL queries
            self._find_sql_queries_in_text(text_data, stream_path)
        
        except Exception as e:
            logger.warning(f"Error processing OLE stream {stream_path}: {str(e)}")
    
    def _extract_connections_from_workbook_binary(self, workbook_data: bytes):
        """Extract connection information from binary workbook data."""
        # Convert to text and search for patterns
        text_data = workbook_data.decode('latin-1', errors='ignore')
        
        # Look for connection strings
        self._find_connection_strings_in_text(text_data, 'Workbook Binary')
        
        # Look for SQL queries
        self._find_sql_queries_in_text(text_data, 'Workbook Binary')
    
    def _find_connection_strings_in_text(self, text: str, source: str):
        """Find potential connection strings in text content."""
        # Common connection string patterns
        patterns = [
            # ODBC Connection Strings
            (r'(?:ODBC;|DSN=)([^;]+)(?:;|$).*?(?:DATABASE|DB)=([^;]+)', 'ODBC'),
            # OLEDB Connection Strings
            (r'Provider=([^;]+);.*?(?:Data Source|Server|Location)=([^;]+)', 'OLEDB'),
            # SQL Server Connection Strings
            (r'Server=([^;]+);.*?Database=([^;]+)', 'SQL Server'),
            # Generic connection strings
            (r'(?:Data Source|Server|DSN)=([^;]+);.*?(?:Initial Catalog|Database)=([^;]+)', 'Database'),
            # Connection strings with embedded credentials (look for patterns safely)
            (r'(?:User ID|UID)=([^;]+);.*?(?:Password|PWD)=([^;]+)', 'Database with Authentication'),
            # CSV file connections
            (r'(?:Text|CSV);.*?HDR=(?:Yes|No);.*?(?:DBQ|Source)=([^;]+)', 'CSV'),
            # Jet/ACE database connections
            (r'(?:Microsoft\.(?:Jet|ACE)\.OLEDB\.\d+\.\d+);.*?(?:Data Source|DBQ)=([^;]+)', 'Access')
        ]
        
        for pattern, conn_type in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE | re.DOTALL)
            
            for match in matches:
                # Extract full connection string (approximate)
                start_pos = max(0, match.start() - 50)
                end_pos = min(len(text), match.end() + 200)
                context = text[start_pos:end_pos]
                
                # Find beginning and end of connection string
                conn_start = context.find('Provider=')
                if conn_start < 0:
                    conn_start = context.find('ODBC;')
                if conn_start < 0:
                    conn_start = context.find('DSN=')
                if conn_start < 0:
                    conn_start = context.find('Driver=')
                if conn_start < 0:
                    conn_start = max(0, 50)  # Use default if no marker found
                
                # Find end of connection string (usually semicolon or quote)
                conn_end = context.find('";', conn_start)
                if conn_end < 0:
                    conn_end = context.find("';", conn_start)
                if conn_end < 0:
                    conn_end = min(len(context), conn_start + 200)  # Limit to reasonable length
                
                conn_string = context[conn_start:conn_end].strip()
                
                # Create connection record
                target_sources = []
                
                # Extract data source information
                if len(match.groups()) >= 2:
                    if match.group(2) and match.group(2) not in target_sources:
                        target_sources.append(match.group(2))
                if len(match.groups()) >= 1:
                    if match.group(1) and match.group(1) not in target_sources:
                        target_sources.append(match.group(1))
                
                connection = {
                    'Name': f"{source} Connection {len(self.connections) + 1}",
                    'Type': conn_type,
                    'ConnectionString': conn_string,
                    'TargetDataSources': target_sources,
                    'QueryText': '',
                    'WorksheetName': '',
                    'ComplexityScore': 0,
                    'PurposeDescription': ''
                }
                
                self.connections.append(connection)
    
    def _find_sql_queries_in_text(self, text: str, source: str):
        """Find potential SQL queries in text content."""
        # Common SQL query patterns
        sql_patterns = [
            r'SELECT\s+.+?\s+FROM\s+.+?(?:WHERE|GROUP\s+BY|ORDER\s+BY|;|$)',
            r'EXEC\s+\w+\s+(?:@\w+\s*=\s*[\'"]\w+[\'"]\s*,\s*)*(?:@\w+\s*=\s*[\'"]\w+[\'"]\s*)?',
            r'UPDATE\s+\w+\s+SET\s+.+?(?:WHERE|;|$)',
            r'INSERT\s+INTO\s+\w+\s*\(.+?\)\s*VALUES\s*\(.+?\)',
            r'DELETE\s+FROM\s+\w+\s+(?:WHERE|;|$)'
        ]
        
        for pattern in sql_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE | re.DOTALL)
            
            for match in matches:
                query_text = match.group(0).strip()
                
                # Limit to reasonable size and remove extra whitespace
                query_text = re.sub(r'\s+', ' ', query_text[:1000])
                
                # Extract table names (data sources)
                tables = set()
                table_pattern = r'FROM\s+(\w+)|JOIN\s+(\w+)'
                table_matches = re.finditer(table_pattern, query_text, re.IGNORECASE)
                
                for table_match in table_matches:
                    if table_match.group(1):
                        tables.add(table_match.group(1))
                    elif table_match.group(2):
                        tables.add(table_match.group(2))
                
                connection = {
                    'Name': f"{source} SQL Query {len(self.connections) + 1}",
                    'Type': 'SQL',
                    'ConnectionString': '',
                    'TargetDataSources': list(tables),
                    'QueryText': query_text,
                    'WorksheetName': '',
                    'ComplexityScore': 0,
                    'PurposeDescription': ''
                }
                
                self.connections.append(connection)
    
    def _looks_like_connection_string(self, text: str) -> bool:
        """Check if a text looks like a connection string."""
        if not text or len(text) < 10:
            return False
            
        # Look for common connection string keywords
        keywords = ['Provider=', 'Data Source=', 'Server=', 'DSN=', 'Driver=', 'Database=', 'Initial Catalog=']
        
        return any(keyword in text for keyword in keywords)
    
    def _looks_like_sql_query(self, text: str) -> bool:
        """Check if a text looks like an SQL query."""
        if not text or len(text) < 10:
            return False
            
        # Check for SQL keywords at the beginning of the text
        sql_starters = ['SELECT', 'INSERT', 'UPDATE', 'DELETE', 'CREATE', 'ALTER', 'DROP', 'EXEC', 'WITH']
        text_upper = text.upper().strip()
        
        return any(text_upper.startswith(keyword) for keyword in sql_starters)
    
    def _determine_connection_type(self, connection_string: str) -> str:
        """Determine the type of connection from a connection string."""
        if not connection_string:
            return 'Unknown'
            
        for conn_type, patterns in CONNECTION_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, connection_string, re.IGNORECASE):
                    return conn_type
        
        return 'Unknown'
    
    def _extract_data_sources(self, connection_string: str) -> List[str]:
        """Extract target data sources from a connection string."""
        sources = []
        
        # Try to extract database/catalog name
        db_match = re.search(r'(?:Initial Catalog|Database|DBQ)=([^;]+)', connection_string, re.IGNORECASE)
        if db_match and db_match.group(1):
            sources.append(db_match.group(1))
        
        # Try to extract server/data source
        server_match = re.search(r'(?:Data Source|Server|DSN)=([^;]+)', connection_string, re.IGNORECASE)
        if server_match and server_match.group(1):
            sources.append(server_match.group(1))
        
        # Try to extract file path for file-based connections
        file_match = re.search(r'(?:DBQ|Source|File)=([^;]+\.(?:csv|txt|xls|xlsx|mdb|accdb))', connection_string, re.IGNORECASE)
        if file_match and file_match.group(1):
            sources.append(file_match.group(1))
        
        return sources
    
    def _calculate_complexity(self, connection: Dict[str, Any]):
        """Calculate complexity score for a connection."""
        score = 0
        
        # Check connection string complexity
        if connection['ConnectionString']:
            # Add points for connection string length
            score += min(len(connection['ConnectionString']) / 50, 5)
            
            # Add points for specific connection features
            if 'Trusted_Connection' in connection['ConnectionString']:
                score += 1
            if 'Integrated Security' in connection['ConnectionString']:
                score += 1
            if 'User ID' in connection['ConnectionString'] or 'UID' in connection['ConnectionString']:
                score += 2
            if 'Password' in connection['ConnectionString'] or 'PWD' in connection['ConnectionString']:
                score += 2
        
        # Check query complexity
        if connection['QueryText']:
            query_text = connection['QueryText'].upper()
            
            # Count SQL keywords
            for keyword in SQL_KEYWORDS:
                if keyword in query_text:
                    score += 1
            
            # Count joins
            join_count = query_text.count('JOIN')
            score += join_count * 2
            
            # Check for subqueries
            if '(' in query_text and 'SELECT' in query_text[query_text.find('('):]:
                score += 5
            
            # Check for window functions
            if 'OVER (' in query_text:
                score += 3
            
            # Check for complex conditions
            if 'CASE' in query_text:
                score += 3
            
            # Check for complex grouping
            if 'GROUP BY' in query_text:
                score += 2
            
            # Check for CTEs
            if 'WITH ' in query_text and ' AS (' in query_text:
                score += 4
        
        # Add points for multiple data sources
        score += len(connection['TargetDataSources'])
        
        # Scale score from 0-10 with logarithmic scaling
        if score > 0:
            connection['ComplexityScore'] = min(round(1 + math.log(score + 1, 2), 1), 10)
        else:
            connection['ComplexityScore'] = 0
    
    def _infer_purpose(self, connection: Dict[str, Any]):
        """Infer purpose description based on connection details."""
        purpose = []
        
        # Check connection type
        if connection['Type'] != 'Unknown':
            purpose.append(f"{connection['Type']} connection")
        
        # Check target data sources
        if connection['TargetDataSources']:
            sources = ', '.join(connection['TargetDataSources'])
            purpose.append(f"accessing {sources}")
        
        # Check query
        if connection['QueryText']:
            query_text = connection['QueryText'].upper()
            
            # Check query type
            if query_text.startswith('SELECT'):
                purpose.append("data retrieval")
                
                # Look for aggregation functions
                agg_functions = ['COUNT', 'SUM', 'AVG', 'MIN', 'MAX']
                if any(func in query_text for func in agg_functions):
                    purpose.append("data aggregation")
                
                # Look for filtering
                if 'WHERE' in query_text:
                    purpose.append("filtered data")
                
                # Look for grouping
                if 'GROUP BY' in query_text:
                    purpose.append("grouped data")
            
            elif query_text.startswith('UPDATE'):
                purpose.append("data update")
            
            elif query_text.startswith('INSERT'):
                purpose.append("data insertion")
            
            elif query_text.startswith('DELETE'):
                purpose.append("data deletion")
        
        # Check complexity score
        if connection['ComplexityScore'] >= 7:
            purpose.append("complex analysis")
        elif connection['ComplexityScore'] >= 4:
            purpose.append("moderate analysis")
        
        # Check worksheet reference
        if connection['WorksheetName']:
            purpose.append(f"used in worksheet '{connection['WorksheetName']}'")
        
        # Combine all inferred purposes
        if purpose:
            connection['PurposeDescription'] = 'A ' + ' for '.join(purpose)
        else:
            connection['PurposeDescription'] = 'Unknown purpose'


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Analyze Excel files for database connections and CSV connections'
    )
    
    parser.add_argument(
        'file_path',
        type=str,
        help='Path to the Excel file to analyze'
    )
    
    parser.add_argument(
        '--output',
        type=str,
        default='connections.json',
        help='Output file path (default: connections.json)'
    )
    
    parser.add_argument(
        '--format',
        choices=['json', 'csv', 'txt'],
        default='json',
        help='Output format (default: json)'
    )
    
    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Enable verbose output'
    )
    
    return parser.parse_args()


def output_connections(connections, output_file, output_format):
    """Output connections to the specified format."""
    if output_format == 'json':
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(connections, f, indent=2)
    
    elif output_format == 'csv':
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'Name', 'Type', 'ConnectionString', 'TargetDataSources',
                'QueryText', 'WorksheetName', 'ComplexityScore', 'PurposeDescription'
            ])
            writer.writeheader()
            
            for conn in connections:
                # Convert list to string for CSV output
                conn_copy = conn.copy()
                if isinstance(conn_copy['TargetDataSources'], list):
                    conn_copy['TargetDataSources'] = ', '.join(conn_copy['TargetDataSources'])
                writer.writerow(conn_copy)
    
    elif output_format == 'txt':
        with open(output_file, 'w', encoding='utf-8') as f:
            for i, conn in enumerate(connections, 1):
                f.write(f"Connection {i}:\n")
                f.write(f"  Name: {conn['Name']}\n")
                f.write(f"  Type: {conn['Type']}\n")
                f.write(f"  Connection String: {conn['ConnectionString']}\n")
                f.write(f"  Target Data Sources: {', '.join(conn['TargetDataSources'])}\n")
                f.write(f"  Query Text: {conn['QueryText']}\n")
                f.write(f"  Worksheet Name: {conn['WorksheetName']}\n")
                f.write(f"  Complexity Score: {conn['ComplexityScore']}\n")
                f.write(f"  Purpose Description: {conn['PurposeDescription']}\n")
                f.write("\n")


def main():
    """Main function to run the analyzer."""
    args = parse_arguments()
    
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    try:
        # Check if file exists
        if not os.path.isfile(args.file_path):
            logger.error(f"File not found: {args.file_path}")
            sys.exit(1)
        
        # Analyze the file
        logger.info(f"Analyzing file: {args.file_path}")
        analyzer = ExcelConnectionAnalyzer(args.file_path)
        connections = analyzer.analyze()
        
        # Output connections
        logger.info(f"Found {len(connections)} connections")
        output_connections(connections, args.output, args.format)
        logger.info(f"Connections output to {args.output} in {args.format} format")
        
        # Print summary to console
        print(f"\nFound {len(connections)} connections in {args.file_path}")
        for i, conn in enumerate(connections, 1):
            print(f"\nConnection {i}:")
            print(f"  Name: {conn['Name']}")
            print(f"  Type: {conn['Type']}")
            print(f"  Target Data Sources: {', '.join(conn['TargetDataSources'])}")
            print(f"  Worksheet Name: {conn['WorksheetName'] if conn['WorksheetName'] else 'N/A'}")
            print(f"  Complexity Score: {conn['ComplexityScore']}")
            print(f"  Purpose: {conn['PurposeDescription']}")
        
        return 0
    
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
