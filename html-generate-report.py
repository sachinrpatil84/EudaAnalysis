def generate_report(excel_file, session):
    """
    Generate an HTML report for an analyzed Excel file.
    
    Args:
        excel_file (ExcelFile): Excel file record
        session (sqlalchemy.Session): Database session
        
    Returns:
        str: HTML report
    """
    # Get related data
    macros = session.query(Macro).filter(Macro.excel_file_id == excel_file.id).all()
    formulas = session.query(Formula).filter(Formula.excel_file_id == excel_file.id).all()
    connections = session.query(DatabaseConnection).filter(DatabaseConnection.excel_file_id == excel_file.id).all()
    worksheets = session.query(Worksheet).filter(Worksheet.excel_file_id == excel_file.id).all()
    
    # Group formulas by type
    formula_types = {}
    for formula in formulas:
        if formula.formula_type in formula_types:
            formula_types[formula.formula_type] += 1
        else:
            formula_types[formula.formula_type] = 1
    
    # Sort by frequency
    formula_types = {k: v for k, v in sorted(formula_types.items(), key=lambda item: item[1], reverse=True)}
    
    # Calculate complexity class
    complexity_class = "Low"
    if excel_file.complexity_score >= 7:
        complexity_class = "High"
    elif excel_file.complexity_score >= 4:
        complexity_class = "Medium"
    
    # Begin HTML content
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Analysis Report: {excel_file.filename}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }}
        .header {{
            background-color: #0078d4;
            color: white;
            padding: 20px;
            border-radius: 5px 5px 0 0;
            margin-bottom: 0;
        }}
        .summary-container {{
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            margin-bottom: 30px;
        }}
        .summary-box {{
            background-color: #f5f5f5;
            border-radius: 5px;
            padding: 15px;
            flex: 1;
            min-width: 200px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .metric {{
            font-size: 24px;
            font-weight: bold;
            margin: 10px 0;
        }}
        .metric-label {{
            font-size: 14px;
            color: #666;
        }}
        .complexity-low {{
            color: #107C10;
        }}
        .complexity-medium {{
            color: #FF8C00;
        }}
        .complexity-high {{
            color: #E81123;
        }}
        .section {{
            background-color: white;
            border-radius: 5px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .section-title {{
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
            margin-top: 0;
            color: #0078d4;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}
        th, td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background-color: #f2f2f2;
            font-weight: 600;
        }}
        tr:hover {{
            background-color: #f5f5f5;
        }}
        .chart {{
            margin: 20px 0;
            height: 300px;
        }}
        .remediation {{
            background-color: #f0f8ff;
            border-left: 5px solid #0078d4;
            padding: 15px;
            margin: 20px 0;
        }}
        .remediation-title {{
            font-weight: 600;
            margin-top: 0;
        }}
        .code {{
            font-family: Consolas, Monaco, 'Courier New', monospace;
            background-color: #f7f7f7;
            padding: 8px;
            border-radius: 3px;
            font-size: 90%;
            overflow-x: auto;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Excel File Analysis Report</h1>
        <h2>{excel_file.filename}</h2>
    </div>

    <div class="section">
        <div class="summary-container">
            <div class="summary-box">
                <div class="metric-label">File Size</div>
                <div class="metric">{format_file_size(excel_file.file_size_kb * 1024)}</div>
            </div>
            <div class="summary-box">
                <div class="metric-label">Worksheets</div>
                <div class="metric">{excel_file.worksheet_count}</div>
            </div>
            <div class="summary-box">
                <div class="metric-label">Complexity Score</div>
                <div class="metric complexity-{complexity_class.lower()}">{excel_file.complexity_score:.1f}/10</div>
            </div>
            <div class="summary-box">
                <div class="metric-label">Can Be Remediated</div>
                <div class="metric">{'Yes' if excel_file.can_be_remediated else 'No'}</div>
            </div>
        </div>

        <p><strong>File Path:</strong> {excel_file.file_path}</p>
        
        <div class="remediation">
            <h3 class="remediation-title">Remediation Notes</h3>
            <p>{excel_file.remediation_notes}</p>
        </div>
    </div>
    """

    # Worksheets Section
    if worksheets:
        html += f"""
    <div class="section">
        <h2 class="section-title">Worksheets ({len(worksheets)})</h2>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Visibility</th>
                    <th>Rows</th>
                    <th>Columns</th>
                    <th>Formulas</th>
                    <th>Charts</th>
                    <th>Tables</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for ws in worksheets:
            html += f"""
                <tr>
                    <td>{ws.name}</td>
                    <td>{ws.visibility}</td>
                    <td>{ws.row_count}</td>
                    <td>{ws.column_count}</td>
                    <td>{ws.formula_count}</td>
                    <td>{ws.chart_count}</td>
                    <td>{ws.table_count}</td>
                </tr>
            """
        
        html += """
            </tbody>
        </table>
    </div>
        """

    # Macros Section
    if macros:
        # Get top 5 most complex macros
        complex_macros = sorted(macros, key=lambda m: m.complexity_score, reverse=True)[:5]
        
        html += f"""
    <div class="section">
        <h2 class="section-title">Macros ({len(macros)})</h2>
        
        <h3>Top {len(complex_macros)} Most Complex Macros</h3>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Module</th>
                    <th>Lines</th>
                    <th>Complexity</th>
                    <th>Purpose</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for macro in complex_macros:
            html += f"""
                <tr>
                    <td>{macro.name}</td>
                    <td>{macro.module_name}</td>
                    <td>{macro.line_count}</td>
                    <td>{macro.complexity_score:.2f}</td>
                    <td>{macro.purpose_description}</td>
                </tr>
            """
        
        html += """
            </tbody>
        </table>
    </div>
        """

    # Formulas Section
    if formulas:
        html += f"""
    <div class="section">
        <h2 class="section-title">Formulas ({len(formulas)})</h2>
        
        <h3>Formula Types Distribution</h3>
        <table>
            <thead>
                <tr>
                    <th>Formula Type</th>
                    <th>Count</th>
                    <th>Percentage</th>
                </tr>
            </thead>
            <tbody>
        """
        
        total_formulas = len(formulas)
        for formula_type, count in formula_types.items():
            percentage = (count / total_formulas) * 100
            html += f"""
                <tr>
                    <td>{formula_type}</td>
                    <td>{count}</td>
                    <td>{percentage:.1f}%</td>
                </tr>
            """
        
        html += """
            </tbody>
        </table>
        
        <h3>Most Complex Formulas</h3>
        <table>
            <thead>
                <tr>
                    <th>Worksheet</th>
                    <th>Cell</th>
                    <th>Formula</th>
                    <th>Type</th>
                    <th>Complexity</th>
                </tr>
            </thead>
            <tbody>
        """
        
        # Get top 5 most complex formulas
        complex_formulas = sorted(formulas, key=lambda f: f.complexity_score, reverse=True)[:5]
        for formula in complex_formulas:
            html += f"""
                <tr>
                    <td>{formula.worksheet_name}</td>
                    <td>{formula.cell_reference}</td>
                    <td><div class="code">{formula.formula_text}</div></td>
                    <td>{formula.formula_type}</td>
                    <td>{formula.complexity_score:.2f}</td>
                </tr>
            """
        
        html += """
            </tbody>
        </table>
    </div>
        """

    # Database Connections Section
    if connections:
        html += f"""
    <div class="section">
        <h2 class="section-title">Database Connections ({len(connections)})</h2>
        <table>
            <thead>
                <tr>
                    <th>Type</th>
                    <th>Target Database</th>
                    <th>Worksheet</th>
                    <th>Query</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for conn in connections:
            # Truncate long query text for display
            query_display = conn.query_text[:100] + ('...' if len(conn.query_text or '') > 100 else '')
            html += f"""
                <tr>
                    <td>{conn.connection_type}</td>
                    <td>{conn.target_database or 'N/A'}</td>
                    <td>{conn.worksheet_name or 'N/A'}</td>
                    <td><div class="code">{query_display}</div></td>
                </tr>
            """
        
        html += """
            </tbody>
        </table>
    </div>
        """

    # Close HTML document
    html += """
    <div class="section">
        <p style="text-align: center; color: #666; font-size: 12px;">
            Generated on: """ + time.strftime("%Y-%m-%d %H:%M:%S") + """<br>
            Excel File Analysis Tool
        </p>
    </div>
</body>
</html>
    """
    
    return html
