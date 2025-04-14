def generate_report(excel_file, session):
    """
    Generate an HTML report for an analyzed Excel file that can be rendered in Streamlit.
    
    Args:
        excel_file (ExcelFile): Excel file record
        session (sqlalchemy.Session): Database session
        
    Returns:
        str: HTML report compatible with Streamlit's st.markdown() with unsafe_allow_html=True
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
    complexity_color = "green"
    if excel_file.complexity_score >= 7:
        complexity_class = "High"
        complexity_color = "red"
    elif excel_file.complexity_score >= 4:
        complexity_class = "Medium"
        complexity_color = "orange"
    
    # Begin HTML content - use simpler styling that works well in Streamlit
    html = f"""
    <style>
        .report-header {{
            background-color: #0078d4;
            color: white;
            padding: 10px 15px;
            border-radius: 5px;
            margin-bottom: 15px;
        }}
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 10px;
            margin-bottom: 20px;
        }}
        .summary-box {{
            background-color: #f5f5f5;
            border-radius: 5px;
            padding: 10px;
            text-align: center;
        }}
        .metric {{
            font-size: 20px;
            font-weight: bold;
            margin: 5px 0;
        }}
        .metric-label {{
            font-size: 12px;
            color: #666;
        }}
        .section {{
            background-color: #f9f9f9;
            border-left: 4px solid #0078d4;
            padding: 10px 15px;
            margin: 15px 0;
        }}
        .section-title {{
            color: #0078d4;
            margin-top: 0;
            margin-bottom: 10px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
            font-size: 14px;
        }}
        th, td {{
            padding: 8px 10px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background-color: #f2f2f2;
        }}
        .code {{
            font-family: monospace;
            background-color: #f7f7f7;
            padding: 4px;
            border-radius: 3px;
            font-size: 12px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            max-width: 300px;
        }}
        .remediation {{
            background-color: #f0f8ff;
            border-left: 4px solid #0078d4;
            padding: 10px 15px;
            margin: 10px 0;
        }}
    </style>

    <div class="report-header">
        <h2>Excel File Analysis Report: {excel_file.filename}</h2>
    </div>

    <div class="summary-grid">
        <div class="summary-box">
            <div class="metric-label">File Size</div>
            <div class="metric">{format_file_size(excel_file.file_size_kb * 1024)}</div>
        </div>
        <div class="summary-box">
            <div class="metric-label">Worksheets</div>
            <div class="metric">{excel_file.worksheet_count}</div>
        </div>
        <div class="summary-box">
            <div class="metric-label">Complexity</div>
            <div class="metric" style="color: {complexity_color};">{excel_file.complexity_score:.1f}/10</div>
        </div>
        <div class="summary-box">
            <div class="metric-label">Remediable</div>
            <div class="metric">{'Yes' if excel_file.can_be_remediated else 'No'}</div>
        </div>
    </div>

    <p><strong>File Path:</strong> {excel_file.file_path}</p>
    
    <div class="remediation">
        <h3 style="margin-top: 0;">Remediation Notes</h3>
        <p>{excel_file.remediation_notes}</p>
    </div>
    """

    # Worksheets Section
    if worksheets:
        html += f"""
    <div class="section">
        <h3 class="section-title">Worksheets ({len(worksheets)})</h3>
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
        <h3 class="section-title">Macros ({len(macros)})</h3>
        
        <h4>Top {len(complex_macros)} Complex Macros</h4>
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
        <h3 class="section-title">Formulas ({len(formulas)})</h3>
        
        <h4>Formula Types</h4>
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
        
        <h4>Most Complex Formulas</h4>
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
            truncated_formula = formula.formula_text
            if len(truncated_formula) > 50:
                truncated_formula = truncated_formula[:50] + "..."
                
            html += f"""
                <tr>
                    <td>{formula.worksheet_name}</td>
                    <td>{formula.cell_reference}</td>
                    <td><div class="code">{truncated_formula}</div></td>
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
        <h3 class="section-title">Database Connections ({len(connections)})</h3>
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
            query_display = conn.query_text[:50] + ('...' if len(conn.query_text or '') > 50 else '')
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

    # Footer
    html += f"""
    <div style="text-align: center; color: #666; font-size: 12px; margin-top: 20px;">
        Generated on: {time.strftime("%Y-%m-%d %H:%M:%S")}
    </div>
    """
    
    return html
