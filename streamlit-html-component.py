def generate_report(excel_file, session):
    """
    Generate an HTML report for an analyzed Excel file suitable for Streamlit rendering.
    
    Args:
        excel_file (ExcelFile): Excel file record
        session (sqlalchemy.Session): Database session
        
    Returns:
        str: HTML report as a string
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
    
    # Calculate complexity class and color
    complexity_color = "#107C10"  # Green for low
    if excel_file.complexity_score >= 7:
        complexity_color = "#E81123"  # Red for high
    elif excel_file.complexity_score >= 4:
        complexity_color = "#FF8C00"  # Orange for medium
    
    # Create HTML content for Streamlit
    html = f"""
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 100%; padding: 0; margin: 0;">
        <!-- Header -->
        <div style="background-color: #0078d4; color: white; padding: 15px; border-radius: 5px; margin-bottom: 15px;">
            <h2 style="margin: 0;">Excel Analysis: {excel_file.filename}</h2>
        </div>
        
        <!-- Summary Stats -->
        <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 15px;">
            <div style="flex: 1; min-width: 120px; background-color: #f5f5f5; padding: 10px; border-radius: 5px; text-align: center;">
                <div style="font-size: 12px; color: #666;">File Size</div>
                <div style="font-size: 18px; font-weight: bold;">{format_file_size(excel_file.file_size_kb * 1024)}</div>
            </div>
            <div style="flex: 1; min-width: 120px; background-color: #f5f5f5; padding: 10px; border-radius: 5px; text-align: center;">
                <div style="font-size: 12px; color: #666;">Worksheets</div>
                <div style="font-size: 18px; font-weight: bold;">{excel_file.worksheet_count}</div>
            </div>
            <div style="flex: 1; min-width: 120px; background-color: #f5f5f5; padding: 10px; border-radius: 5px; text-align: center;">
                <div style="font-size: 12px; color: #666;">Complexity</div>
                <div style="font-size: 18px; font-weight: bold; color: {complexity_color};">{excel_file.complexity_score:.1f}/10</div>
            </div>
            <div style="flex: 1; min-width: 120px; background-color: #f5f5f5; padding: 10px; border-radius: 5px; text-align: center;">
                <div style="font-size: 12px; color: #666;">Remediable</div>
                <div style="font-size: 18px; font-weight: bold;">{'Yes' if excel_file.can_be_remediated else 'No'}</div>
            </div>
        </div>
        
        <p><strong>File Path:</strong> {excel_file.file_path}</p>
        
        <!-- Remediation Notes -->
        <div style="background-color: #f0f8ff; border-left: 4px solid #0078d4; padding: 10px 15px; margin: 15px 0;">
            <h3 style="margin-top: 0; color: #0078d4;">Remediation Notes</h3>
            <p style="margin-bottom: 0;">{excel_file.remediation_notes}</p>
        </div>
    """

    # Worksheets Section
    if worksheets:
        html += """
        <!-- Worksheets Section -->
        <div style="background-color: #f9f9f9; border-left: 4px solid #0078d4; padding: 10px 15px; margin: 15px 0;">
            <h3 style="margin-top: 0; color: #0078d4;">Worksheets</h3>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <thead>
                        <tr>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Name</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Visibility</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Rows</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Columns</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Formulas</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Charts</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Tables</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        for ws in worksheets:
            html += f"""
                        <tr>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{ws.name}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{ws.visibility}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{ws.row_count}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{ws.column_count}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{ws.formula_count}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{ws.chart_count}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{ws.table_count}</td>
                        </tr>
            """
        
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """

    # Macros Section
    if macros:
        # Get top 5 most complex macros
        complex_macros = sorted(macros, key=lambda m: m.complexity_score, reverse=True)[:5]
        
        html += f"""
        <!-- Macros Section -->
        <div style="background-color: #f9f9f9; border-left: 4px solid #0078d4; padding: 10px 15px; margin: 15px 0;">
            <h3 style="margin-top: 0; color: #0078d4;">Macros ({len(macros)})</h3>
            <h4 style="margin-bottom: 10px;">Top {len(complex_macros)} Complex Macros</h4>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <thead>
                        <tr>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Name</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Module</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Lines</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Complexity</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Purpose</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        for macro in complex_macros:
            html += f"""
                        <tr>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{macro.name}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{macro.module_name}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{macro.line_count}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{macro.complexity_score:.2f}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{macro.purpose_description}</td>
                        </tr>
            """
        
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """

    # Formulas Section
    if formulas:
        html += f"""
        <!-- Formulas Section -->
        <div style="background-color: #f9f9f9; border-left: 4px solid #0078d4; padding: 10px 15px; margin: 15px 0;">
            <h3 style="margin-top: 0; color: #0078d4;">Formulas ({len(formulas)})</h3>
            <h4 style="margin-bottom: 10px;">Formula Types</h4>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <thead>
                        <tr>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Formula Type</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Count</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Percentage</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        total_formulas = len(formulas)
        for formula_type, count in formula_types.items():
            percentage = (count / total_formulas) * 100
            html += f"""
                        <tr>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{formula_type}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{count}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{percentage:.1f}%</td>
                        </tr>
            """
        
        html += """
                    </tbody>
                </table>
            </div>
            
            <h4 style="margin-bottom: 10px; margin-top: 15px;">Most Complex Formulas</h4>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <thead>
                        <tr>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Worksheet</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Cell</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Formula</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Type</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Complexity</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        # Get top 5 most complex formulas
        complex_formulas = sorted(formulas, key=lambda f: f.complexity_score, reverse=True)[:5]
        for formula in complex_formulas:
            truncated_formula = formula.formula_text
            if len(truncated_formula) > 40:
                truncated_formula = truncated_formula[:40] + "..."
                
            html += f"""
                        <tr>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{formula.worksheet_name}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{formula.cell_reference}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd; font-family: monospace; background-color: #f7f7f7;">{truncated_formula}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{formula.formula_type}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{formula.complexity_score:.2f}</td>
                        </tr>
            """
        
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """

    # Database Connections Section
    if connections:
        html += f"""
        <!-- Database Connections Section -->
        <div style="background-color: #f9f9f9; border-left: 4px solid #0078d4; padding: 10px 15px; margin: 15px 0;">
            <h3 style="margin-top: 0; color: #0078d4;">Database Connections ({len(connections)})</h3>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <thead>
                        <tr>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Type</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Target Database</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Worksheet</th>
                            <th style="text-align: left; padding: 8px; background-color: #f2f2f2; border-bottom: 1px solid #ddd;">Query</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        for conn in connections:
            # Truncate long query text for display
            query_display = conn.query_text[:40] + ('...' if len(conn.query_text or '') > 40 else '')
            html += f"""
                        <tr>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{conn.connection_type}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{conn.target_database or 'N/A'}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd;">{conn.worksheet_name or 'N/A'}</td>
                            <td style="text-align: left; padding: 8px; border-bottom: 1px solid #ddd; font-family: monospace; background-color: #f7f7f7;">{query_display}</td>
                        </tr>
            """
        
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """

    # Footer
    html += f"""
        <div style="text-align: center; color: #666; font-size: 12px; margin-top: 20px; margin-bottom: 10px;">
            Generated on: {time.strftime("%Y-%m-%d %H:%M:%S")}
        </div>
    </div>
    """
    
    return html

# Implementation in Streamlit chat
import streamlit as st
import components.v1 as components
from database.connection import get_db_session
from database.models import ExcelFile

def display_excel_report_in_chat(file_id):
    """
    Display an Excel file analysis report in Streamlit chat.
    
    Args:
        file_id (int): ID of the Excel file to display
    """
    # Get database session
    session = get_db_session()
    
    try:
        # Get Excel file from database
        excel_file = session.query(ExcelFile).filter(ExcelFile.id == file_id).first()
        
        if not excel_file:
            st.error(f"Excel file with ID {file_id} not found.")
            return
        
        # Generate HTML report
        html_report = generate_report(excel_file, session)
        
        # Display report in Streamlit using components.html
        # Set height to None to let it expand based on content
        components.html(html_report, height=None, scrolling=True)
        
    finally:
        # Close session
        session.close()

# Use in chat bot context
def process_message(user_input):
    """
    Process a user message in the chat interface.
    
    Args:
        user_input (str): User's input message
    """
    # For example purposes
    if "show report" in user_input.lower() and "file" in user_input.lower():
        # Extract file ID using regex
        import re
        match = re.search(r'file\s+(\d+)', user_input)
        
        if match:
            file_id = int(match.group(1))
            # Display report
            st.write(f"Here's the analysis report for file #{file_id}:")
            display_excel_report_in_chat(file_id)
        else:
            st.write("Please specify a file ID, for example: 'show report for file 123'")
    else:
        # Handle other message types
        st.write("How can I help you with Excel file analysis today?")

# Main Streamlit chat app
def main():
    st.title("Excel Analysis Chat")
    
    # Initialize chat history
    if "messages" not in st.session_state:
        st.session_state.messages = []
    
    # Display chat messages
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            if message["role"] == "assistant" and "html_content" in message:
                # Display HTML content using components.html
                components.html(message["html_content"], height=None, scrolling=True)
            else:
                st.write(message["content"])
    
    # Get user input
    user_input = st.chat_input("Ask about Excel files...")
    
    if user_input:
        # Add user message to chat history
        st.session_state.messages.append({"role": "user", "content": user_input})
        
        # Display user message
        with st.chat_message("user"):
            st.write(user_input)
        
        # Process the message
        with st.chat_message("assistant"):
            if "show report" in user_input.lower() and "file" in user_input.lower():
                # Extract file ID
                import re
                match = re.search(r'file\s+(\d+)', user_input)
                
                if match:
                    file_id = int(match.group(1))
                    
                    # Get database session
                    session = get_db_session()
                    
                    try:
                        # Get Excel file from database
                        excel_file = session.query(ExcelFile).filter(ExcelFile.id == file_id).first()
                        
                        if excel_file:
                            # Generate HTML report
                            response_text = f"Here's the analysis report for {excel_file.filename}:"
                            st.write(response_text)
                            
                            html_report = generate_report(excel_file, session)
                            
                            # Display HTML using components.html
                            components.html(html_report, height=None, scrolling=True)
                            
                            # Add to chat history
                            st.session_state.messages.append({
                                "role": "assistant", 
                                "content": response_text,
                                "html_content": html_report
                            })
                        else:
                            response_text = f"I couldn't find an Excel file with ID {file_id}."
                            st.write(response_text)
                            st.session_state.messages.append({"role": "assistant", "content": response_text})
                    finally:
                        # Close session
                        session.close()
                else:
                    response_text = "Please specify which Excel file report you'd like to see. Example: 'show report for file 123'"
                    st.write(response_text)
                    st.session_state.messages.append({"role": "assistant", "content": response_text})
            else:
                # Handle other types of messages
                response_text = "How can I help you analyze Excel files today?"
                st.write(response_text)
                st.session_state.messages.append({"role": "assistant", "content": response_text})

if __name__ == "__main__":
    main()
