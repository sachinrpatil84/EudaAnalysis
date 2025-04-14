import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime
import re
import time

def analyze_excel_euda(file):
    """
    Analyze an Excel EUDA file and return analysis results.
    """
    try:
        # Load the Excel file with all sheets
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        
        # Initialize analysis results
        analysis = {
            "file_name": file.name,
            "sheet_count": len(sheet_names),
            "formulas_count": 0,
            "named_ranges": [],
            "macros_detected": False,
            "data_tables": [],
            "sheets_analysis": [],
            "complexity_score": 0,
            "risk_areas": []
        }
        
        # Check if file contains macros by looking at extension
        if file.name.lower().endswith(('.xlsm', '.xls')):
            analysis["macros_detected"] = True
            analysis["risk_areas"].append("Contains macros which should be reviewed for security")
        
        # Analyze each sheet
        total_cells = 0
        formula_cells = 0
        
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                rows, cols = df.shape
                total_cells += rows * cols
                
                # Extract cell details including formulas 
                sheet_formulas = 0
                for col in df.columns:
                    # Check for potential formula indicators in cell values
                    if df[col].dtype == 'object':
                        potential_formulas = df[col].astype(str).str.contains(r'^\s*=', regex=True).sum()
                        sheet_formulas += potential_formulas
                
                # Detect potential data tables
                if rows > 5 and cols > 3:
                    analysis["data_tables"].append({
                        "sheet": sheet_name,
                        "dimensions": f"{rows}x{cols}"
                    })
                
                # Record sheet analysis
                analysis["sheets_analysis"].append({
                    "name": sheet_name,
                    "rows": rows,
                    "columns": cols,
                    "formula_count": sheet_formulas
                })
                
                # Update total formula count
                analysis["formulas_count"] += sheet_formulas
                formula_cells += sheet_formulas
            except Exception as sheet_error:
                analysis["sheets_analysis"].append({
                    "name": sheet_name,
                    "error": str(sheet_error)
                })
                analysis["risk_areas"].append(f"Error analyzing sheet {sheet_name}: {str(sheet_error)}")
        
        # Calculate complexity score (simple heuristic)
        formula_ratio = formula_cells / total_cells if total_cells > 0 else 0
        complexity_base = min(10, analysis["sheet_count"]) + min(10, len(analysis["data_tables"]))
        formula_complexity = min(30, analysis["formulas_count"] / 10)
        analysis["complexity_score"] = round(complexity_base + formula_complexity + (formula_ratio * 50))
        
        # Add risk assessments based on analysis
        if analysis["complexity_score"] > 70:
            analysis["risk_areas"].append("High complexity suggests difficult maintenance")
        
        if analysis["formulas_count"] > 100:
            analysis["risk_areas"].append("Large number of formulas increases error risk")
            
        if analysis["sheet_count"] > 10:
            analysis["risk_areas"].append("Large number of sheets may indicate poor organization")
        
        return analysis
    
    except Exception as e:
        return {"error": str(e)}

def generate_report(analysis):
    """
    Generate a formatted report from the analysis results.
    """
    if "error" in analysis:
        return f"⚠️ Error analyzing file: {analysis['error']}"
    
    report = f"""
    # EUDA Analysis Report: {analysis['file_name']}
    
    ## Summary
    - **Analyzed on:** {datetime.now().strftime('%Y-%m-%d %H:%M')}
    - **Number of sheets:** {analysis['sheet_count']}
    - **Total formulas detected:** {analysis['formulas_count']}
    - **Data tables found:** {len(analysis['data_tables'])}
    - **Macros present:** {'Yes' if analysis['macros_detected'] else 'No'}
    - **Complexity score:** {analysis['complexity_score']}/100
    
    ## Sheet Details
    """
    
    for sheet in analysis['sheets_analysis']:
        report += f"\n### {sheet['name']}\n"
        if "error" in sheet:
            report += f"- Error: {sheet['error']}\n"
        else:
            report += f"- Dimensions: {sheet['rows']} rows × {sheet['columns']} columns\n"
            report += f"- Formulas: {sheet['formula_count']}\n"
    
    if analysis['data_tables']:
        report += "\n## Detected Data Tables\n"
        for table in analysis['data_tables']:
            report += f"- Sheet '{table['sheet']}': {table['dimensions']}\n"
    
    if analysis['risk_areas']:
        report += "\n## Risk Assessment\n"
        for risk in analysis['risk_areas']:
            report += f"- ⚠️ {risk}\n"
    
    report += f"""
    ## Recommendations
    
    Based on the complexity score of {analysis['complexity_score']}/100, this EUDA is 
    rated as **{'High' if analysis['complexity_score'] > 70 else 'Medium' if analysis['complexity_score'] > 40 else 'Low'} Complexity**.
    
    """
    
    if analysis['complexity_score'] > 70:
        report += """
    **Suggested actions:**
    - Consider migrating this EUDA to a more structured application platform
    - Implement comprehensive testing regime before any changes
    - Document all business rules and calculations
    """
    elif analysis['complexity_score'] > 40:
        report += """
    **Suggested actions:**
    - Implement version control for this spreadsheet
    - Create documentation for maintenance
    - Review formula logic for potential simplification
    """
    else:
        report += """
    **Suggested actions:**
    - Basic documentation recommended
    - Consider implementing cell protection to prevent accidental changes
    - Regular reviews to ensure continued fitness for purpose
    """
    
    return report

def main():
    st.set_page_config(page_title="EUDA Analyzer Chatbot", layout="wide")
    
    # Initialize session state variables
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": "Hello! I can analyze your Excel EUDAs. Upload an Excel file to begin."}
        ]
    
    if "uploaded_file_processed" not in st.session_state:
        st.session_state.uploaded_file_processed = False
    
    # Display header
    st.title("EUDA Analyzer Chatbot")
    st.markdown("""
    **Welcome!** Upload your Excel End User Developed Applications (EUDAs) for analysis.
    This tool will help you understand the complexity, risks, and potential improvements.
    """)
    
    # Create a tab layout for better organization
    tab1, tab2 = st.tabs(["Chat Interface", "File Upload"])
    
    with tab2:
        st.subheader("Upload Excel EUDA File")
        
        # File uploader
        uploaded_file = st.file_uploader("Select an Excel file (.xlsx, .xls, .xlsm)", 
                                         type=["xlsx", "xls", "xlsm"],
                                         key="file_uploader")
        
        # Process button to give user control over when analysis runs
        if uploaded_file is not None:
            if st.button("Analyze File"):
                with st.spinner("Analyzing your Excel EUDA..."):
                    try:
                        # Save filename to session state before processing
                        file_name = uploaded_file.name
                        
                        # Process the file
                        analysis_results = analyze_excel_euda(uploaded_file)
                        report = generate_report(analysis_results)
                        
                        # Add messages to chat history
                        st.session_state.messages.append({"role": "user", 
                                                         "content": f"I've uploaded {file_name} for analysis."})
                        st.session_state.messages.append({"role": "assistant", 
                                                         "content": report})
                        
                        # Set flag to prevent reprocessing
                        st.session_state.uploaded_file_processed = True
                        
                        # Success message
                        st.success("Analysis complete! View results in the Chat Interface tab.")
                    except Exception as e:
                        st.error(f"Error analyzing file: {str(e)}")
    
    with tab1:
        # Chat message display with fixed height container and scrollbar
        st.subheader("Chat History")
        
        # Create a container with fixed height for chat messages
        chat_container = st.container()
        
        # Force the container to have a scrollbar by setting its height with custom CSS
        st.markdown("""
        <style>
        .chat-container {
            height: 400px;
            overflow-y: auto;
            border: 1px solid #ddd;
            padding: 10px;
            border-radius: 5px;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Chat message display within a scrollable div
        chat_html = '<div class="chat-container">'
        
        for message in st.session_state.messages:
            role_class = "assistant" if message["role"] == "assistant" else "user"
            chat_html += f'<div class="{role_class}-message" style="margin-bottom: 15px; padding: 10px; border-radius: 5px; background-color: {"#f0f7ff" if role_class == "assistant" else "#e6f7e6"};">'
            chat_html += f'<strong>{message["role"].capitalize()}</strong><br>'
            chat_html += f'{message["content"].replace("#", "").replace("**", "<b>").replace("</b>", "</b>")}'
            chat_html += '</div>'
        
        chat_html += '</div>'
        
        with chat_container:
            st.markdown(chat_html, unsafe_allow_html=True)
            
            # Always scroll to the bottom of the chat (using Javascript)
            st.markdown("""
            <script>
                // Scroll chat container to bottom
                const chatContainer = document.querySelector('.chat-container');
                if (chatContainer) {
                    chatContainer.scrollTop = chatContainer.scrollHeight;
                }
            </script>
            """, unsafe_allow_html=True)
        
        # User input area for chat
        user_input = st.text_input("Type your message here:", key="user_input")
        
        if user_input:
            # Add user message to chat history
            st.session_state.messages.append({"role": "user", "content": user_input})
            
            # Generate response based on user input
            if "analyze" in user_input.lower() and "file" in user_input.lower():
                response = "Please go to the 'File Upload' tab to upload and analyze an Excel file."
            elif "clear" in user_input.lower() or "reset" in user_input.lower():
                # Reset the chat
                st.session_state.messages = [
                    {"role": "assistant", "content": "Chat history cleared. I can analyze your Excel EUDAs. Upload an Excel file to begin."}
                ]
                st.session_state.uploaded_file_processed = False
                st.rerun()
            elif "help" in user_input.lower():
                response = """
                I can help you analyze Excel EUDA files. Here's what I can do:
                
                1. Analyze the structure and complexity of your Excel files
                2. Identify potential risk areas in your EUDA
                3. Provide recommendations based on best practices
                
                To begin, go to the 'File Upload' tab and upload an Excel file.
                """
            elif "euda" in user_input.lower() and ("what" in user_input.lower() or "mean" in user_input.lower()):
                response = """
                EUDA stands for End User Developed Application. These are typically spreadsheets or databases created by business users rather than IT professionals.
                
                Common characteristics of EUDAs include:
                - Created to solve specific business problems
                - Often developed incrementally over time
                - May contain complex formulas, macros, or data connections
                - Usually maintained by business users rather than IT
                """
            else:
                response = "I'm here to help analyze Excel EUDA files. You can ask me questions about EUDA analysis or go to the 'File Upload' tab to upload a file."
            
            # Add assistant response to chat history
            st.session_state.messages.append({"role": "assistant", "content": response})
            
            # Force a refresh to show the new messages
            st.rerun()

if __name__ == "__main__":
    main()
