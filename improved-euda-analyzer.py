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
    
    st.title("EUDA Analyzer Chatbot")
    st.markdown("""
    **Welcome!** Upload your Excel End User Developed Applications (EUDAs) for analysis.
    This tool will help you understand the complexity, risks, and potential improvements.
    """)
    
    # Initialize chat history
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": "Hello! I can analyze your Excel EUDAs. Upload an Excel file to begin."}
        ]
    
    # Initialize processing state
    if "processing" not in st.session_state:
        st.session_state.processing = False
    
    # Initialize completed flag
    if "analysis_complete" not in st.session_state:
        st.session_state.analysis_complete = False
    
    # Display chat messages in a container with fixed height and scrolling
    chat_container = st.container()
    with chat_container:
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
    
    # Create two columns for file upload and status
    col1, col2 = st.columns([3, 1])
    
    with col1:
        # File uploader widget
        uploaded_file = st.file_uploader("Upload an Excel file (.xlsx, .xls, .xlsm)", type=["xlsx", "xls", "xlsm"], key="file_uploader")
    
    with col2:
        if st.session_state.processing:
            st.warning("Processing file... Please wait")
        elif st.session_state.analysis_complete:
            st.success("Analysis complete! See results above.")
    
    # Process file if uploaded and not already processing
    if uploaded_file is not None and not st.session_state.processing:
        # Set processing flag
        st.session_state.processing = True
        st.session_state.analysis_complete = False
        
        # Force a rerun to show the processing indicator
        st.rerun()
    
    # If processing flag is set, do the analysis
    if st.session_state.processing and not st.session_state.analysis_complete:
        try:
            # Get the file from session state
            uploaded_file = st.session_state.file_uploader
            
            if uploaded_file:
                with st.spinner("Analyzing your Excel EUDA..."):
                    # Analyze the uploaded file
                    analysis_results = analyze_excel_euda(uploaded_file)
                    report = generate_report(analysis_results)
                    
                    # Add user message about the upload
                    st.session_state.messages.append({"role": "user", "content": f"I've uploaded {uploaded_file.name} for analysis."})
                    
                    # Add assistant message with the analysis report
                    st.session_state.messages.append({"role": "assistant", "content": report})
                    
                    # Set flags
                    st.session_state.processing = False
                    st.session_state.analysis_complete = True
                    
                    # Force a rerun to display the new messages
                    st.rerun()
        except Exception as e:
            # Handle any errors
            st.session_state.messages.append({"role": "assistant", "content": f"Error processing file: {str(e)}"})
            st.session_state.processing = False
            st.rerun()
    
    # User input area - always available
    user_input = st.chat_input("Ask a question about EUDA analysis or upload a file", disabled=st.session_state.processing)
    
    # Handle text input from user
    if user_input and not st.session_state.processing:
        # Add user message to chat history
        st.session_state.messages.append({"role": "user", "content": user_input})
        
        # Generate response based on user input
        if "analyze" in user_input.lower() and "file" in user_input.lower():
            response = "Please upload an Excel file using the file uploader above for analysis."
        elif "help" in user_input.lower() or "can you" in user_input.lower():
            response = """
            I can help you analyze Excel EUDA files. Here's what I can do:
            
            1. Analyze the structure and complexity of your Excel files
            2. Identify potential risk areas in your EUDA
            3. Provide recommendations based on best practices
            4. Detect formulas, data tables, and macros
            
            To begin, simply upload an Excel file using the file uploader above.
            """
        elif "euda" in user_input.lower() and ("what" in user_input.lower() or "mean" in user_input.lower()):
            response = """
            EUDA stands for End User Developed Application. These are typically spreadsheets or databases created by business users rather than IT professionals.
            
            Common characteristics of EUDAs include:
            - Created to solve specific business problems
            - Often developed incrementally over time
            - May contain complex formulas, macros, or data connections
            - Usually maintained by business users rather than IT
            
            While EUDAs provide flexibility, they can introduce risks like undocumented logic, single points of failure, and security issues.
            """
        elif "clear" in user_input.lower() or "reset" in user_input.lower():
            # Reset the session state
            st.session_state.messages = [
                {"role": "assistant", "content": "Chat history cleared. I can analyze your Excel EUDAs. Upload an Excel file to begin."}
            ]
            st.session_state.analysis_complete = False
            st.rerun()
            response = ""  # Not used but needed to avoid error
        else:
            response = "I'm here to help analyze Excel EUDA files. Please upload a file or ask specific questions about EUDA analysis."
        
        # Add assistant response to chat history
        if response:  # Only add if not empty (for clear/reset case)
            st.session_state.messages.append({"role": "assistant", "content": response})
        
        # Force a rerun to display the new messages
        st.rerun()

if __name__ == "__main__":
    main()
