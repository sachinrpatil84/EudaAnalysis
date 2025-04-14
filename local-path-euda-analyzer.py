import streamlit as st
import pandas as pd
import os
from datetime import datetime
import re

def analyze_excel_euda(file_path):
    """
    Analyze an Excel EUDA file from local path and return analysis results.
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            return {"error": f"File not found: {file_path}"}
        
        # Load the Excel file with all sheets
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        # Get just the filename without path
        file_name = os.path.basename(file_path)
        
        # Initialize analysis results
        analysis = {
            "file_name": file_name,
            "file_path": file_path,
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
        if file_path.lower().endswith(('.xlsm', '.xls')):
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
    - **File path:** {analysis['file_path']}
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
    
    # Initialize chat history
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": "Hello! I can analyze your Excel EUDAs. Enter the file path to begin."}
        ]
    
    st.title("EUDA Analyzer Chatbot")
    st.markdown("""
    **Welcome!** I can analyze Excel End User Developed Applications (EUDAs) from your local disk.
    This tool will help you understand the complexity, risks, and potential improvements.
    
    Simply type the full path to your Excel file, and I'll analyze it for you.
    """)
    
    # Display chat messages
    chat_container = st.container()
    with chat_container:
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
    
    # User input area for chat
    user_input = st.chat_input("Type the path to your Excel file or ask a question")
    
    if user_input:
        # Add user message to chat history
        st.session_state.messages.append({"role": "user", "content": user_input})
        
        # Check if input looks like a file path with an Excel extension
        if os.path.splitext(user_input.strip())[1].lower() in ['.xlsx', '.xls', '.xlsm'] or "analyze" in user_input.lower() and ".xls" in user_input.lower():
            # Extract file path from the message if needed
            file_path = user_input.strip()
            if not os.path.splitext(file_path)[1]:
                # If no extension, look for the last word that might be a path
                parts = user_input.split()
                for part in parts:
                    if os.path.splitext(part)[1].lower() in ['.xlsx', '.xls', '.xlsm']:
                        file_path = part
                        break
            
            with st.spinner(f"Analyzing Excel file: {file_path}"):
                try:
                    # Analyze the file
                    analysis_results = analyze_excel_euda(file_path)
                    report = generate_report(analysis_results)
                    
                    # Add assistant message with the analysis report
                    st.session_state.messages.append({"role": "assistant", "content": report})
                except Exception as e:
                    # Handle any errors
                    error_msg = f"Error processing file: {str(e)}"
                    st.session_state.messages.append({"role": "assistant", "content": error_msg})
        
        # Handle other types of queries
        elif "clear" in user_input.lower() or "reset" in user_input.lower():
            # Reset the chat
            st.session_state.messages = [
                {"role": "assistant", "content": "Chat history cleared. I can analyze your Excel EUDAs. Enter the file path to begin."}
            ]
            st.rerun()
        elif "help" in user_input.lower() or "how" in user_input.lower():
            response = """
            **How to use the EUDA Analyzer:**
            
            1. Type the complete file path to your Excel file
               Example: `C:/Users/YourName/Documents/example.xlsx`
               
            2. I'll analyze the file and provide a detailed report
            
            3. You can ask follow-up questions about EUDA best practices
            
            4. Type "clear" or "reset" to start a new conversation
            """
            st.session_state.messages.append({"role": "assistant", "content": response})
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
            st.session_state.messages.append({"role": "assistant", "content": response})
        else:
            response = "I'm here to help analyze Excel EUDA files. Please provide the full path to your Excel file, or ask a specific question about EUDA analysis."
            st.session_state.messages.append({"role": "assistant", "content": response})
        
        # Force a rerun to display the new messages
        st.rerun()

if __name__ == "__main__":
    main()
