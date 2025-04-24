import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime
import uuid
import re

# Configure upload directory
UPLOAD_FOLDER = "uploaded_eudas"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def save_uploaded_file(uploaded_file):
    """
    Save the uploaded file to the server's designated folder
    and return the path where it was saved
    """
    # Create a unique filename to avoid collisions
    file_extension = os.path.splitext(uploaded_file.name)[1]
    unique_filename = f"{uuid.uuid4()}{file_extension}"
    file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
    
    # Save uploaded file to disk
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    return file_path, uploaded_file.name

def analyze_excel_euda(file_path, original_filename):
    """
    Analyze an Excel EUDA file and return analysis results.
    """
    try:
        # Load the Excel file with all sheets
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        # Initialize analysis results
        analysis = {
            "file_name": original_filename,
            "server_path": file_path,
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
        if original_filename.lower().endswith(('.xlsm', '.xls')):
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

def cleanup_old_files():
    """Remove files older than 24 hours from the upload folder"""
    current_time = datetime.now().timestamp()
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        # Check when the file was last modified
        file_mod_time = os.path.getmtime(file_path)
        # Remove if older than 24 hours (86400 seconds)
        if current_time - file_mod_time > 86400:
            try:
                os.remove(file_path)
            except:
                pass

def main():
    st.set_page_config(page_title="EUDA Analyzer Chatbot", layout="wide")
    
    # Clean up old files on startup
    cleanup_old_files()
    
    # Initialize chat history
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": "Hello! I can analyze your Excel EUDAs. Upload your file to begin."}
        ]
    
    if "processing_file" not in st.session_state:
        st.session_state.processing_file = False
    
    # Create a two-column layout
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.title("EUDA Analyzer Chatbot")
    
    with col2:
        # File upload widget in the sidebar for better UI separation
        uploaded_file = st.file_uploader("Upload an Excel file", 
                                        type=["xlsx", "xls", "xlsm"],
                                        key="file_uploader",
                                        help="Select an Excel file to analyze")
    
    # Display chat messages
    chat_container = st.container()
    with chat_container:
        # Add some custom CSS to improve chat display
        st.markdown("""
        <style>
        .stChatMessage {
            padding: 15px;
            margin-bottom: 10px;
            border-radius: 8px;
        }
        
        .stTextInput > div > div > input {
            font-size: 16px;
            line-height: 1.6;
        }
        </style>
        """, unsafe_allow_html=True)
        
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
    
    # Create a separate container for the process button
    button_container = st.container()
    
    # Process the uploaded file if available and button is clicked
    if uploaded_file is not None and not st.session_state.processing_file:
        with button_container:
            if st.button("Process Excel File"):
                # Set processing flag
                st.session_state.processing_file = True
                
                with st.spinner(f"Uploading and analyzing {uploaded_file.name}..."):
                    try:
                        # Save the file to the server
                        server_path, original_filename = save_uploaded_file(uploaded_file)
                        
                        # Add user message about the upload
                        st.session_state.messages.append(
                            {"role": "user", 
                             "content": f"I've uploaded {original_filename} for analysis."}
                        )
                        
                        # Analyze the file
                        analysis_results = analyze_excel_euda(server_path, original_filename)
                        report = generate_report(analysis_results)
                        
                        # Add assistant message with the analysis report
                        st.session_state.messages.append(
                            {"role": "assistant", 
                             "content": report}
                        )
                    except Exception as e:
                        # Handle any errors
                        error_msg = f"Error processing file: {str(e)}"
                        st.session_state.messages.append(
                            {"role": "assistant", 
                             "content": error_msg}
                        )
                    finally:
                        # Reset processing flag
                        st.session_state.processing_file = False
                        
                        # Rerun to update UI
                        st.rerun()
    
    # User input area for chat
    user_input = st.chat_input("Ask a question or provide feedback", 
                              key="user_input",
                              disabled=st.session_state.processing_file)
    
    if user_input:
        # Add user message to chat history
        st.session_state.messages.append({"role": "user", "content": user_input})
        
        # Generate response based on user input
        if "clear" in user_input.lower() or "reset" in user_input.lower():
            # Reset the chat
            st.session_state.messages = [
                {"role": "assistant", "content": "Chat history cleared. I can analyze your Excel EUDAs. Upload your file to begin."}
            ]
            st.rerun()
        elif "help" in user_input.lower() or "how" in user_input.lower():
            response = """
            **How to use the EUDA Analyzer:**
            
            1. Upload your Excel file using the file uploader in the top-right
            2. Click the "Process Excel File" button that appears
            3. Wait for the analysis to complete
            4. Review the report in this chat window
            5. Ask follow-up questions as needed
            
            The system analyzes End User Developed Applications (EUDAs) in Excel and provides insights on complexity, risks, and recommendations.
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
            response = "I'm here to help analyze Excel EUDA files. Please upload a file using the uploader in the top-right corner, or ask a specific question about EUDA analysis."
            st.session_state.messages.append({"role": "assistant", "content": response})
        
        # Force a rerun to display the new messages
        st.rerun()

if __name__ == "__main__":
    main()
