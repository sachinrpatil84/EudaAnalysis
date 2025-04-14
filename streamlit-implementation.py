import streamlit as st
from database.connection import get_db_session
from database.models import ExcelFile

def display_excel_report(file_id):
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
        
        # Display report in Streamlit using markdown with HTML
        st.markdown(html_report, unsafe_allow_html=True)
        
    finally:
        # Close session
        session.close()

# Example usage in a Streamlit chatbot
def handle_report_request(message):
    """
    Handle a user request to view an Excel file report.
    
    Args:
        message (str): User message
    """
    # Example: Extract file ID from message (in practice, you might use NLP or regex)
    # For demo, assume the message is "show report for file 123"
    import re
    match = re.search(r'file\s+(\d+)', message)
    
    if match:
        file_id = int(match.group(1))
        display_excel_report(file_id)
    else:
        st.markdown("Please specify which Excel file report you'd like to see. Example: 'show report for file 123'")

# In your Streamlit app
if "messages" not in st.session_state:
    st.session_state.messages = []

# Display chat history
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        if message["role"] == "assistant" and "html_report" in message:
            # Display HTML report
            st.markdown(message["html_report"], unsafe_allow_html=True)
        else:
            # Display regular message
            st.markdown(message["content"])

# Get user input
user_input = st.chat_input("Ask about Excel files...")

# Process user input
if user_input:
    # Add user message to chat history
    st.session_state.messages.append({"role": "user", "content": user_input})
    
    # Display user message
    with st.chat_message("user"):
        st.markdown(user_input)
    
    # Check if user is requesting a report
    if "report" in user_input.lower() and "file" in user_input.lower():
        # Get database session
        session = get_db_session()
        
        try:
            # Extract file ID (simplified example)
            import re
            match = re.search(r'file\s+(\d+)', user_input)
            
            if match:
                file_id = int(match.group(1))
                excel_file = session.query(ExcelFile).filter(ExcelFile.id == file_id).first()
                
                if excel_file:
                    # Generate report
                    html_report = generate_report(excel_file, session)
                    
                    # Add assistant response with HTML report
                    response = f"Here's the analysis report for {excel_file.filename}:"
                    st.session_state.messages.append({
                        "role": "assistant", 
                        "content": response,
                        "html_report": html_report
                    })
                    
                    # Display assistant response
                    with st.chat_message("assistant"):
                        st.markdown(response)
                        st.markdown(html_report, unsafe_allow_html=True)
                else:
                    # File not found
                    response = f"I couldn't find an Excel file with ID {file_id}."
                    st.session_state.messages.append({"role": "assistant", "content": response})
                    
                    with st.chat_message("assistant"):
                        st.markdown(response)
            else:
                # Invalid request format
                response = "Please specify which Excel file report you'd like to see. Example: 'show report for file 123'"
                st.session_state.messages.append({"role": "assistant", "content": response})
                
                with st.chat_message("assistant"):
                    st.markdown(response)
        finally:
            # Close session
            session.close()
    else:
        # Handle other types of requests
        # Process with your chatbot logic
        response = "How can I help you with Excel files today?"
        st.session_state.messages.append({"role": "assistant", "content": response})
        
        with st.chat_message("assistant"):
            st.markdown(response)
