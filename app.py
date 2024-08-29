import streamlit as st
from docx import Document
from pptx import Presentation
import io
import base64

# Define a password
PASSWORD = "satish"

# Prompt the user for the password
def check_password():
    def password_entered():
        if st.session_state["password"] == PASSWORD:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # remove the password from the state
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password
        st.text_input("Enter Password", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error
        st.text_input("Enter Password", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct
        return True

if check_password():
    # Your Streamlit app code goes here
    st.title("Document Text Replacer")
    
    st.sidebar.header("Upload File")
    uploaded_file = st.sidebar.file_uploader("Upload a File", type=["docx", "pptx"])
    
    st.sidebar.header("Find and Replace")
    user_find = st.sidebar.text_input("Find:")
    user_replace = st.sidebar.text_input("Replace with:")
    
    st.sidebar.header("Select Segments")
    selected_segment = st.sidebar.selectbox("Select a segment", options=segment_options_ordered)
    
    if selected_segment:
        selected_segments = get_segments_up_to(selected_segment)
        st.sidebar.write(f"Selected Segments: {', '.join(selected_segments)}")
    
    with st.expander("Segment Replacements"):
        segment_replace_inputs = {}
        for segment in selected_segments:
            st.subheader(segment)
            segment_replace_inputs[segment] = {}
            for subsegment in segment_options[segment]:
                segment_replace_inputs[segment][subsegment] = st.text_input(f"Replace {subsegment} with:", key=f"{segment}_{subsegment}")
    
    with st.expander("Company Replacements"):
        company_replace_inputs = {value: st.text_input(f"Replace {value} with:", key=f"COMPANY_{value}") for value in company_options}
    
    st.sidebar.header("Download Options")
    custom_filename = st.sidebar.text_input("Enter Filename:", "")
    
    if st.button("Update File"):
        if uploaded_file:
            try:
                file_content = io.BytesIO(uploaded_file.getvalue())
                filename = uploaded_file.name
                
                find_replace_pairs = [(find_str, replace_str) for segment in segment_replace_inputs.values() for find_str, replace_str in segment.items() if replace_str]
                
                for find_str, replace_str in company_replace_inputs.items():
                    if replace_str:
                        find_replace_pairs.append((find_str, replace_str))
                
                if user_find and user_replace:
                    find_replace_pairs.append((user_find, user_replace))
                
                if filename.endswith('.docx'):
                    doc = Document(file_content)
                    replace_word_in_docx(doc, find_replace_pairs)
                    output_buffer = io.BytesIO()
                    doc.save(output_buffer)
                    output_buffer.seek(0)
                    
                    b64 = base64.b64encode(output_buffer.read()).decode()
                    href = f'<a download="{custom_filename if custom_filename else "modified_document"}.docx" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" target="_blank">Download Updated Word File</a>'
                    st.markdown(href, unsafe_allow_html=True)
                
                elif filename.endswith('.pptx'):
                    ppt = Presentation(file_content)
                    replace_text_in_pptx(ppt, find_replace_pairs)
                    output_buffer = io.BytesIO()
                    ppt.save(output_buffer)
                    output_buffer.seek(0)
                    
                    b64 = base64.b64encode(output_buffer.read()).decode()
                    href = f'<a download="{custom_filename if custom_filename else "modified_presentation"}.pptx" href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" target="_blank">Download Updated PPT File</a>'
                    st.markdown(href, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Upload a File.")
