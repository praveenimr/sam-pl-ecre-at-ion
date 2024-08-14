import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.util import Pt
import io
import base64

# Options for segments and their subsegments
segment_options = {
    'SEGMENTTA': ['SEGMENTTA','SUBSEGA1', 'SUBSEGA2', 'SUBSEGA3', 'SUBSEGA4', 'SUBSEGA5', 'SUBSEGA6'],
    'SEGMENTTB': ['SEGMENTTB','SUBSEGB1', 'SUBSEGB2', 'SUBSEGB3', 'SUBSEGB4', 'SUBSEGB5', 'SUBSEGB6'],
    'SEGMENTTC': ['SEGMENTTC','SUBSEGC1', 'SUBSEGC2', 'SUBSEGC3', 'SUBSEGC4', 'SUBSEGC5', 'SUBSEGC6'],
    'SEGMENTTD': ['SEGMENTTD','SUBSEGD1', 'SUBSEGD2', 'SUBSEGD3', 'SUBSEGD4', 'SUBSEGD5', 'SUBSEGD6'],
    'SEGMENTTE': ['SEGMENTTE','SUBSEGE1', 'SUBSEGE2', 'SUBSEGE3', 'SUBSEGE4', 'SUBSEGE5', 'SUBSEGE6'],
    'SEGMENTTF': ['SEGMENTTF','SUBSEGF1', 'SUBSEGF2', 'SUBSEGF3', 'SUBSEGF4', 'SUBSEGF5', 'SUBSEGF6'],
}
company_options = ['COMPANYA', 'COMPANYB', 'COMPANYC', 'COMPANYD', 'COMPANYE', 'COMPANYF', 'COMPANYG', 'COMPANYH', 'COMPANYI', 'COMPANYJ', 'COMPANYK', 'COMPANYL', 'COMPANYM', 'COMPANYN', 'COMPANYO', 'COMPANYP', 'COMPANYQ', 'COMPANYR', 'COMPANYS', 'COMPANYT']

def replace_text_case_insensitive(paragraphs, find_str, replace_str, font_name="Segoe UI"):
    find_str_lower = find_str.lower()
    for para in paragraphs:
        text = para.text
        text_lower = text.lower()
        start = 0
        while True:
            start = text_lower.find(find_str_lower, start)
            if start == -1:
                break
            end = start + len(find_str)
            para.text = para.text[:start] + replace_str + para.text[end:]
            for run in para.runs:
                if run.text == replace_str:
                    run.font.name = font_name
            text = para.text
            text_lower = text.lower()
            start = end

def replace_text_in_pptx(slides, find_str, replace_str, font_name="Segoe UI"):
    find_str_lower = find_str.lower()
    for slide in slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                text = paragraph.text
                text_lower = text.lower()
                start = 0
                while True:
                    start = text_lower.find(find_str_lower, start)
                    if start == -1:
                        break
                    end = start + len(find_str)
                    paragraph.text = paragraph.text[:start] + replace_str + paragraph.text[end:]
                    for run in paragraph.runs:
                        if run.text == replace_str:
                            run.font.name = font_name
                    text = paragraph.text
                    text_lower = text.lower()
                    start = end

def replace_word_in_docx(doc, find_replace_pairs):
    for find_str, replace_str in find_replace_pairs:
        # Replace in regular paragraphs
        replace_text_case_insensitive(doc.paragraphs, find_str, replace_str)
        
        # Replace in table cells
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_text_case_insensitive(cell.paragraphs, find_str, replace_str)  # Iterate over paragraphs within the cell
                    
        # Replace in headers and footers
        for section in doc.sections:
            replace_text_case_insensitive(section.header.paragraphs, find_str, replace_str)
            replace_text_case_insensitive(section.footer.paragraphs, find_str, replace_str)


def main():
    st.title("Sample Creation")
    
    st.sidebar.header("Upload File")
    uploaded_file = st.sidebar.file_uploader("Upload a .docx or .pptx file", type=["docx", "pptx"])
    
    st.sidebar.header("Find and Replace")
    user_find = st.sidebar.text_input("Find:")
    user_replace = st.sidebar.text_input("Replace with:")
    
    st.sidebar.header("Select Segments")
    selected_segments = st.sidebar.multiselect("Select segments to replace", options=list(segment_options.keys()))
    
    # st.header("Replacements")
    
    segment_replace_inputs = {}
    for segment in selected_segments:
        st.subheader(segment)
        segment_replace_inputs[segment] = {}
        for subsegment in segment_options[segment]:
            segment_replace_inputs[segment][subsegment] = st.text_input(f"Replace {subsegment} with:", key=f"{segment}_{subsegment}")
    
    st.subheader("KEY-COMPANY Replacements")
    company_replace_inputs = {value: st.text_input(f"Replace {value} with:", key=f"COMPANY_{value}") for value in company_options}
    
    st.sidebar.header("Download Options")
    custom_filename = st.sidebar.text_input("Enter Filename:",)
    
    if st.button("Replace Text"):
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
                    href = f'<a download="{custom_filename}.docx" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" target="_blank">Download Modified Document</a>'
                    st.markdown(href, unsafe_allow_html=True)
                
                elif filename.endswith('.pptx'):
                    ppt = Presentation(file_content)
                    replace_ppt_in_pptx(ppt, find_replace_pairs)
                    output_buffer = io.BytesIO()
                    ppt.save(output_buffer)
                    output_buffer.seek(0)
                    
                    b64 = base64.b64encode(output_buffer.read()).decode()
                    href = f'<a download="{custom_filename}.pptx" href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" target="_blank">Download Modified Presentation</a>'
                    st.markdown(href, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please upload a .docx or .pptx file.")
            
if __name__ == "__main__":
    main()
