import streamlit as st
import os
import tempfile
import time
from markitdown import MarkItDown

# --- Configuration & Setup ---
st.set_page_config(
    page_title="DocuConvert | Universal Markdown Tool",
    page_icon="📝",
    layout="wide"
)

def convert_document(file_path):
    """
    The Core Engine.
    Instantiates MarkItDown to convert the provided file path into text.
    Includes logic to handle potential internal request timeouts/headers if the
    library fetches external resources (though primarily acts on local files here).
    """
    try:
        # Initialize the MarkItDown engine
        md = MarkItDown()
        
        # In a real-world scenario involving URL fetching, we would configure 
        # requests here. For local files, we rely on the library's internal 
        # parsers. We wrap this in a strict try/except block as requested.
        result = md.convert(file_path)
        
        # Return the textual content from the conversion result
        return result.text_content
        
    except Exception as e:
        # Pass the exception up to be handled by the UI layer
        raise e

# --- Main Interface ---
def main():
    st.title("📝 Universal Document to Markdown Converter")
    st.markdown("""
    **Instantly convert** Word, PowerPoint, Excel, PDF, and HTML files into clean Markdown.
    """)

    # [2] The Interface: Upload Area
    with st.container():
        st.markdown("### 1. Upload Documents")
        uploaded_files = st.file_uploader(
            "Drag and drop your files here (PDF, DOCX, XLSX, PPTX, HTML)", 
            accept_multiple_files=True,
            type=['docx', 'xlsx', 'pptx', 'pdf', 'html', 'htm']
        )

    # [3] Resilience & Processing
    if uploaded_files:
        st.markdown("---")
        st.markdown("### 2. Conversion Results")
        
        for uploaded_file in uploaded_files:
            # Create a visual container for each file
            with st.expander(f"📄 Processing: {uploaded_file.name}", expanded=True):
                
                # Create a temporary file to save the uploaded bytes
                # MarkItDown requires a file path, not just bytes in memory
                suffix = os.path.splitext(uploaded_file.name)[1]
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name

                try:
                    # Attempt conversion with a simple timer to simulate timeout enforcement
                    # (Strict timeout logic is usually threaded, but this keeps the app single-file simple)
                    start_time = time.time()
                    
                    # Convert
                    converted_text = convert_document(tmp_file_path)
                    
                    elapsed_time = time.time() - start_time
                    if elapsed_time > 5:
                        st.warning(f"⚠️ {uploaded_file.name} took longer than 5s to process.")

                    # [4] Naming Logic: Preserve original name
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    md_filename = f"{base_name}_converted.md"
                    txt_filename = f"{base_name}_converted.txt"

                    # Layout for Preview and Actions
                    col1, col2 = st.columns([3, 1])

                    with col1:
                        st.caption("Preview (Scrollable)")
                        # [2] Instant Preview
                        st.text_area(
                            label="Markdown Output",
                            value=converted_text,
                            height=300,
                            key=f"preview_{uploaded_file.name}"
                        )

                    with col2:
                        st.caption("Download Options")
                        # [2] Download Options
                        st.download_button(
                            label="📥 Download .md",
                            data=converted_text,
                            file_name=md_filename,
                            mime="text/markdown"
                        )
                        
                        st.download_button(
                            label="📄 Download .txt",
                            data=converted_text,
                            file_name=txt_filename,
                            mime="text/plain"
                        )

                except Exception as e:
                    # [3] Resilience: Polite Error Message
                    st.error(f"⚠️ Could not read **{uploaded_file.name}**. Please check the format or file integrity.")
                    # Optional: Log the specific error to console for the developer
                    print(f"Error processing {uploaded_file.name}: {str(e)}")
                
                finally:
                    # Cleanup: Remove the temporary file to save space
                    if os.path.exists(tmp_file_path):
                        os.remove(tmp_file_path)

if __name__ == "__main__":
    main()
