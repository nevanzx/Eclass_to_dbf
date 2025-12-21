#!/usr/bin/env python3
"""
Script to add PDF viewer section to the app.py file
"""

import os
import re

# Read the current app.py file
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Find the exact location to insert before 'if __name__ == "__main__":'
marker = 'if __name__ == "__main__":'
if marker in content:
    parts = content.rsplit(marker, 1)
    before_main = parts[0]
    after_main = parts[1]
    
    # The new PDF viewer section
    pdf_viewer_code = '''    # Check if we have a dataframe to work with
    if 'df' in st.session_state and st.session_state.df is not None:
        df = st.session_state.df

        # Get JLE data if available
        jle_data = st.session_state.get('jle_data', None) if 'jle_data' in st.session_state else None

        # Show PDF from Word report if it was generated
        if st.session_state.get('pdf_from_word_generated', False) and 'pdf_from_word_bytes' in st.session_state:
            # Create side-by-side layout for DBF viewer and PDF viewer
            viewer_col1, viewer_col2 = st.columns(2)

            with viewer_col1:
                # Show DBF viewer in its own card
                with st.container(border=True):
                    st.subheader("📋 Updated DBF Content")
                    st.dataframe(df, use_container_width=True, height=400)

            with viewer_col2:
                # Show PDF from Word report if it was generated
                with st.container(border=True):
                    st.subheader("📄 Generated PDF Report")
                    
                    # Convert PDF bytes to base64 for embedding in HTML
                    import base64
                    pdf_bytes = st.session_state.pdf_from_word_bytes
                    if isinstance(pdf_bytes, str):
                        pdf_bytes = pdf_bytes.encode('latin-1')
                    elif isinstance(pdf_bytes, bytearray):
                        pdf_bytes = bytes(pdf_bytes)

                    pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')

                    # Embed the PDF in an iframe
                    pdf_display = f'<iframe src="data:application/pdf;base64,{pdf_base64}" width="100%" height="600" type="application/pdf">'
                    st.markdown(pdf_display, unsafe_allow_html=True)

                    # Download button for PDF from Word
                    st.download_button(
                        label="📥 Download PDF Report",
                        data=st.session_state.pdf_from_word_bytes,
                        file_name=st.session_state.pdf_from_word_filename,
                        mime="application/pdf",
                        key="download_pdf_from_word_btn"
                    )


'''

    # Combine the parts with the new section
    new_content = before_main + pdf_viewer_code + marker + after_main
    
    # Write the updated content back
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print('PDF viewer section added successfully!')
else:
    print('Could not find the main function marker')