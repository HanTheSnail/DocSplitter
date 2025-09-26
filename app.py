import streamlit as st
import io
import zipfile
from docx import Document

def _top_level_tables(body):
    """Return a list of top-level <w:tbl> elements (not nested in cells)."""
    # body.iterchildren() preserves document order; filter on localname 'tbl'
    return [el for el in body.iterchildren() if el.tag.endswith('}tbl')]

def extract_tables_from_docx(uploaded_file):
    """
    Extract each TOP-LEVEL table as a standalone .docx by
    loading the original, removing all other body content, and saving.
    This preserves formatting, merges, widths, styles, images, hyperlinks, and SDTs (drop-downs).
    """
    # Get raw bytes once; UploadedFile can be read multiple times safely via getvalue()
    file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

    # Count top-level tables
    probe_doc = Document(io.BytesIO(file_bytes))
    probe_body = probe_doc._element.body
    tables = _top_level_tables(probe_body)

    if not tables:
        return [], "No tables found in the document."

    extracted_docs = []

    # For each table, make a fresh copy of the original doc and delete everything else
    for i in range(len(tables)):
        d = Document(io.BytesIO(file_bytes))
        b = d._element.body

        tbl_index = 0
        # Iterate over a static copy since we'll remove from the tree
        for child in list(b.iterchildren()):
            if child.tag.endswith('}tbl'):
                tbl_index += 1
                if tbl_index == i + 1:
                    # keep this table
                    continue
            # remove any non-target table or any paragraph/other element
            b.remove(child)

        # Save this single-table document to memory
        buf = io.BytesIO()
        d.save(buf)
        buf.seek(0)

        extracted_docs.append({
            "name": f"table_{i+1}.docx",
            "content": buf.getvalue()
        })

    return extracted_docs, None


def create_zip_download(docs, original_filename):
    """Create a zip file containing all extracted table documents"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        base_name = original_filename.rsplit(".", 1)[0]
        for doc in docs:
            new_filename = f"{base_name}_{doc['name']}"
            zip_file.writestr(new_filename, doc["content"])
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def main():
    st.set_page_config(
        page_title="Word Table Splitter",
        page_icon="📊",
        layout="centered"
    )

    # Session state
    if "processed_results" not in st.session_state:
        st.session_state.processed_results = []
    if "processing_complete" not in st.session_state:
        st.session_state.processing_complete = False

    st.title("📊 Word Document Table Splitter")
    st.markdown("Upload Word documents (.docx) and automatically split each **top-level** table into separate documents — formatting and drop-downs preserved.")

    uploaded_files = st.file_uploader(
        "Choose Word documents",
        type=["docx"],
        accept_multiple_files=True,
        help="Select one or more .docx files containing tables you want to split",
        key="docx_uploader"
    )

    if uploaded_files:
        st.write(f"📁 {len(uploaded_files)} file(s) uploaded")

        if st.button("🔄 Split Tables", type="primary", key="split_button"):
            st.session_state.processed_results = []
            st.session_state.processing_complete = False

            all_results = []

            # Use placeholders so we can cleanly remove progress UI
            progress_ph = st.empty()
            status_text = st.empty()
            progress_bar = progress_ph.progress(0)

            total = len(uploaded_files)
            for idx, uploaded_file in enumerate(uploaded_files, start=1):
                status_text.text(f"Processing: {uploaded_file.name}")
                progress_bar.progress(idx / total)

                extracted_docs, error = extract_tables_from_docx(uploaded_file)

                if error:
                    st.error(f"❌ Error with {uploaded_file.name}: {error}")
                    continue

                if extracted_docs:
                    all_results.append({
                        "filename": uploaded_file.name,
                        "docs": extracted_docs,
                        "count": len(extracted_docs)
                    })
                    st.success(f"✅ {uploaded_file.name}: Found {len(extracted_docs)} top-level table(s)")

            st.session_state.processed_results = all_results
            st.session_state.processing_complete = True

            status_text.empty()
            progress_ph.empty()

        if st.session_state.processing_complete and st.session_state.processed_results:
            st.markdown("---")
            st.subheader("📥 Download Split Documents")

            for idx, result in enumerate(st.session_state.processed_results):
                st.markdown(f"**{result['filename']}** — {result['count']} table(s)")

                if result["count"] == 1:
                    doc = result["docs"][0]
                    st.download_button(
                        label=f"📄 Download {doc['name']}",
                        data=doc["content"],
                        file_name=f"{result['filename'].rsplit('.', 1)[0]}_{doc['name']}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"single_download_{idx}"
                    )
                else:
                    zip_data = create_zip_download(result["docs"], result["filename"])
                    st.download_button(
                        label="📦 Download All Tables as ZIP",
                        data=zip_data,
                        file_name=f"{result['filename'].rsplit('.', 1)[0]}_tables.zip",
                        mime="application/zip",
                        key=f"zip_download_{idx}"
                    )

            if len(st.session_state.processed_results) > 1:
                st.markdown("---")
                master_zip_buffer = io.BytesIO()
                with zipfile.ZipFile(master_zip_buffer, "w", zipfile.ZIP_DEFLATED) as master_zip:
                    for result in st.session_state.processed_results:
                        base_name = result["filename"].rsplit(".", 1)[0]
                        for doc in result["docs"]:
                            filename = f"{base_name}_{doc['name']}"
                            master_zip.writestr(filename, doc["content"])
                master_zip_buffer.seek(0)
                st.download_button(
                    label="📦 Download All Files as One ZIP",
                    data=master_zip_buffer.getvalue(),
                    file_name="all_split_tables.zip",
                    mime="application/zip",
                    type="primary",
                    key="master_zip_download"
                )

        elif st.session_state.processing_complete and not st.session_state.processed_results:
            st.warning("⚠️ No tables were found in any of the uploaded documents.")

    with st.expander("ℹ️ How to use"):
        st.markdown("""
        1) **Upload** your Word documents (.docx)  
        2) **Click** “Split Tables”  
        3) **Download** the split documents:
           - Single table → direct .docx
           - Multiple tables → ZIP archive
        
        **Notes**
        - Only **top-level** tables in the main document body are split (tables nested inside cells aren’t split individually).
        - Because we keep the original table XML, **drop-downs, checkboxes (SDTs), merges, widths, borders, styles, hyperlinks, and images** are preserved.
        """)

if __name__ == "__main__":
    main()
