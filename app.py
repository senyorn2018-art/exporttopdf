import streamlit as st
import os
import tempfile
import sys
from docx2pdf import convert

# បើកដំណើរការ COM សម្រាប់ Windows ដើម្បីកុំឱ្យជាប់គាំងពេលបម្លែងច្រើន File
if sys.platform == "win32":
    import pythoncom

st.set_page_config(
    page_title="Word to PDF Converter Pro",
    page_icon="📑",
    layout="wide"
)

st.markdown("""
<style>
    .main-title {
        text-align: center;
        background: linear-gradient(135deg, #1E88E5, #43A047);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.3rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        text-align: center;
        color: #6c757d;
        font-size: 1rem;
        margin-bottom: 1.8rem;
    }
    .badge-size {
        background-color: #f1f3f5;
        color: #495057;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

if "converted_files" not in st.session_state:
    st.session_state.converted_files = {}

# អនុគមន៍ជំនួយសម្រាប់ Convert ឯកសារនីមួយៗ
def convert_docx_to_pdf(uploaded_file, target_dir):
    if sys.platform == "win32":
        pythoncom.CoInitialize() # កំណត់ Thread សម្រាប់ MS Word API
    
    # ប្រើឈ្មោះមានសុវត្ថិភាពដើម្បីការពារបញ្ហាអក្សរខ្មែរ ឬ Space
    safe_docx_path = os.path.join(target_dir, f"input_{uploaded_file.file_id}.docx")
    safe_pdf_path = os.path.join(target_dir, f"output_{uploaded_file.file_id}.pdf")
    
    with open(safe_docx_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    convert(safe_docx_path, safe_pdf_path)
    
    with open(safe_pdf_path, "rb") as f:
        pdf_bytes = f.read()
        
    return pdf_bytes

st.markdown('<div class="main-title">📑 Word to PDF Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">បម្លែងឯកសារ Word (.docx) ទៅជា PDF បានរហ័ស ងាយស្រួល និងទាញយកម្ដងមួយៗ</div>', unsafe_allow_html=True)

left_pad, main_col, right_pad = st.columns([1, 4, 1])

with main_col:
    uploaded_files = st.file_uploader(
        "📂 ទម្លាក់ ឬជ្រើសរើស File Word នៅទីនេះ (អាចជ្រើសរើសលើសពីមួយ)",
        type=["docx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.write("")
        col_header, col_btn_all = st.columns([2, 1])
        
        with col_header:
            st.markdown(f"#### 📋 បញ្ជីឯកសារបានបញ្ចូល ({len(uploaded_files)})")
        
        with col_btn_all:
            convert_all = st.button("🚀 Convert ទាំងអស់ (Convert All)", use_container_width=True, type="primary")

        # Logic បម្លែង File ទាំងអស់
        if convert_all:
            progress_bar = st.progress(0)
            status_text = st.empty()
            total_files = len(uploaded_files)

            with tempfile.TemporaryDirectory() as temp_dir:
                for idx, file in enumerate(uploaded_files):
                    status_text.text(f"កំពុងបម្លែង ({idx + 1}/{total_files}): {file.name}...")
                    try:
                        pdf_data = convert_docx_to_pdf(file, temp_dir)
                        st.session_state.converted_files[file.name] = pdf_data
                    except Exception as e:
                        st.error(f"បរាជ័យលើ File {file.name}: {e}")

                    progress_bar.progress((idx + 1) / total_files)

            status_text.success("✅ បានបម្លែងឯកសារទាំងអស់រួចរាល់!")
            st.rerun()

        st.divider()

        # បង្ហាញបញ្ជី File នីមួយៗ
        for index, uploaded_file in enumerate(uploaded_files):
            file_size_kb = uploaded_file.size / 1024
            size_label = f"{file_size_kb:.1f} KB" if file_size_kb < 1024 else f"{file_size_kb/1024:.2f} MB"
            pdf_filename = os.path.splitext(uploaded_file.name)[0] + ".pdf"

            with st.container():
                col_info, col_action = st.columns([3, 1])

                with col_info:
                    st.markdown(
                        f"""
                        <div style="padding: 4px 0;">
                            <b>{index + 1}. 📄 {uploaded_file.name}</b>
                            &nbsp;<span class="badge-size">{size_label}</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                with col_action:
                    if uploaded_file.name in st.session_state.converted_files:
                        pdf_data = st.session_state.converted_files[uploaded_file.name]
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_data,
                            file_name=pdf_filename,
                            mime="application/pdf",
                            key=f"dl_{index}",
                            use_container_width=True
                        )
                    else:
                        if st.button("⚡ Convert", key=f"btn_{index}", use_container_width=True):
                            with st.spinner("កំពុងបម្លែង..."):
                                with tempfile.TemporaryDirectory() as temp_dir:
                                    try:
                                        pdf_data = convert_docx_to_pdf(uploaded_file, temp_dir)
                                        st.session_state.converted_files[uploaded_file.name] = pdf_data
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"បរាជ័យ: {e}")

                st.divider()