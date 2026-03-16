import streamlit as st
import streamlit.components.v1 as components
import os
import io
import base64
from translator import TranslationService
from pptx_processor import PPTXProcessor
from pdf_processor import PDFProcessor


def auto_download(data, filename, mime_type):
    """Trigger automatic file download using JavaScript with Blob for large files."""
    b64 = base64.b64encode(data).decode()
    # Use Blob approach to handle large files reliably
    components.html(
        f'''
        <script>
            // Decode base64 to binary
            const b64 = "{b64}";
            const byteChars = atob(b64);
            const byteArrays = [];
            const sliceSize = 512;
            for (let offset = 0; offset < byteChars.length; offset += sliceSize) {{
                const slice = byteChars.slice(offset, offset + sliceSize);
                const byteNumbers = new Array(slice.length);
                for (let i = 0; i < slice.length; i++) {{
                    byteNumbers[i] = slice.charCodeAt(i);
                }}
                byteArrays.push(new Uint8Array(byteNumbers));
            }}
            const blob = new Blob(byteArrays, {{type: "{mime_type}"}});
            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = url;
            link.download = "{filename}";
            document.body.appendChild(link);
            setTimeout(() => {{
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            }}, 100);
        </script>
        ''',
        height=0,
    )

st.set_page_config(page_title="SAP PPT/PDF Translator", page_icon="📊", layout="wide")

col1, col2 = st.columns([0.1, 0.9])
with col1:
    st.image("logo.png", width=150)
with col2:
    st.title("SAP PPT 자동 번역 프로그램")

st.markdown("""
영문 SAP PPT 또는 PDF 파일을 드래그 앤 드롭하면 서식을 유지한 채 한글로 번역해줍니다.
`PPTX` → 번역된 `PPTX` | `PDF` → 번역된 `DOCX`
""")

with st.sidebar:
    st.header("⚙️ 설정")
    service_type = st.selectbox("번역 엔진 선택", ["Smart (OpenAI -> Free)", "Free (Google)", "DeepL", "OpenAI"])
    
    api_key = None
    if service_type in ["OpenAI", "Smart (OpenAI -> Free)"]:
        api_key = st.text_input("OpenAI API Key", type="password", help="OpenAI API 키는 https://platform.openai.com/api-keys 에서 발급받을 수 있습니다. 신용카드 등록 및 충전이 필요할 수 있습니다.")
        if service_type == "Smart (OpenAI -> Free)":
            st.caption("✨ OpenAI로 우선 번역하며, 키가 없거나 한도 초과 시 무료 엔진으로 자동 전환됩니다.")
    elif service_type == "DeepL":
        api_key = st.text_input("DeepL API Key", type="password", help="DeepL 계정 페이지의 'Authentication Key'를 복사해 넣으세요.")
    else:
        st.info("🔓 무료 엔진은 API 키가 필요하지 않습니다.")

    st.divider()
    translation_level = st.radio(
        "번역 수준",
        ["normal", "thorough"],
        index=0,
        format_func=lambda x: {
            "normal": "🔸 표준 (권장 - 제목/라벨/도형 유지, 본문 번역)",
            "thorough": "🔻 전체 (대부분 번역, 제목만 유지)"
        }[x],
        help="SAP 프레젠테이션에는 '표준'을 권장합니다."
    )

    st.info("💡 SAP 전문 용어(MRP, BDC 등)는 `glossary.json`의 정의를 따릅니다.")

uploaded_file = st.file_uploader("PPTX 또는 PDF 파일을 업로드하세요", type=["pptx", "pdf"])

if uploaded_file is not None:
    # Determine file type
    file_ext = os.path.splitext(uploaded_file.name)[1].lower()
    is_pdf = file_ext == ".pdf"

    # Validate API key only for non-free services
    is_key_required = service_type in ["DeepL", "OpenAI"]
    if is_key_required and not api_key:
        st.warning(f"⚠️ {service_type} 사용을 위해 API 키를 입력해주세요.")
    else:
        if st.button("번역 시작"):
            with st.status("번역 중...", expanded=True) as status:
                try:
                    file_bytes = uploaded_file.getvalue()
                    file_size = len(file_bytes)

                    # Log file info for debugging
                    st.write(f"📁 파일 크기: {file_size / 1024:.2f} KB")
                    st.write(f"📄 파일 형식: {'PDF' if is_pdf else 'PPTX'}")

                    if not is_pdf:
                        import zipfile
                        if not zipfile.is_zipfile(io.BytesIO(file_bytes)):
                            st.error("⚠️ 업로드된 파일이 유효한 PPTX(Zip) 형식이 아닙니다. 혹시 구버전 PPT(97-2003) 파일인가요? .pptx 형식만 지원합니다.")
                            st.stop()

                    # Use BytesIO
                    input_stream = io.BytesIO(file_bytes)
                    output_stream = io.BytesIO()

                    translator = TranslationService(service_type=service_type, api_key=api_key)

                    progress_bar = st.progress(0)

                    def update_progress(progress):
                        progress_bar.progress(progress)

                    # Ensure pointer is at start
                    input_stream.seek(0)

                    if is_pdf:
                        # PDF → DOCX
                        processor = PDFProcessor(translator, translation_level=translation_level)
                        output_result, errors = processor.process_pdf(input_stream, output_stream, progress_callback=update_progress)
                        output_filename = os.path.splitext(uploaded_file.name)[0] + "_KO.docx"
                        output_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    else:
                        # PPTX → PPTX
                        processor = PPTXProcessor(translator, translation_level=translation_level)
                        output_result, errors = processor.process_presentation(input_stream, output_stream, progress_callback=update_progress)
                        output_filename = os.path.splitext(uploaded_file.name)[0] + "_KO.pptx"
                        output_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

                    # Get results from output stream
                    output_data = output_stream.getvalue()
                    if not output_data:
                        st.error("⚠️ 번역 결과 파일이 비어 있습니다. 처리 중 오류가 발생했을 수 있습니다.")
                        st.stop()

                    status.update(label="번역 완료!", state="complete", expanded=False)
                    st.success(f"✅ 번역이 완료되었습니다. (파일 크기: {len(output_data)/1024/1024:.2f} MB)")
                    if is_pdf:
                        st.info("📝 PDF는 텍스트를 추출하여 DOCX(Word) 파일로 변환됩니다.")

                    if errors:
                        with st.expander("📝 번역 시 발생한 일부 오류 (디버깅용)", expanded=False):
                            for err in errors:
                                st.write(f"- {err}")

                    # Auto-download the translated file
                    auto_download(output_data, output_filename, output_mime)

                    # Also show manual download button as fallback
                    st.download_button(
                        label="📥 다운로드가 안 됐다면 여기를 클릭",
                        data=output_data,
                        file_name=output_filename,
                        mime=output_mime
                    )
                except Exception as e:
                    st.error(f"번역 중 오류 발생: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
