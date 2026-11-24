import streamlit as st
import whisper
import tempfile
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from io import BytesIO

# 页面设定
st.set_page_config(
    page_title="语音转文字工具",
    page_icon="🎙️",
    layout="centered"
)

st.title("🎙️ 语音转文字工具")
st.markdown("上传 MP3 / WAV / M4A 等语音档，自动转换成文字并生成 Word 文件")

# 侧边栏设定
with st.sidebar:
    st.header("设定")
    model_size = st.selectbox(
        "Whisper 模型大小",
        options=["tiny", "base", "small", "medium"],
        index=2,  # 预设 small
        help="模型越大准确度越高，但速度越慢。建议先用 small 或 base。"
    )
    language = st.selectbox(
        "语言",
        options=["zh", "en", "自动侦测"],
        index=0
    )
    st.markdown("---")
    st.caption("使用 OpenAI Whisper 进行语音辨识")

# 上传档案
uploaded_file = st.file_uploader(
    "请上传语音档案",
    type=["mp3", "wav", "m4a", "ogg", "flac", "mp4"],
    help="支援常见音频格式"
)

@st.cache_resource
def load_whisper_model(size: str):
    """加载模型（使用缓存，避免重复加载）"""
    return whisper.load_model(size)

def create_word_document(text: str) -> BytesIO:
    """产生 Word 文件并回传 BytesIO"""
    doc = Document()
    
    # 设定中文字体
    style = doc.styles['Normal']
    font = style.font
    font.name = '微软雅黑'
    font.size = Pt(12)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    
    doc.add_heading('语音转文字结果', level=1)
    doc.add_paragraph(text)
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if uploaded_file is not None:
    st.audio(uploaded_file, format=f'audio/{uploaded_file.type.split("/")[-1]}')
    
    if st.button("开始转换", type="primary"):
        with st.spinner("正在加载模型与转换中，请稍候..."):
            # 储存成临时档案
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            
            try:
                # 加载模型
                model = load_whisper_model(model_size)
                
                # 转写
                lang = None if language == "自动侦测" else language
                result = model.transcribe(
                    tmp_path,
                    language=lang,
                    fp16=False,
                    verbose=False
                )
                text = result["text"].strip()
                
                # 显示结果
                st.success("转换完成！")
                st.subheader("转换结果")
                st.text_area("文字内容", value=text, height=300)
                
                # 产生 Word 并提供下载
                word_buffer = create_word_document(text)
                st.download_button(
                    label="📥 下载 Word 文件",
                    data=word_buffer,
                    file_name="语音转文字结果.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"发生错误：{str(e)}")
            finally:
                # 删除临时档案
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
else:
    st.info("请先上传一个语音档案")
