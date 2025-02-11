import sys
try:
    from llama_index.readers.file.video_audio import VideoAudioReader
    # llama_index 내부에서 VideoAudioParser를 참조할 때 VideoAudioReader를 대신 사용하도록 패치합니다.
    import llama_index.readers.file.video_audio as va_module
    va_module.VideoAudioParser = VideoAudioReader
except ImportError:
    pass

from llama_index.readers.llama_parsers import LlamaParseReader

# LlamaParseReader 인스턴스 초기화
loader = LlamaParseReader()

# 파싱할 PDF 파일의 경로를 지정합니다.
file_path = "/workspaces/automation/data/input/0211/KB Yes!365 건강보험(세만기)(무배당)(25.01)_0214_요약서_v1.1.pdf"

prompt = (
    "Extract all tables from the document by detecting the page ranges for each special clause "
    "('상해관련 특별약관', '질병관련 특별약관', '상해및질병관련특별약관'). "
    "Within each table, identify and extract any highlighted or colored text (e.g., red, blue, or highlighter-marked text). "
    "Return the results in a structured JSON format, separating sections by clause."
)

documents = loader.load_data(file_path=file_path, prompt=prompt)

for doc in documents:
    print(doc.text)