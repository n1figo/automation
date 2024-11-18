from langchain_community.document_loaders import PyPDFLoader
from langchain_community.llms import LlamaCpp
from langchain.prompts import PromptTemplate

def find_special_terms_pages(pdf_path):
    # PDF 로더 초기화
    loader = PyPDFLoader(pdf_path)
    pages = loader.load_and_split()
    
    special_terms_pages = {
        "상해관련 특별약관": [],
        "질병관련 특별약관": []
    }

    # 각 페이지 분석
    for i, page in enumerate(pages):
        text = page.page_content
        
        # 키워드 검색
        if "상해관련 특별약관" in text:
            special_terms_pages["상해관련 특별약관"].append(i + 1)
        if "질병관련 특별약관" in text:
            special_terms_pages["질병관련 특별약관"].append(i + 1)

    return special_terms_pages

if __name__ == "__main__":
    # 필요한 패키지 설치
    # pip install langchain-community
    
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    results = find_special_terms_pages(pdf_path)
    
    # 결과 출력
    for term, pages in results.items():
        if pages:
            print(f"{term}: {pages} 페이지에서 발견됨")
        else:
            print(f"{term}: 발견되지 않음")
            # 69페이지 텍스트 추출
            page_69 = loader.load_and_split()[68]  # 0-based index
            with open("page_69.txt", "w", encoding="utf-8") as f:
                f.write(page_69.page_content)

            # Llama 모델로 상해관련 특별약관 검색

            llm = LlamaCpp(
                model_path="/path/to/llama/model.gguf",
                temperature=0.1,
                max_tokens=2000
            )

            prompt = PromptTemplate(
                input_variables=["content"],
                template="다음 보험약관 내용에서 상해관련 특별약관을 찾아서 설명해주세요:\n\n{content}"
            )

            response = llm(prompt.format(content=page_69.page_content))
            print("\nLLM 분석 결과:")
            print(response)