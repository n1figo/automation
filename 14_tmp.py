import re
import PyPDF2
from transformers import AutoTokenizer, AutoModelForQuestionAnswering
import torch

def load_pdf_text(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        texts_by_page = {i+1: page.extract_text() for i, page in enumerate(reader.pages)}
    return texts_by_page

def normalize_text(text):
    # 모든 특수문자 제거 및 공백 제거
    return re.sub(r'\W+', '', text.lower())

# 방법 1: 정규화된 텍스트를 사용한 검색
def find_end_phrases_normalized(texts_by_page):
    end_phrases = [
        "질병관련특약자세한사항은반드시약관을참고하시기바랍니다",
        "상해및질병관련특약자세한사항은반드시약관을참고하시기바랍니다",
        "재진단암진단비특약자세한사항은반드시약관을참고하시기바랍니다"
    ]

    results = []
    for page_num, text in texts_by_page.items():
        normalized_text = normalize_text(text)
        for phrase in end_phrases:
            if phrase in normalized_text:
                original_phrase = re.search(r'◇\s*.*?(?=◇|\Z)', text, re.DOTALL)
                if original_phrase:
                    original_phrase = original_phrase.group().strip()
                    results.append((page_num, original_phrase))
                    print(f"정규화된 텍스트 검색 방법: 페이지 {page_num}에서 발견: {original_phrase}")
    
    return results

# 방법 2: 부분 문자열 매칭
def find_end_phrases_partial(texts_by_page):
    end_phrases = [
        "질병관련 특약",
        "상해 및 질병 관련 특약",
        "재진단암진단비 특약"
    ]

    results = []
    for page_num, text in texts_by_page.items():
        for phrase in end_phrases:
            if phrase in text:
                full_phrase = re.search(f'◇.*{re.escape(phrase)}.*?(?=◇|\Z)', text, re.DOTALL)
                if full_phrase:
                    full_phrase = full_phrase.group().strip()
                    results.append((page_num, full_phrase))
                    print(f"부분 문자열 매칭 방법: 페이지 {page_num}에서 발견: {full_phrase}")
    
    return results

# 방법 3: RAG를 사용한 검색 (수정됨)
def setup_rag_model():
    tokenizer = AutoTokenizer.from_pretrained("deepset/roberta-base-squad2")
    model = AutoModelForQuestionAnswering.from_pretrained("deepset/roberta-base-squad2")
    return tokenizer, model

def find_end_phrases_rag(texts_by_page, tokenizer, model):
    question = "이 문서에서 질병관련 특약, 상해 및 질병 관련 특약, 또는 재진단암진단비 특약에 대한 언급이 있습니까? 있다면 전체 문구를 알려주세요."
    
    results = []
    for page_num, text in texts_by_page.items():
        inputs = tokenizer(question, text, return_tensors="pt", max_length=512, truncation=True)
        
        with torch.no_grad():
            outputs = model(**inputs)
        
        answer_start = torch.argmax(outputs.start_logits)
        answer_end = torch.argmax(outputs.end_logits) + 1
        answer = tokenizer.decode(inputs["input_ids"][0][answer_start:answer_end])
        
        if any(phrase in answer for phrase in ["질병관련 특약", "상해 및 질병 관련 특약", "재진단암진단비 특약"]):
            full_phrase = re.search(r'◇.*?(?=◇|\Z)', answer, re.DOTALL)
            if full_phrase:
                full_phrase = full_phrase.group().strip()
                results.append((page_num, full_phrase))
                print(f"RAG 방법: 페이지 {page_num}에서 발견: {full_phrase}")
    
    return results

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    texts_by_page = load_pdf_text(pdf_path)

    print("방법 1: 정규화된 텍스트를 사용한 검색")
    normalized_results = find_end_phrases_normalized(texts_by_page)

    print("\n방법 2: 부분 문자열 매칭")
    partial_results = find_end_phrases_partial(texts_by_page)

    print("\n방법 3: RAG를 사용한 검색")
    tokenizer, model = setup_rag_model()
    rag_results = find_end_phrases_rag(texts_by_page, tokenizer, model)

    print("\n결과 비교:")
    print(f"정규화된 텍스트 검색 방법: {len(normalized_results)} 결과 찾음")
    print(f"부분 문자열 매칭 방법: {len(partial_results)} 결과 찾음")
    print(f"RAG 방법: {len(rag_results)} 결과 찾음")

if __name__ == "__main__":
    main()