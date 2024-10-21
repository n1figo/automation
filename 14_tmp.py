import re
import PyPDF2
from transformers import AutoTokenizer, AutoModelForQuestionAnswering
import torch

def load_pdf_text(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        texts_by_page = {i+1: page.extract_text() for i, page in enumerate(reader.pages)}
    return texts_by_page

# 방법 1: 정규 표현식을 사용한 검색
def find_end_phrases_regex(texts_by_page):
    end_phrases = [
        r"◇\s*질병관련\s*특약\s*\(자세한\s*사항은\s*반드시\s*약관을\s*참고하시기\s*바랍니다\.\s*\)",
        r"◇\s*상해\s*및\s*질병\s*관련\s*특약\s*\(자세한\s*사항은\s*반드시\s*약관을\s*참고하시기\s*바랍니다\.\s*\)",
        r"◇\s*재진단암진단비\s*특약\s*\(자세한\s*사항은\s*반드시\s*약관을\s*참고하시기\s*바랍니다\.\s*\)"
    ]
    
    combined_pattern = re.compile('|'.join(end_phrases), re.IGNORECASE)

    results = []
    for page_num, text in texts_by_page.items():
        matches = combined_pattern.findall(text)
        if matches:
            for match in matches:
                results.append((page_num, match.strip()))
                print(f"정규 표현식 방법: 페이지 {page_num}에서 발견: {match.strip()}")
    
    return results

# 방법 2: 단순 문자열 검색
def find_end_phrases_simple(texts_by_page):
    end_phrases = [
        "◇ 질병관련 특약 (자세한 사항은 반드시 약관을 참고하시기 바랍니다.)",
        "◇ 상해 및 질병 관련 특약(자세한 사항은 반드시 약관을 참고하시기 바랍니다.)",
        "◇ 재진단암진단비 특약(자세한 사항은 반드시 약관을 참고하시기 바랍니다.)"
    ]

    results = []
    for page_num, text in texts_by_page.items():
        for phrase in end_phrases:
            if phrase in text:
                results.append((page_num, phrase))
                print(f"단순 문자열 검색 방법: 페이지 {page_num}에서 발견: {phrase}")
    
    return results

# 방법 3: RAG를 사용한 검색
def setup_rag_model():
    tokenizer = AutoTokenizer.from_pretrained("deepset/roberta-base-squad2")
    model = AutoModelForQuestionAnswering.from_pretrained("deepset/roberta-base-squad2")
    return tokenizer, model

def find_end_phrases_rag(texts_by_page, tokenizer, model):
    question = "이 문서에서 질병관련 특약, 상해 및 질병 관련 특약, 또는 재진단암진단비 특약에 대한 언급이 있습니까? 있다면 어디에 있습니까?"
    
    results = []
    for page_num, text in texts_by_page.items():
        inputs = tokenizer(question, text, return_tensors="pt", max_length=512, truncation=True)
        
        with torch.no_grad():
            outputs = model(**inputs)
        
        answer_start = torch.argmax(outputs.start_logits)
        answer_end = torch.argmax(outputs.end_logits) + 1
        answer = tokenizer.decode(inputs["input_ids"][0][answer_start:answer_end])
        
        if any(phrase in answer for phrase in ["질병관련 특약", "상해 및 질병 관련 특약", "재진단암진단비 특약"]):
            results.append((page_num, answer.strip()))
            print(f"RAG 방법: 페이지 {page_num}에서 발견: {answer.strip()}")
    
    return results

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    texts_by_page = load_pdf_text(pdf_path)

    print("방법 1: 정규 표현식을 사용한 검색")
    regex_results = find_end_phrases_regex(texts_by_page)

    print("\n방법 2: 단순 문자열 검색")
    simple_results = find_end_phrases_simple(texts_by_page)

    print("\n방법 3: RAG를 사용한 검색")
    tokenizer, model = setup_rag_model()
    rag_results = find_end_phrases_rag(texts_by_page, tokenizer, model)

    print("\n결과 비교:")
    print(f"정규 표현식 방법: {len(regex_results)} 결과 찾음")
    print(f"단순 문자열 검색 방법: {len(simple_results)} 결과 찾음")
    print(f"RAG 방법: {len(rag_results)} 결과 찾음")

if __name__ == "__main__":
    main()