import torch
from transformers import AutoTokenizer, AutoModelForCausalLM
from huggingface_hub import login
from sentence_transformers import SentenceTransformer
import faiss
import numpy as np
import PyPDF2
from typing import Dict, List, Tuple
import logging
import os
from dotenv import load_dotenv

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 환경변수 로드
load_dotenv()

class TableEndDetector:
    def __init__(self, model_name: str = "beomi/KoAlpaca-350m-v1.1"): # beomi/KoAlpaca-350m-v1.1  # TinyLlama/TinyLlama-1.1B-Chat-v1.0
        """
        Initialize the detector with specified models
        :param model_name: TinyLlama model name
        """
        # Hugging Face 토큰 설정 및 로그인
        hf_token = os.getenv("HUGGING_FACE_TOKEN")
        if not hf_token:
            raise ValueError("HUGGING_FACE_TOKEN이 환경변수에 설정되어 있지 않습니다.")
        login(token=hf_token)

        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        logger.info(f"Using device: {self.device}")
        
        # TinyLlama 모델 초기화
        logger.info("TinyLlama 모델 로딩 중...")
        self.tokenizer = AutoTokenizer.from_pretrained(model_name)
        self.model = AutoModelForCausalLM.from_pretrained(
            model_name,
            torch_dtype=torch.float16 if self.device == "cuda" else torch.float32,
            device_map='auto'
        )
        logger.info("TinyLlama 모델 로딩 완료")
        
        # 한국어 임베딩 모델 초기화
        logger.info("한국어 임베딩 모델 로딩 중...")
        self.embedding_model = SentenceTransformer('jhgan/ko-sbert-nli')
        logger.info("한국어 임베딩 모델 로딩 완료")

    def create_embeddings(self, texts: List[str]) -> np.ndarray:
        """
        텍스트 리스트의 임베딩을 생성
        """
        try:
            embeddings = self.embedding_model.encode(
                texts,
                batch_size=32,
                show_progress_bar=True,
                normalize_embeddings=True
            )
            return embeddings
        except Exception as e:
            logger.error(f"임베딩 생성 중 오류 발생: {e}")
            raise

    def setup_rag_index(self, texts: List[str]) -> Tuple[faiss.Index, List[str]]:
        """
        RAG 검색을 위한 FAISS 인덱스 설정
        """
        embeddings = self.create_embeddings(texts)
        dimension = embeddings.shape[1]
        
        nlist = min(len(texts) // 10, 100)  # 클러스터 수 동적 설정
        quantizer = faiss.IndexFlatIP(dimension)
        index = faiss.IndexIVFFlat(quantizer, dimension, nlist, faiss.METRIC_INNER_PRODUCT)
        
        if not index.is_trained:
            index.train(embeddings.astype('float32'))
        index.add(embeddings.astype('float32'))
        
        return index, texts

    def split_text(self, text: str, max_chunk_size: int = 3000) -> List[str]:
        """
        긴 텍스트를 작은 청크로 분할
        """
        words = text.split()
        chunks = []
        current_chunk = []
        current_length = 0
        
        for word in words:
            if current_length + len(word) + 1 > max_chunk_size:
                chunks.append(' '.join(current_chunk))
                current_chunk = [word]
                current_length = len(word)
            else:
                current_chunk.append(word)
                current_length += len(word) + 1
                
        if current_chunk:
            chunks.append(' '.join(current_chunk))
        
        return chunks

    def generate_response(self, prompt: str) -> str:
        """
        TinyLlama를 사용하여 응답 생성
        """
        try:
            # 긴 텍스트 처리를 위해 청크로 분할
            chunks = self.split_text(prompt)
            responses = []
            
            for chunk in chunks:
                # TinyLlama 형식에 맞게 프롬프트 포맷팅
                formatted_prompt = f"<human>: {chunk}\n<assistant>:"
                
                # 토큰화 및 길이 제한
                inputs = self.tokenizer(
                    formatted_prompt,
                    return_tensors="pt",
                    truncation=True,
                    max_length=4096
                )
                
                outputs = self.model.generate(
                    **inputs,
                    max_new_tokens=512,
                    temperature=0.1,
                    top_p=0.95,
                    do_sample=True,
                    pad_token_id=self.tokenizer.eos_token_id
                )
                
                response = self.tokenizer.decode(outputs[0], skip_special_tokens=True)
                response = response.split("<assistant>:")[-1].strip()
                
                if response and response != "없음":
                    responses.append(response)
            
            # 모든 응답 중 가장 관련성 높은 것 선택
            return responses[0] if responses else "없음"
            
        except Exception as e:
            logger.error(f"응답 생성 중 오류 발생: {e}")
            return "오류 발생"

    def find_table_end_sections(
        self,
        texts_by_page: Dict[int, str],
        target_phrases: List[str] = None,
        similarity_threshold: float = 0.75
    ) -> List[Tuple[int, str]]:
        """
        표 끝 부분 검출
        """
        if target_phrases is None:
            target_phrases = [
                "질병관련 특약",
                "상해 및 질병 관련 특약",
                "재진단암진단비 특약"
            ]

        logger.info("페이지 분석 시작")
        pages_text = list(texts_by_page.values())
        index, indexed_texts = self.setup_rag_index(pages_text)
        
        results = []
        for phrase in target_phrases:
            logger.info(f"'{phrase}' 검색 중...")
            phrase_embedding = self.create_embeddings([phrase])
            D, I = index.search(phrase_embedding.astype('float32'), k=5)
            
            for similarity, idx in zip(D[0], I[0]):
                if similarity < similarity_threshold:
                    continue
                    
                page_num = list(texts_by_page.keys())[idx]
                text = texts_by_page[page_num]
                
                prompt = f"""
                다음 텍스트에서 보험 약관의 표 끝을 나타내는 문구를 찾아주세요.
                특히 다음과 같은 형식의 문구를 찾아주세요:
                - "[특약 이름] (자세한 사항은 반드시 약관을 참고하시기 바랍니다.)"

                텍스트:
                {text}

                답변 형식:
                - 문구를 찾은 경우: 전체 문구를 정확히 알려주세요
                - 찾지 못한 경우: "없음"
                """
                
                response = self.generate_response(prompt)
                
                if "없음" not in response and any(phrase in response for phrase in target_phrases):
                    results.append((page_num, response))
                    logger.info(f"페이지 {page_num}에서 표 끝 문구 발견: {response}")

        logger.info("페이지 분석 완료")
        return sorted(results, key=lambda x: x[0])

def main():
    try:
        # PDF 파일 경로 설정
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        
        logger.info("PDF 처리 시작")
        # PDF 텍스트 추출
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            texts_by_page = {i+1: page.extract_text() for i, page in enumerate(reader.pages)}
        logger.info(f"총 {len(texts_by_page)} 페이지 로드됨")
        
        # TableEndDetector 초기화 및 실행
        detector = TableEndDetector()
        results = detector.find_table_end_sections(texts_by_page)
        
        # 결과 출력
        if results:
            print("\n발견된 표 끝 문구:")
            for page_num, text in results:
                print(f"\n페이지 {page_num}:")
                print(text)
        else:
            print("표 끝 문구를 찾지 못했습니다.")
            
    except FileNotFoundError:
        logger.error(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
        raise
    except Exception as e:
        logger.error(f"처리 중 오류 발생: {e}")
        raise

if __name__ == "__main__":
    main()