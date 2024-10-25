# main.py
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"], # Vercel 배포 후에는 Vercel URL도 추가
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/api/analyze")
async def analyze_pdf(file: UploadFile = File(...)):
    # PDF 분석 로직
    return {
        "results": [
            {
                "section": "기본계약",
                "changes": ["보장금액 변경", "보험기간 추가"],
                "status": "completed"
            }
        ]
    }