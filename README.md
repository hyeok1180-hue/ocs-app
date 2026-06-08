# 💊 OCS 처방 분석기

대웅제약 영업사원용 병원 처방통계 자동 분석 웹앱

## 파일 구성
```
ocs_app/
├── app.py                          ← 웹앱 메인 코드
├── requirements.txt                ← 필요 라이브러리 목록
├── 약제급여목록및급여상한금액표_2026_4_1__공개용_1부.xlsx  ← 보험약가 파일
└── README.md
```

## 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 배포
1. GitHub에 이 폴더 전체 업로드
2. share.streamlit.io 접속 → New app
3. 저장소 선택 → app.py 선택 → Deploy
