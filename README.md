# 포스코퓨처엠 스마트 밀폐공간 안전설계 및 온열질환 예방 프로그램

**Smart Confined-space Safety Design & Heat Stress Prevention Program**

밀폐공간 작업 시 필요한 환기량, 국소냉방 열량, 온열질환 예방 프로토콜을
현장 조건 입력만으로 자동 산출하는 스마트 안전설계 프로그램입니다.

## 🔑 주요 기능

- 🌡️ 체감온도 자동 산출 (NOAA Heat Index)
- 📐 밀폐공간 체적 계산 (사각/원통/결합형)
- 💨 법적 필요 송풍기 용량 계산 (Q = V × 0.4)
- ❄️ 국소냉방(Spot Cooling) 필요 열량 계산
- 🚨 폭염대비 온열질환 예방수칙 3단계 자동 안내
- 🛡️ 작업 내용별 맞춤 안전 조치 자동 매칭
- 📊 조회수 카운터 및 결과서 PDF 저장

## 🚀 로컬 실행 방법

```bash
# 1. 저장소 클론
git clone https://github.com/<사용자명>/posco-confined-space-app.git
cd posco-confined-space-app

# 2. 가상환경 생성 (권장)
python -m venv venv
source venv/bin/activate    # Windows: venv\Scripts\activate

# 3. 패키지 설치
pip install -r requirements.txt

# 4. Streamlit 실행
streamlit run app.py
