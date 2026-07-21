# 포스코퓨처엠 스마트 밀폐공간 안전설계 및 온열질환 예방 프로그램

**Smart Confined-space Safety Design & Heat Stress Prevention Program**

밀폐공간 작업 시 필요한 환기량, 환기팬 사용 가능 여부, 국소냉방 열량,
온열질환 예방 프로토콜을 현장 조건 입력만으로 자동 산출하는
스마트 안전설계 프로그램입니다.

## 🔑 주요 기능

- 🌡️ 체감온도 자동 산출 (NOAA Heat Index)
- 📐 밀폐공간 체적 계산 (사각/원통/결합형)
- 💨 **환기량 및 환기팬 가동 시간 자동 계산 (NEW)**
  - 작업 전 초기 급기 시간 산정
  - 적용 환기팬 용량·대수 입력 시 **사용 가능/불가 자동 판정**
- ❄️ 국소냉방(Spot Cooling) 필요 열량 계산
- 🚨 폭염대비 온열질환 예방수칙 3단계 자동 안내
- 🛡️ 작업 내용별 맞춤 안전 조치 자동 매칭
- 📊 조회수 카운터 및 결과서 PDF 저장

## 🧮 핵심 계산식

- 체적: V = W × L × H  또는  (π × D² / 4) × H
- 작업 중 법적 필요 환기량: Q = V × 0.4 (m³/min)
- 작업 전 총 필요 환기량: V × 10 (m³)
- 최소 가동 시간: (V × 10) ÷ 적용 환기팬 총 용량
- 사용 불가 판정: 적용 총 용량 < 법적 필요 환기량

## 🚀 로컬 실행 방법

```bash
git clone https://github.com/<사용자명>/posco-confined-space-app.git
cd posco-confined-space-app

python -m venv venv
source venv/bin/activate    # Windows: venv\Scripts\activate

pip install -r requirements.txt
streamlit run app.py
