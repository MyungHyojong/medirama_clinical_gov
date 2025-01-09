# Clinical Trial Data Processor

## 소개

이 애플리케이션은 Streamlit과 OpenAI API를 사용하여 임상시험 데이터를 처리하고, CSV 또는 XLSX 파일로부터 데이터를 분석 및 변환하여 HTML, Excel, 그리고 PPT로 출력할 수 있는 도구입니다. 연구 데이터를 시각화하고 GPT를 활용해 추가 정보를 추출합니다.

https://medirama-clinical-gov.streamlit.app/
---

## 주요 기능

### 데이터 처리 및 변환
- **CSV 파일 업로드**: 임상시험 데이터를 포함한 CSV 파일 업로드.
- **데이터 처리**: 업로드된 데이터를 GPT를 통해 분석하고 결과를 추가.
- **결과 데이터 다운로드**: 처리된 데이터를 Excel 및 HTML 형식으로 다운로드.

### PPT 생성
- **CSV/XLSX 데이터로 PPT 생성**: 업로드된 데이터를 바탕으로 PPT 슬라이드를 생성.
- **열 너비 및 슬라이드 행 개수 설정**: 슬라이드에 표시할 데이터와 형식을 사용자 지정.

### GPT 기반 데이터 분석
- **임상시험 데이터 분석**: GPT를 활용하여 실험 유형, 근거, 설명, 관련 유전자 및 신뢰 점수 추출.

---

## 설치 방법

### 필수 요구 사항
- Python 3.9 이상
- 아래 패키지를 설치:
  ```bash
  pip install streamlit pandas openai python-pptx requests
  ```

### OpenAI API 키 설정
- Streamlit `secrets.toml` 파일을 설정:
  ```toml
  [openai]
  api_key = "your_openai_api_key"
  ```

---

## 사용법
![Clinical Trial Protocol](https://i.ibb.co/SN0ZKdG/i2.png)

### 1) 데이터 업로드 및 처리
1. **CSV 업로드**:
   - CSV 파일을 업로드합니다.
2. **`Process Data` 버튼 클릭**:
   - 데이터가 GPT를 통해 처리되고, 결과가 테이블로 표시됩니다.
3. **결과 다운로드**:
   - 처리된 데이터를 Excel 및 HTML 형식으로 다운로드 가능합니다.

### 2) PPT 생성
1. **CSV/XLSX 파일 업로드**:
   - PPT로 변환할 데이터를 포함한 파일 업로드.
2. **열 및 슬라이드 설정**:
   - 표시할 열, 슬라이드당 행 수, 열 너비를 설정합니다.
3. **PPT 생성**:
   - "Generate PPT" 버튼을 클릭하여 PPT 생성 후 다운로드합니다.

---

## 주요 코드 구성

- **`find_common_substring`**: 두 문자열 간 공통 부분 추출.
- **`extract_last_in_brackets`**: 문자열에서 괄호 안 마지막 텍스트 추출.
- **`highlight_substring`**: 텍스트에서 특정 문자열 강조.
- **`create_ppt_from_dfs`**: DataFrame을 PPT 형식으로 변환.
- **`zip_html_files`**: HTML 콘텐츠를 압축하여 ZIP 파일 생성.
- **`process_files`**: 데이터 처리를 수행하고 결과를 JSON 형식으로 구성.

---

## 주의사항

1. **OpenAI API 사용량**:
   - OpenAI API 호출은 사용량을 소비하므로 적절한 요금제를 선택하세요.

2. **데이터 품질**:
   - 업로드된 데이터가 정확하고 잘 구성되어 있어야 결과 품질이 향상됩니다.

3. **데이터 보안**:
   - 업로드한 파일이 민감한 정보를 포함할 수 있으므로 보안을 유지하세요.

---

## 기여 및 피드백

이 프로젝트에 기여하거나 피드백을 제공하려면 [GitHub 저장소 링크](#)를 통해 문의하세요.

---

## 라이선스

이 프로젝트는 MIT 라이선스에 따라 배포됩니다.
