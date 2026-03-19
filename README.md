# Excel to Confluence 변환기

Excel (`.xlsx`) 파일을 Confluence에 임포트할 수 있는 Word (`.docx`) 문서로 변환합니다.

여러 시트에 걸친 **테이블**, **이미지**, **병합 셀**, **텍스트** 를 모두 지원합니다.

---

## 동작 방식

```
Excel (.xlsx)
  └── Sheet 1  ──►  제목 + 텍스트 + 테이블 + 이미지
  └── Sheet 2  ──►  제목 + 텍스트 + 테이블 + 이미지
  └── Sheet 3  ──►  ...
                          │
                          ▼
                    output.docx
                          │
                          ▼
              Confluence Word 임포트
```

각 시트는 Word 문서의 섹션으로 변환되며 다음 요소를 보존합니다:
- 셀 텍스트 및 폰트 스타일 (굵기, 기울임, 색상)
- 병합 셀 및 배경색을 포함한 테이블 구조
- 원래 행 순서에 맞춰 삽입된 이미지

---

## 요구 사항

- Python 3.9 이상
- `requirements.txt` 에 명시된 pip 패키지

---

## 설치

```bash
git clone https://github.com/pineskyeo/excel-to-confluence.git
cd excel-to-confluence

pip install -r requirements.txt
```

---

## 사용법

### 1. Excel 파일 변환

```bash
python converter.py input.xlsx
```

변환된 파일은 동일한 디렉토리에 `input.docx` 로 저장됩니다.

출력 경로를 직접 지정하려면:

```bash
python converter.py input.xlsx output.docx
```

### 2. Confluence에 임포트

1. Confluence 페이지 편집기를 열고 **편집** 클릭
2. **···** (더 보기) 메뉴 클릭
3. **Word 문서 임포트** 선택
4. 생성된 `.docx` 파일 업로드

---

## 샘플 파일로 빠르게 테스트

텍스트, 테이블, 병합 셀, 이미지가 포함된 샘플 Excel 파일을 생성합니다:

```bash
python create_sample.py
```

`sample.xlsx` 파일이 생성되며 3개의 시트로 구성됩니다:

| 시트 | 내용 |
|------|------|
| Project Overview | 제목 텍스트 + 이미지 + 업무 테이블 |
| Technical Specs | 병합 셀 헤더 + 이미지 + 테이블 2개 |
| Summary | KPI 텍스트 + 이미지 + 요약 테이블 |

이후 변환 실행:

```bash
python converter.py sample.xlsx
```

---

## 프로젝트 구조

```
excel-to-confluence/
├── converter.py        # 메인 변환기
├── create_sample.py    # 테스트용 샘플 Excel 생성기
├── requirements.txt    # Python 의존성 목록
└── README.md
```

---

## 지원하는 Excel 요소

| 요소 | 지원 여부 |
|------|-----------|
| 텍스트 셀 | ✅ |
| 굵기 / 기울임 폰트 | ✅ |
| 폰트 색상 | ✅ |
| 셀 배경색 | ✅ |
| 테이블 (이름 지정) | ✅ |
| 테이블 (자동 감지) | ✅ |
| 병합 셀 | ✅ |
| 이미지 삽입 (PNG/JPEG) | ✅ |
| 다중 시트 | ✅ |
| 수식 | ✅ (계산된 값만 변환) |
| 차트 | ❌ (변환 제외) |
