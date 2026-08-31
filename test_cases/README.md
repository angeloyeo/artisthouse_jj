# 리포트 입력 테스트 케이스

이 폴더는 `inputTemplate.xlsx`의 컬럼 구조를 기준으로 만든 테스트용 XLSX 파일과 실행 결과 보고서를 보관합니다.

## 구성

- `inputs/`: 브라우저에서 직접 불러볼 테스트 입력 파일
- `test_results.xlsx`: 현재 `app.js`의 검증/정규화 로직 기준 실행 결과 요약
- `generate_test_cases.py`: 테스트 파일과 결과 보고서를 다시 생성하는 스크립트

## 재생성

프로젝트 루트에서 아래 명령을 실행합니다.

```powershell
$env:PYTHONIOENCODING='utf-8'; py test_cases\generate_test_cases.py
```

## 주요 케이스

- `tc01_valid_progress.xlsx`: 정상 3개월 성장 데이터
- `tc02_long_evaluation.xlsx`: 종합평가가 매우 긴 경우
- `tc03_scores_out_of_range.xlsx`: 점수가 0~100 범위를 벗어나는 경우
- `tc04_missing_student_info.xlsx`: 학생 기본 정보가 모두 비어 있는 경우
- `tc05_missing_required_column.xlsx`: 필수 컬럼이 빠진 경우
- `tc06_invalid_dates.xlsx`: 리포트시점이 해석 불가한 경우
- `tc07_blank_last_identity_fallback.xlsx`: 마지막 행의 이름/학교/학년이 비어 있는 경우
- `tc08_zero_detail_scores.xlsx`: 실기 세부항목 점수가 0인 경우
- `tc09_text_and_comma_numbers.xlsx`: 문자열 숫자와 콤마 포함 숫자
- `tc10_unsorted_dates.xlsx`: 행 순서상 날짜가 역순인 경우
- `tc11_header_only.xlsx`: 헤더만 있고 데이터가 없는 경우
- `tc12_wrong_univ_template_marker.xlsx`: `성적` 컬럼이 있는 잘못된 템플릿
- `tc13_missing_optional_current_school.xlsx`: 현재 지원 가능 학교 컬럼이 빠진 경우
- `tc14_score_units.xlsx`: 점수에 `점` 단위가 붙어 있는 경우
- `tc15_non_numeric_scores.xlsx`: 숫자 칸에 숫자가 전혀 없는 문자를 넣은 경우
