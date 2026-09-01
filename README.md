# 올리브영 올영세일(브세) TOP100 브랜드 순위 리포트

올리브영 올영세일(3월/6월/9월/12월) 기간 동안의 TOP100 브랜드 순위를 엑셀에서 읽어와
탭으로 탐색 가능한 HTML 리포트로 만들고, GitHub Pages에 자동 배포하는 개인용 리포팅 도구입니다.

**라이브 리포트:** https://kimco12.github.io/oliveyoung-brand-sale-rank-pages/

> 이 저장소(oliveyoung-brand-sale-rank)는 비공개(private)이며 소스코드/엑셀 로직 백업용입니다.
> 실제 리포트는 index.html만 담긴 별도의 공개 저장소 `oliveyoung-brand-sale-rank-pages`로
> 배포되어 위 링크로만 결과물을 볼 수 있고, 소스코드는 노출되지 않습니다.
>
> 여기 담긴 순위 데이터는 AI 스크립트가 엑셀을 그대로 파싱해 자동 생성한 결과이니,
> 실제 의사결정에 쓰기 전에는 원본 엑셀과 대조해 검토하세요.

## 파일 구성

| 파일 | 역할 |
|---|---|
| `generate_brand_rank_report.py` | 엑셀 → HTML 리포트 생성 + GitHub 커밋/Pages 배포까지 한 번에 처리하는 메인 스크립트 |
| `_brand_rank_report_template.html` | 리포트의 HTML/CSS/JS 템플릿 (`generate_brand_rank_report.py`가 이 파일을 읽어 데이터를 주입) |
| `index.html` | 위 스크립트로 생성된 최종 리포트 (GitHub Pages가 이 파일을 서비스) |
| `oliveyoung_brand rank wide.py` | 올리브영 모바일 랭킹 페이지(`m.oliveyoung.co.kr`)를 Selenium으로 크롤링해 브랜드 실시간 순위를 엑셀에 가로형(wide)으로 누적 저장하는 별도 수집 스크립트 |

`oliveyoung_brand rank wide.py`는 실시간 랭킹을 수집하는 **독립적인 도구**이며,
현재 `generate_brand_rank_report.py`가 읽는 브세 TOP100 엑셀과 자동으로 연동되어 있지는 않습니다
(수집된 데이터를 리포트에 반영하려면 브세 엑셀에 수동으로 옮겨야 합니다).

## 리포트 기능

- **기간별 탭**: 세일 기간(예: 26년 3월/6월/9월)마다 브랜드 × 일자 표로 순위 확인, 전일 대비 변동(▲상승 ▼하락 – 동일 NEW신규진입) 배지 표시
- **검색 / TOP N 필터**: 브랜드명 검색, TOP 10~전체(이탈 브랜드 포함) 필터
- **이미지로 다운로드**: 현재 필터와 무관하게 TOP100 전체를 표 이미지(PNG)로 다운로드
- **기간 비교 탭**: 두 세일 기간을 골라 브랜드별 "기간 평균 순위"로 비교. 신규진입/이탈 카드를 클릭하면 해당 브랜드만 표에 필터링됨

## 사용 방법 (엑셀 데이터가 갱신됐을 때)

1. 같은 폴더의 `올리브영 26년 브세 top100_k.xlsx`를 최신 데이터로 갱신
2. `generate_brand_rank_report.py` 실행 → 로컬 HTML 재생성 + GitHub Pages 자동 재배포

```bash
python generate_brand_rank_report.py
```

GitHub 인증 토큰은 코드에 포함되어 있지 않으며, 로컬의 `github_config_brandrank.json` 또는
환경변수 `GITHUB_TOKEN`에서 읽습니다(이 파일은 저장소에 올라가지 않습니다).

## 엑셀 레이아웃 인식 규칙

`generate_brand_rank_report.py`는 엑셀 시트 구조가 바뀌어도 아래 규칙만 지키면 자동으로 새 기간 블록을 인식합니다.

- 셀 값이 정확히 **"순위"**인 위치를 블록의 기준점으로 삼음
- 제목: 같은 열에서 기준행 위쪽 5행 이내에 있는 첫 텍스트 셀
- 날짜: 기준행 바로 위 행, 순위 열 오른쪽 7칸
- 데이터: 기준행 다음 행부터 순위 1~100, 순위 열 오른쪽 7칸이 요일별(일~토) 브랜드
