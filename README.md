# Trip.com 무인·자동화 호텔 vs 일반 유인호텔 리뷰 데이터셋 크롤링

연구 목적: 무인·자동화 호텔과 일반 유인호텔에 대해 Trip.com의 실제 투숙객 리뷰와 리뷰 첨부 사진을 수집한다. 이를 통해 서비스 운영 방식의 차이가 고객 경험·평가에 어떻게 나타나는지 분석할 수 있는 데이터셋을 만든다. Survey나 지각 변수(PU, PEOU 등)는 사용하지 않는다.

---

## 0-A. 2026-10-06 live 실행에서 확인된 사실과 코드 변경 (아래 본문보다 우선)

실제 Trip.com을 대상으로 probe와 전체 실행을 하면서 확인한 내용이다. 아래 1~12절 중 이 절과 다른 서술은 이 절이 우선한다. 최종 수치는 `outputs/05`, `06`, `07` 파일을 본다.

| 항목 | live에서 확인된 사실 | 처리 |
|---|---|---|
| 리뷰 수 상한 | 공개 리뷰 페이지(`…/review.html`)는 익명 방문자에게 "Show More"를 5페이지(15개 × 5 = 75개)까지만 제공한다 | 우회하지 않는다. 호텔당 최신순 최대 75개를 수집하고, 리뷰가 75개보다 많은 호텔은 `PARTIAL / VISIBLE_LIMIT_REACHED`로 기록한다 |
| 리뷰 JSON | 정렬·페이지 넘김 시 페이지가 `getClientHotelCommentList`를 호출한다 (`orderBy=1`이 Most Recent) | 응답을 수신만 해서 파싱한다. API를 직접 호출하지 않는다 |
| 정렬·더보기 UI | 정렬은 `Sort by:` 옆 읽기 전용 input 드롭다운, 더보기는 파란색 "Show More" 버튼이다. 리뷰 카드마다 같은 글자의 본문 펼치기 링크가 있다 | `config.SORT_OPEN_SELECTORS`, `LOAD_MORE_BUTTON_TEXTS`와 `crawl_tripcom_reviews._click_load_more_button`로 처리한다 |
| 표본 정의 | 첫 화면에는 "Most relevant" 순서의 15개가 먼저 들어 있다 | 최신순 응답으로 받은 리뷰만 표본(`04`)에 넣는다. 기본 순서로만 본 리뷰는 `work/reviews/*.jsonl`에만 남는다(`newest_rank`가 비어 있음) |
| 상세·검색 페이지 | 자동화 브라우저로 호텔 상세 페이지와 `/hotels/list`를 열면 로그인 페이지로 이동한다 | 우회하지 않는다. 호텔 메타와 시설 정보는 공개 리뷰 페이지에 포함된 데이터에서 읽는다 (`common.parse_review_page_meta`) |
| 로봇 검증 근거 | 리뷰 페이지 데이터에 시설 ID 전체 목록과 상위 시설 태그 이름이 있다. ID 364가 "Service robots"다 | 무인 후보는 364 유무로 live 재확인하고, 유인 후보는 364가 없어야 통과한다. 추가로 수집된 리뷰 본문의 로봇 언급 수를 `qa_final.py`로 점검한다 |
| 가격 | 날짜를 지정한 상세 페이지를 쓸 수 없다 | 리뷰 페이지의 `displayPrice`(세전 최저가, USD, 크롤링 당일 기본 체크인 날짜)를 모든 호텔에 동일하게 쓴다. +28일 고정 조건은 적용되지 않았다 |
| 중국 유인 후보 discovery | 키워드 검색 목록은 로그인 벽 뒤에 있다 | 공개 도시 목록 페이지(`/hotels/{city}-hotels-list-{id}/`)와 그 지역(zone) 페이지에서 후보 링크를 모은 뒤, 후보마다 리뷰 페이지에서 검증한다. 같은 브랜드 우선 검색은 불가능해졌다 |
| 리뷰 사진 URL | 리뷰 객체의 `imageList`는 `/0230….jpg` 형태의 상대경로다 | `config.REVIEW_IMAGE_BASE`를 붙여 원본 해상도로 받는다 |
| crawl state | 재선정·재매칭으로 hotel_id가 바뀔 수 있다 | `work/crawl_state.json`을 Trip.com 호텔 코드 기준으로 저장한다 |

추가된 파일: `qa_final.py` (최종 QA, `outputs/07_final_qa.xlsx` 생성). `_prefetch_seed_meta.py`, `_prediscover.py`, `_discover_star_fix.py`는 이번 실행에서 대기 시간을 줄이려고 쓴 보조 스크립트이며, 같은 일을 `match_controls.py`가 단독으로도 수행한다.

### live 실행 결과 (2026-10-06)

| 항목 | 값 |
|---|---|
| 무인/자동화 호텔 | 100 (PHYSICAL_AI 91, AUTOMATION_ONLY 9). 중국 76, 일본 22, 한국 1, 미국 1 |
| Trip.com "Service robots" 태그 live 확인 | PHYSICAL_AI 91개 중 83개. 나머지 8개(Henn na)는 공식 근거만 있어 confidence MEDIUM |
| 유인 호텔 | 100 (기존 후보풀 22, live discovery 78). 전부 Trip.com 로봇 태그 없음 |
| pair | 100쌍 / 200개 호텔. 전부 같은 도시, 성급 동일 61쌍, ±1 39쌍, 중복 호텔·코드 0 |
| 크롤링 상태 | SUCCESS 1, PARTIAL 199, FAILED 0 |
| PARTIAL 사유 | 194개는 75개 상한(`VISIBLE_LIMIT_REACHED`), 5개는 Trip.com이 나열하는 리뷰가 75개 미만(`ALL_LISTED_REVIEWS_COLLECTED`) |
| 리뷰 | 14,701 (무인/자동화 7,375, 유인 7,326). 전부 최신순, 원문 |
| 사진 | 사진 있는 리뷰 4,352, 다운로드 13,069장, 실패 0 |

- 선정 규칙 보정: live 확인 후에는 generic 태그 호텔이 모두 HIGH가 되어, 기존 규칙대로면 초과분 제거가 공식 근거 호텔(AUTOMATION_ONLY 포함)로 넘어갔다. generic 태그 호텔을 confidence와 무관하게 먼저 제거하도록 단계를 추가했다(`validate_hotels.balance_select`). 그 결과 reserve 9개는 JI Hotel 7개(항저우 5, 상하이 1, 청두 1), Orange Hotel 쑤저우 1개, Henn na Hotel Osaka Shinsaibashi 1개다.
- 유인 후보에서 이름으로 제외한 유형: 아파트·민박·호스텔, smart·무인·영화관 컨셉 (`config.NON_CONTROL_NAME_RE`).
- 사진 저장 위치: 원본이 평균 1MB를 넘어 총 13.8GB다. C: 여유 공간이 부족해 실제 파일은 `E:\tripcom_hotel_review_pipeline_data\images`에 있고, `dataset/images`는 그 폴더로 연결된 junction이다. `04`의 `photo_N` 상대경로는 그대로 유효하다. 프로젝트 폴더를 다른 PC로 옮길 때는 E:의 폴더를 함께 복사해야 한다.
- 그룹 간 언어 구성이 다르다(무인/자동화 쪽 중국어 비중이 더 높음). 분석에서 통제할 것.

## 0. 이전 상태 (2026-10-05 오프라인 실행 결과, 참고용)

| 단계 | 상태 |
|---|---|
| 무인/자동화 100 선정 | **완료 (desk evidence 기준)**: PHYSICAL_AI 91, AUTOMATION_ONLY 9. Reserve 9 |
| 유인 100 매칭 | **25 / 100 잠정 매칭** (일본 23, 한국 1, 미국 1: SAME_CITY 24, NEARBY_CITY 1) |
| 중국 75 pair | 기존 후보풀에 중국 유인호텔이 0개라서 **live discovery 필요** |
| 리뷰·사진 크롤링 | **미실행**. 작성 환경에서 외부 네트워크가 차단되어 있음 |

로컬 PC에서 `python run_all.py`를 실행하면 아래가 진행된다. 각 단계의 Resume 기능이 그대로 작동한다.

1. live 재검증
2. 중국 유인호텔 discovery
3. 100 pair 확정
4. 리뷰·사진 수집
5. 로그·요약 생성

---

## 1. 설치

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

두 원본 Excel은 `inputs/` 폴더에 둔다. 코드는 이 파일을 읽기만 하고 절대 덮어쓰지 않는다.

- `호텔_표본_재구성_PhysicalAI100_매칭계획.xlsx`
- `유인호텔_리스트.xlsx`

## 2. 실행 순서

```bash
python run_all.py                  # 전체를 한 번에 끝까지 실행 (단계별 확인 없음)
```

본 실행 전에 Trip.com 페이지 구조가 파서와 맞는지 먼저 확인하는 것을 권장한다. 익명 방문자가 받을 수 있는 리뷰 수도 이때 확인된다.

```bash
python run_all.py --offline                      # 몇 초: 선정 + 잠정 매칭 (pairs_current.csv 생성)
python crawl_tripcom_reviews.py --probe H076     # 호텔 1곳 진단 (브라우저 창 표시)
```

`--probe`는 캡처된 JSON URL, 파싱된 리뷰 수와 샘플, 정렬·페이지 넘김 결과를 출력한다. 결과는 `work/probe/`에만 저장되고 실제 데이터셋과 crawl state는 건드리지 않는다. 리뷰가 0개로 나오면 `work/debug/probe_*.html`과 `*_xhr.txt`를 확인해 `config.py`의 셀렉터나 `REVIEW_RESPONSE_PATTERNS`를 조정한다.

단계별 실행도 가능하다. 각 스크립트는 독립 실행되며 resume-safe하다.

| 순서 | 스크립트 | 하는 일 |
|---|---|---|
| 1 | `validate_hotels.py` | 후보 109개를 읽고 desk + live Trip.com 시설 정보로 분류한 뒤 정확히 100개를 선정 |
| 2 | `match_controls.py` | 유인 후보풀을 검증하고 중국 도시는 discovery를 실행한 뒤 Hungarian 1:1 매칭 |
| 3 | `crawl_tripcom_reviews.py` | 호텔당 최신 리뷰 최대 100개 수집 |
| 4 | `download_review_images.py` | 리뷰 첨부 사진을 리뷰당 최대 10장 다운로드하고 `04` CSV 생성 |
| 5 | `make_summary.py` | `05` 크롤링 로그, `06` 표본 요약, 터미널 Summary 출력 |

옵션:

- `--offline`: 네트워크 없이 선정과 잠정 매칭만 수행
- `--skip-discovery`: discovery 생략
- `--reselect`: live 선정 결과를 다시 계산
- `--retry-partial`: PARTIAL 호텔도 재시도

## 3. 호텔 분류 기준

| operational_type | 기준 |
|---|---|
| PHYSICAL_AI | 실제 물리적 로봇이 고객 접점 서비스를 수행: 프런트·리셉션·배송·안내·식음료 배송 로봇 |
| AUTOMATION_ONLY | 서비스 로봇은 미확인이지만 체크인·체크아웃·프런트의 핵심 접점이 자동화됨: 키오스크, 홀로그램 프런트, 무인 리셉션, 모바일 셀프 체크인. 객실 태블릿만 있는 경우는 제외 |
| CONVENTIONAL | 직원 중심 서비스. 로봇과 핵심 프런트 자동화가 모두 없음 |
| UNCERTAIN | 근거가 충돌하거나 부족함. 두 그룹 모두에서 제외 |

- 체인 이름만으로 판정하지 않는다. 예를 들어 Henn-na는 지점별로 PHYSICAL_AI 16개, AUTOMATION_ONLY 9개(홀로그램 또는 Express 체크인 기계)로 나뉜다.
- Trip.com 시설 라벨 "Express check-in/check-out"은 보통 직원이 처리하는 빠른 체크인을 뜻한다. 그래서 약한 신호로만 기록하고 자동화 근거로 쓰지 않는다.
- live 판정은 페이지 전체 텍스트가 아니라 시설(facility) 영역 텍스트만 사용한다. 리뷰 본문에 "robot"이 나와도 분류에 반영되지 않는다.

**confidence**

| 등급 | 기준 |
|---|---|
| HIGH | 지점 단위의 최신 1차 근거(공식 지점 페이지·보도자료), 이번 검토 주기에 확인한 Trip.com 시설, 또는 desk 근거와 live 근거의 일치 |
| MEDIUM | 이전 목록에서 이월된 OTA 태그(미재검증), 또는 간접적인 자동화 근거 |
| LOW | 체인 수준 근거뿐이거나 근거가 충돌 |

현재(오프라인) 기준으로 HIGH 29, MEDIUM 71이다. 이 중 AUTOMATION_ONLY는 HIGH 4, MEDIUM 5다. Express 지점 2개처럼 브랜드 공지만 근거인 경우는 MEDIUM으로 두었다. 중국 기존 66개는 원본 파일의 Trip.com "Service robots" 태그를 이월한 것이라 MEDIUM이다. live 실행에서 태그가 다시 확인되면 HIGH로 올라간다. 확인되지 않으면 AUTOMATION_ONLY나 UNCERTAIN으로 바뀌고, 그 경우 reserve가 빈자리를 채운다.

## 4. 무인/자동화 100 선정 규칙

- 후보는 `Physical_AI_Final_100` 100개와 `Excluded_9` 9개다. 기존 설계에서 "Physical AI 아님"으로 빠졌던 Henn-na 9개 지점은 새 3분류에서 AUTOMATION_ONLY 후보가 된다.
- 적격 후보는 109개로 100개보다 9개 많다. 그래서 브랜드×도시 셀 집중도가 가장 높은 곳부터 1개씩 제거했다.
  - 제거 대상은 근거가 generic "Service robots" 태그뿐이고 confidence가 HIGH가 아닌 호텔로 한정했다.
  - 동률이면 도시 전체 수가 큰 곳, 그다음 원본 행 번호가 뒤인 호텔을 먼저 제거한다.
- 제거된 9개는 유효한 호텔이므로 RESERVE로 보존한다. 매칭할 유인호텔이 없는 무인호텔이 생기면 같은 국가의 reserve로 교체하고, 교체 내역을 기록한다.
  - JI Hotel 항저우 5개, 상하이 1개, 청두 1개, 난징 1개
  - Orange Hotel 쑤저우 1개
- 결과 브랜드 비율(무인/자동화 100 기준). 브랜드 편중이 있지만 지시대로 제거하지 않고 비율로 명시한다.

| 브랜드 | 비율 |
|---|---|
| JI Hotel | 34% |
| Henn na Hotel | 25% |
| Atour | 13% |
| Orange | 12% |
| 독립·기타 | 16% |

- 국가 구성: 중국 75, 일본 23, 한국 1, 미국 1

## 5. 유인호텔 매칭 규칙

1. 국가가 다르면 절대 매칭하지 않는다.
2. 우선순위는 같은 도시 > 같은 브랜드(또는 같은 모회사) > 동일 성급 > 성급 ±1 > 유사 가격 > 유사 리뷰 수다. 이 순서가 사전식(lexicographic)으로 지켜지도록 가중치 크기를 설계했다(`config.MATCH_WEIGHTS`).
3. 같은 도시에 후보가 없으면 같은 광역권(`NEARBY_CITIES`, 예: Komatsu→Kanazawa)을 사용한다. 최후 수단은 같은 국가(SAME_COUNTRY_ONLY)이며, 이 경우 플래그를 남긴다.
4. 전체 100 pair를 동시에 최적화한다(Hungarian 알고리즘). 같은 유인호텔은 중복 사용하지 않는다.
5. 유인 후보도 live로 검증한다. Trip.com 시설에 로봇이나 셀프체크인 신호가 있으면 탈락시키고 로그에 남긴다. 무인 후보 109개(reserve 포함)는 어떤 경우에도 대조군으로 쓰지 않는다.
6. 중국 discovery는 Trip.com 도시 목록 페이지를 사용한다.
   - 무인호텔과 같은 브랜드 키워드로 먼저 검색한다(예: 같은 도시의 로봇 없는 JI Hotel 지점). 그다음 일반 목록을 검색한다.
   - 찾은 후보는 각각 상세 페이지에서 로봇과 자동화가 없는지 검증한다.
7. 가격 조건은 모든 호텔에 동일하다.
   - 첫 실행일 기준 +28일 체크인, 1박, 성인 2명, 객실 1개, USD 요청
   - 날짜는 `work/run_config.json`에 고정되어 Resume 시에도 같은 조건을 쓴다.
   - 가격을 확보하지 못하면 빈칸으로 둔다. 추정하지 않는다.

discovery가 막히거나 후보가 부족하면 `inputs/manual_control_candidates_TEMPLATE.xlsx`를 사용한다.

1. 행을 채우고 EXAMPLE 행을 삭제한다.
2. `inputs/manual_control_candidates.xlsx`로 저장한다.
3. `match_controls.py`를 다시 실행한다. 수동 후보도 live 검증을 거친다.

## 6. 리뷰 크롤링 방식

- Trip.com 리뷰 API(`getHotelCommentInfo`)는 직접 호출하거나 요청을 재생하면 HTTP 430으로 차단된다. 그래서 API를 직접 호출하지 않는다.
- 공개 리뷰 페이지를 실제 브라우저(Playwright)로 열고, 페이지가 스스로 받는 JSON 응답을 수신만 해서 파싱한다.
- 화면의 정렬 메뉴에서 "Most recent"를 선택한 뒤 다음 페이지 버튼을 한 번씩 클릭한다. 클릭 사이에는 3~7초 랜덤 지연을 둔다.
  - 정렬 메뉴를 찾지 못하면 최대 300개를 수집한 뒤 날짜 기준 최신 100개를 사용한다.
  - 호텔별 `sort_mode`(NEWEST_UI 또는 DEFAULT_ORDER)를 로그에 남긴다.
- review_text는 원문 그대로 저장한다. 번역·요약·교정은 하지 않는다. 원문 필드(`originalContent`)를 표시용 번역문보다 우선한다. 원문 필드가 없고 번역 플래그만 있으면 `text_source=possibly_translated`로 표시한다.
- 로봇이나 무인 관련 키워드 유무와 관계없이 모든 리뷰를 수집한다.
- 호텔 답글은 리뷰로 섞이지 않게 별도 컬럼(`hotel_reply_text`)에 저장한다.
- CAPTCHA, 로그인 요구, HTTP 403/429/430이 발생하면 우회하지 않는다. 해당 호텔을 FAILED 또는 PARTIAL로 기록하고 다음 호텔로 넘어간다.
  - 3곳 연속으로 차단되면 10~15분 대기한다.
  - 대기를 반복해도 차단이 계속되면 깔끔하게 중단한다. 나중에 다시 실행하면 이어서 진행된다.
- 처리 속도 제어: 호텔은 순차 처리한다. 호텔 사이에 8~18초, 10개 호텔(배치)마다 1~2분 쉰다. 모두 `config.py`에서 조정할 수 있다.

## 7. 리뷰 사진

- 리뷰 객체 안의 이미지 목록만 사용한다. 호텔 갤러리나 객실 판매 사진은 리뷰 객체에 없으므로 구조적으로 섞이지 않는다.
- 리뷰당 최대 10장을 Trip.com 표시 순서대로 저장한다. `image_count`에는 실제 첨부 수(10 초과 가능)를, `photos_downloaded`에는 저장된 수를 기록한다.
- 원본 해상도를 우선 요청한다(CDN의 리사이즈 접미사 제거). 실패하면 Trip.com이 준 URL 그대로 받는다.
- 파일은 리사이즈·크롭·재인코딩 없이 그대로 저장하고, 확장자는 content-type을 따른다(jpg/png/webp 등).
- 경로: `dataset/images/{pair_id}/{group}/{hotel_id}/{review_id}_{NN}.{ext}`
- 같은 URL은 한 번만 다운로드한다. 다른 리뷰에 같은 URL이 나오면 로컬에서 파일을 복사하므로, 파일명은 항상 해당 리뷰를 가리킨다.
- 사진 하나가 실패해도 리뷰는 유지된다. 실패 내역은 `work/image_manifest.csv`(status=FAILED, error)에 기록된다.

## 8. 출력 파일

| 파일 | 내용 |
|---|---|
| `outputs/01_final_unmanned_automated_100.xlsx` | 무인/자동화 100 (spec 필드 + 분류 근거, live 키워드, QA 플래그) |
| `outputs/02_final_conventional_100.xlsx` | 유인 100. 100개가 모두 매칭되기 전에는 `02_conventional_matched_DRAFT.xlsx`로 저장 |
| `outputs/03_final_hotel_pairs_100.xlsx` | 100 pair / 200 호텔 (long, wide, match_quality). 미완성 시 `03_hotel_pairs_DRAFT.xlsx` |
| `outputs/04_tripcom_reviews_raw.csv` | 리뷰 1행. spec 컬럼, `photo_1..10`(상대경로), `photo_url_1..10`, 부가 컬럼 |
| `outputs/05_crawling_log.xlsx` | 호텔별 목표·수집·사진·상태·재시도·실패 사유 |
| `outputs/06_sample_summary.xlsx` | 그룹별 요약 (hotel_level 시트에 대한 Excel 수식, 국가·도시·브랜드·성급 분포) |
| `outputs/stages/S1_unmanned_candidates_validated.xlsx` | 후보 109개 전부의 분류·선정·제외·reserve 사유, QA 플래그, 브랜드 집중도 |
| `outputs/stages/S1_control_pool_validated.xlsx` | 유인 후보풀 검증 결과, 탈락 사유, 도시별 커버리지 |
| `outputs/stages/S2_matching_log.xlsx` | 선택된 pair의 점수 구성요소, pair별 상위 대안 4개, 미매칭·교체 내역 |
| `work/` | resume 상태: crawl_state.json, reviews/*.jsonl, raw_responses/, hotel_meta/, image_manifest.csv, debug/ |

`04`의 부가 컬럼: `photos_downloaded`, `photos_failed`, `text_source`, `review_title`, `rating_scale`, `stay_date`, `hotel_reply_text`, `reviewer_extra`(Trip.com이 제공하는 추가 reviewer 정보), `crawl_sort_mode`, `crawled_at`

Methodology 재현성을 위해 남는 기록:

- 원본 JSON 응답 전체(`work/raw_responses/`)
- 모든 분류·제외·매칭 근거(S1, S2)
- 고정된 가격 조건(`run_config.json`)

## 9. Resume 방법

중단되면 같은 명령을 다시 실행하면 된다.

- 리뷰는 수집 즉시 `work/reviews/{Trip.com 호텔 코드}.jsonl`에 추가된다. 이미 있는 `review_id`는 다시 저장하지 않는다. hotel_id가 아닌 Trip.com 코드로 저장하므로, 재선정이나 재매칭으로 번호가 바뀌어도 리뷰가 다른 호텔에 붙지 않는다.
- 재매칭으로 pair 폴더가 바뀐 사진은 다시 다운로드하지 않고 새 폴더로 로컬 복사한다.
- 호텔이 끝날 때마다 `work/crawl_state.json`을 원자적으로 저장한다. SUCCESS 호텔은 건너뛴다.
- 이미 디스크에 있는 이미지와 manifest에 OK로 기록된 이미지는 다시 받지 않는다.
- live로 확정된 선정 결과는 고정되어, 재실행해도 hotel_id가 바뀌지 않는다. 다시 선정하려면 `--reselect`를 사용한다.
- 호텔 상세 메타(성급, 가격 등)는 `work/hotel_meta/`에 캐시되므로 다시 요청하지 않는다.

## 10. 실패한 호텔 처리

| 상태 | 의미 | 재실행 시 |
|---|---|---|
| SUCCESS | min(100, Trip.com 총 리뷰 수) 이상 수집 | 건너뜀 |
| PARTIAL | 일부만 수집 | 사유가 일시적(TIMEOUT 등)이면 자동 재시도. 그 외는 `--retry-partial` 사용 |
| FAILED | 0개 수집 | 자동 재시도 (최대 `MAX_RETRIES`=2회) |

failure_reason 코드:

| 코드 | 의미 |
|---|---|
| CAPTCHA, LOGIN_REQUIRED, HTTP_403/429/430 | 접근 제한. 우회하지 않음 |
| NO_NEXT_PAGE_CONTROL, PAGINATION_EXHAUSTED | 익명 방문자에게 더 보여주는 리뷰가 없음 |
| NO_REVIEW_PAYLOAD | 리뷰 JSON을 찾지 못함 |
| TIMEOUT, ERROR | 시간 초과 또는 기타 오류 |
| NO_REVIEWS_ON_TRIPCOM | Trip.com에 리뷰가 0개 |

진단 방법:

- 실패 호텔의 HTML, 캡처된 XHR 목록, 스크린샷이 `work/debug/`에 저장된다.
- `--probe {hotel_id}`로 개별 호텔을 진단할 수 있다.
- Trip.com UI 문구가 바뀌어 정렬·다음 버튼을 못 찾으면 `config.py`의 `*_SELECTORS`만 수정하면 된다.

pair 단위 처리:

- 한쪽 호텔이 FAILED여도 pair와 다른 쪽 데이터는 유지된다. 분석에서 pair를 제외할지는 `05` 로그로 판단한다.
- 유인호텔이 끝내 매칭되지 않은 무인호텔은 같은 국가의 reserve로 자동 교체된다(S2 `substitutions` 시트).

## 11. 알려진 제약과 주의사항

- **익명 접근 시 리뷰 수 제한 가능성**: 2026년 기준 공개 페이지는 익명 방문자에게 리뷰를 한 묶음(약 10~15개 보고됨)만 주고 페이지네이션이 막혀 있을 수 있다. 그렇다면 호텔당 100개 목표는 달성되지 않고 대부분 PARTIAL(NO_NEXT_PAGE_CONTROL)로 기록된다. 첫 `--probe` 결과로 확인하는 것을 권장한다. 로그인 우회는 구현하지 않았다.
- **평점 스케일**: 리뷰 평점은 Trip.com의 0~10 스케일이다(`rating_scale` 컬럼). 호텔 overall rating은 화면 표시 스케일을 따른다.
- **언어 편차**: 중국 호텔 리뷰는 중국어 비중이 높을 가능성이 크다. 그룹 간 언어 구성 차이(무인 쪽은 중국 75%)를 분석에서 통제해야 한다.
- **QA 플래그 2건** (S1 `qa_flags` 시트)
  - Henn-na Seoul Myeongdong: evidence URL과 상세 URL의 Trip.com ID가 다르다(71975183 ≠ 71590703). live 실행 전에 맞는 ID를 확인할 것.
  - Henn-na Express Osaka Namba Nipponbashi: Trip.com URL slug가 `under-preparation`이다. 리뷰가 적을 수 있다.
- **연구윤리**
  - reviewer_name과 리뷰 사진(얼굴이 포함될 수 있음)은 개인정보다. `config.PSEUDONYMIZE_REVIEWER_NAME=True`로 가명화할 수 있다.
  - Trip.com 이용약관과 소속 기관의 IRB·데이터 보호 규정을 확인한 뒤 실행할 것.

## 12. 주요 설정 (`config.py`)

| 설정 | 기본값 |
|---|---|
| 표본 수 / 리뷰 상한 / 사진 상한 | 100 / 100 / 10 |
| 투숙 조건 | 오늘 +28일 체크인, 1박, 성인 2명, 객실 1개 |
| 지연·배치·차단 쿨다운 | `ACTION_DELAY`, `HOTEL_DELAY`, `BATCH_*`, `BLOCK_*` |
| 분류 키워드 | `ROBOT_KEYWORDS`, `AUTOMATION_STRONG_KEYWORDS`, `AUTOMATION_WEAK_KEYWORDS` |
| 매칭 가중치·인접 도시·브랜드 계열 | `MATCH_WEIGHTS`, `NEARBY_CITIES`, `BRAND_PATTERNS` |
| UI 셀렉터 | `SORT_*`, `NEXT_PAGE_SELECTORS`, `LOAD_MORE_SELECTORS` |
