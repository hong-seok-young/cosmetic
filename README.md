# 화장품 성분표 자동 작성 도구

한글 성분명만 적으면 **영문명 · CAS.NO · 기능(기원·정의)** 을 자동으로 채워
엑셀 성분표를 만들어 주는 도구입니다. 성분 비율은 직접 입력하시면 됩니다.

데이터 출처: 공공데이터포털 **식품의약품안전처\_화장품 원료성분정보** OpenAPI
- End Point: `https://apis.data.go.kr/1471000/CsmtcsIngdCpntInfoService01`

그동안 KCIA 성분사전에서 한 건씩 검색해 옮겨 적던 작업을, 목록만 넣으면
한 번에 처리하도록 자동화한 것입니다.

---

## 1. 준비 (최초 1회)

```bash
# 파이썬 패키지 설치
pip install -r requirements.txt
```

인증키는 `config.ini` 에 이미 넣어 두었습니다(이 파일은 git 에 올라가지 않습니다).
다른 PC에서 새로 받았다면 `config.example.ini` 를 `config.ini` 로 복사한 뒤
공공데이터포털에서 발급한 **일반 인증키(Decoding 키)** 를 넣으세요.

> 인증키는 환경변수 `DATA_GO_KR_KEY` 나 `--service-key` 옵션으로도 줄 수 있습니다.

## 2. 전체 성분 DB 내려받기 (최초 1회 / 가끔 갱신)

```bash
python cosmetic_tool.py fetch
```

성분 전체를 받아 `cache/ingredients.json` 에 저장합니다. 이후 검색은 이 캐시를
사용하므로 빠르고, 인터넷 없이도 동작합니다.

## 3. 입력 엑셀 준비

`input.xlsx` 를 만들고 첫 줄을 머리글로, 아래에 성분명을 적습니다.
템플릿이 필요하면:

```bash
python cosmetic_tool.py template
```

| 한글성분명 | 비율 |
|------------|------|
| 정제수     | to 100 |
| 글리세린   | 5.0 |
| 부틸렌글라이콜 | 3.0 |

- **한글성분명** 열은 필수입니다(`성분명`, `한글` 등이 들어간 머리글도 인식).
- **비율** 열은 선택입니다. 비워 두고 나중에 채워도 됩니다.

## 4. 성분표 만들기

```bash
python cosmetic_tool.py lookup
```

`output.xlsx` 가 생성됩니다.

| 한글성분명 | 영문명 | 성분비율 | CAS.NO | 기능 | 매칭상태 |
|---|---|---|---|---|---|
| 글리세린 | Glycerin | 5.0 | 56-81-5 | 보습제… | 찾음 |

- **매칭상태**: `찾음` / `유사일치` / `다중일치(첫번째)` / `못찾음`
- 못 찾은 성분은 실행 후 화면에도 목록으로 표시됩니다(띄어쓰기·표기 차이 확인용).

### 한 번에 실행

```bash
python cosmetic_tool.py run     # 캐시 없으면 자동 fetch 후 lookup
```

---

## 명령어 요약

| 명령 | 설명 |
|------|------|
| `fetch` | 전체 성분 DB 를 받아 `cache/` 에 저장 |
| `lookup` | `input.xlsx` 의 성분명을 캐시와 대조해 `output.xlsx` 생성 |
| `run` | fetch(필요시) + lookup |
| `template` | 입력 엑셀 템플릿 생성 |
| `probe` | API 응답 원문·필드명 확인(점검용) |

### 자주 쓰는 옵션

```bash
# 입력/출력 파일 경로 지정
python cosmetic_tool.py lookup --input 내성분.xlsx --output 결과.xlsx

# 성분명/비율 열을 머리글 이름이나 0부터 시작하는 번호로 지정
python cosmetic_tool.py lookup --name-col 성분명 --ratio-col 1

# 오퍼레이션명을 직접 지정(자동 탐색 실패 시)
python cosmetic_tool.py fetch --operation getCsmtcsIngdCpntInfoService01 --debug
```

---

## 문제 해결

- **인증키 오류(code 30 등)**: 공공데이터포털에서 해당 API "활용신청"이 승인됐는지,
  Decoding 키를 넣었는지 확인하세요. 신청 직후엔 적용까지 시간이 걸릴 수 있습니다.
- **오퍼레이션을 못 찾음**: `python cosmetic_tool.py probe --service-key <키>` 로
  응답을 확인하고, 상세페이지의 오퍼레이션명을 `--operation` 으로 지정하세요.
- **필드가 이상하게 채워짐**: `probe` 출력의 "필드명"을 알려주시면 매핑을 맞춰
  드리겠습니다(영문명/CAS/기능 컬럼 자동 감지가 데이터에 따라 빗나갈 수 있음).
- **성분이 안 잡힘**: 표기 차이(띄어쓰기, 괄호 등)일 수 있습니다. 도구가 공백·기호를
  무시하고 비교하지만, 그래도 안 되면 KCIA 성분사전 기준 표준명으로 적어 보세요.
