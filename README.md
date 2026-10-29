# Word 텍스트 인용을 Zotero 필드로 자동 연결하기 (Python)

Oct 29, 2026 · @테르비나

## 소개

텍스트로만 적힌 저자-연도 인용 (Yuk et al., 2022)을, 이미 내 Zotero 폴더에 모아 둔 논문과 자동으로 연결해 Zotero 활성 필드로 바꿔 주는 Python 스크립트입니다. docx 파일을 아이콘 위에 끌어다 놓기만 하면 됩니다.

- 입력: 본문에 (저자 et al., 연도) 형식 인용과 끝에 References 목록이 있는 Word 원고(.docx)
- 출력: `원고이름_zotero.docx` — 인용이 Zotero 필드로 바뀌고, 참고문헌 목록은 Zotero Bibliography로 바뀐 파일 (원본은 그대로)
- 스타일: Human Reproduction으로 미리 설정 (Word에서 Refresh 한 번이면 적용)
- 시험 결과: 참고문헌 31편 원고에서 인용 묶음 29개가 모두 연결됨. 31편 중 30편은 자동으로, 다른 폴더에 있던 1편은 확인 질문 한 번으로 연결

## 준비물

| 항목 | 내용 |
| --- | --- |
| Zotero 7 + Word 플러그인 | 인용할 논문이 한 폴더(콜렉션)에 모여 있어야 함 |
| Zotero 데이터 폴더 | `zotero.sqlite`가 있는 폴더. 내 경우 `D:\OneDrive\zotero` (스크립트 기본값) |
| Anaconda (또는 Miniconda) | Windows용 |
| `zotero_link_citations.py` | 본 스크립트 (본 글에 첨부) |
| `zotero_link.bat` | 끌어다 놓기용 실행 파일 (본 글에 첨부) |

Zotero 데이터 폴더 위치는 Zotero → 편집 → 설정 → 고급 → 파일 및 폴더에서 확인합니다. 다른 위치라면 py 파일 윗부분의 `DEFAULT_ZOTERO_DIR` 값을 바꾸거나 실행할 때 `--zotero-dir` 옵션을 씁니다.

## 설치 (처음 한 번만)

1. **전용 conda 환경 만들기.** Anaconda Prompt에서 아래 한 줄을 실행합니다. 세 패키지 모두 conda-forge에 있습니다.
2. **파일 두기.** `zotero_link_citations.py`와 `zotero_link.bat`을 같은 폴더에 둡니다 (예: `D:\tools`). bat은 자기와 같은 폴더에 있는 py 파일을 찾습니다.
3. **바로가기 만들기.** `zotero_link.bat` 오른쪽 클릭 → 추가 옵션 표시 → 바로 가기 만들기. 만든 바로가기를 다운로드 폴더나 바탕화면에 둡니다.

```
conda create -n zotlink -c conda-forge python=3.11 python-docx lxml rapidfuzz
```

bat은 Anaconda를 `%USERPROFILE%\anaconda3`, `C:\ProgramData\anaconda3` 등 흔한 위치에서 자동으로 찾습니다. "Anaconda folder not found"가 나오면 Anaconda Prompt에서 `where conda`로 위치를 확인해 bat 안의 후보 목록에 추가합니다.

conda 환경은 필수는 아닙니다 (base에 `pip install python-docx`만 해도 작동). 다만 따로 두면 base 환경의 패키지가 꼬이지 않고, `conda remove -n zotlink --all` 한 줄로 지울 수 있습니다.

## 사용법

1. **Zotero를 켜고 원고용 폴더를 선택해 둡니다.** 스크립트가 지금 선택된 폴더를 자동으로 알아냅니다.
2. **docx를 바로가기 위에 끌어다 놓습니다.** 여러 개를 한꺼번에 놓아도 되고, 더블클릭하면 파일 선택 창이 뜹니다. 이미 변환된 `_zotero.docx`는 건너뜁니다.
3. **폴더를 확인합니다.** `이 폴더에서만 찾을까요? [Y/n]` → Enter 또는 y. 하위 폴더가 있으면 포함 여부도 묻습니다.
4. **확인 질문에 답합니다.** 확실한 것은 자동으로 연결되고, 애매하거나 폴더에 없는 것만 물습니다 (아래 표).
5. **Word에서 `원고이름_zotero.docx`를 열고 Zotero 탭 → Refresh.** 원본 파일에서 Refresh를 누르면 "You must insert a citation…" 오류가 나니, 파일 이름 끝의 `_zotero`를 꼭 확인합니다.

확인 질문에서 쓸 수 있는 입력:

| 입력 | 동작 |
| --- | --- |
| `1`, `2` … | 후보 목록에서 그 번호의 항목으로 연결 |
| `a` | 전체 라이브러리에서 이 참고문헌과 비슷한 항목 자동 검색 |
| `a 검색어` | 전체 라이브러리에서 검색 (예: `a Lethaby`) |
| `s` / `s 검색어` | 선택한 폴더 안에서만 검색 |
| Enter | 건너뛰기 (그 인용은 텍스트로 남음) |

후보가 제목만 비슷한 다른 논문일 수 있어 저자와 연도를 꼭 확인합니다. 예를 들어 Lethaby 2015 Cochrane 리뷰 자리에 개정판 Bofill Rodriguez 2020이 후보로 뜨기도 합니다.

Anaconda Prompt에서 직접 실행할 수도 있습니다.

```
conda activate zotlink
python D:\tools\zotero_link_citations.py "D:\원고\manuscript.docx"
```

## 동작 원리

핵심은 문서 끝의 References 목록을 다리로 쓰는 것입니다. 본문의 (Yuk et al., 2022)만으로는 같은 저자·같은 연도 논문을 구분할 수 없지만, 목록에는 제목과 DOI가 있기 때문입니다.

1. **DB 읽기:** `zotero.sqlite`를 임시 폴더로 복사해 읽습니다. 원본 DB는 건드리지 않고, Zotero가 켜져 있어도 됩니다.
2. **폴더 한정:** Zotero의 로컬 Connector API(`127.0.0.1:23119`)로 지금 선택된 폴더를 알아내고, 그 폴더(선택 시 하위 폴더 포함) 안에서만 찾습니다. 다른 프로젝트의 동명 논문이 걸리지 않습니다.
3. **참고문헌 ↔ Zotero 항목:** 목록의 항목마다 DOI 일치 → 제1저자+연도 → 제목 유사도 순서로 대조합니다. 확실할 때만 자동 연결하고 나머지는 사용자에게 물습니다.
4. **본문 인용 ↔ 참고문헌:** (A et al., 2006; B and C, 2010) 같은 괄호 인용과 A et al. (2006) 같은 서술형 인용을 찾아, 저자 수(et al. / and / 단독)와 2022a·b 접미사까지 맞춰 목록 항목과 잇습니다.
5. **필드 삽입:** 해당 텍스트를 Zotero가 쓰는 것과 같은 `ADDIN ZOTERO_ITEM CSL_CITATION` 필드로 감쌉니다. 글꼴과 기울임 et al.은 그대로 두고, 참고문헌 목록은 `ZOTERO_BIBL` 필드로 바꿉니다. 표에만 인용된 문헌도 목록에 남도록 넣어 둡니다.
6. **스타일 설정:** 문서 속성 `ZOTERO_PREF`에 Human Reproduction(`http://www.zotero.org/styles/human-reproduction`)을 저장합니다. 참고문헌 서식은 첫 Refresh 때 Zotero가 적용합니다.
7. **선택 기록:** 사용자가 고른 내용을 `원고이름_zotero_map.json`에 저장해, 같은 원고를 다시 돌릴 때 같은 질문을 반복하지 않습니다. 지워도 결과 파일에는 영향이 없습니다.

## 옵션과 문제 해결

| 옵션 | 용도 |
| --- | --- |
| `--collection "폴더명"` | 폴더를 이름이나 경로로 직접 지정 (예: `"Projects/LNG-IUS"`) |
| `--zotero-dir 경로` | `zotero.sqlite`가 다른 곳에 있을 때 |
| `-o 출력.docx` | 결과 파일 이름 지정 |
| `--keep-references` | 텍스트 참고문헌 목록을 Zotero 필드로 바꾸지 않고 그대로 둠 |
| `--auto` | 묻지 않고 확실한 것만 연결 |
| `--yes` | 감지된 폴더와 하위 폴더 포함을 묻지 않고 사용 |
| `--fresh` | 이전 선택 기록(json)을 무시하고 처음부터 |

**Refresh를 눌렀더니 "You must insert a citation before performing this operation".** 원본 파일을 열었을 때 나는 오류입니다. `_zotero.docx`를 열어야 합니다. 인용을 클릭했을 때 회색 음영이 보이면 필드가 들어간 것입니다 (Alt+F9로 필드 코드 확인 가능).

**참고문헌 목록이 줄간격 없이 뭉쳐 보인다.** 초기 버전의 버그로, 현재 버전에서 고쳐졌습니다. 이미 만든 파일은 Zotero 탭 → Document Preferences에서 다른 스타일로 바꿨다가 Human Reproduction으로 돌아오면 정상이 됩니다.

**목록에 "Hum Reprod Oxf Engl", "England." 같은 글자가 붙는다.** PubMed에서 가져온 Zotero 항목의 Journal Abbr와 Place 값입니다. Zotero에서 항목을 고치면 사라집니다.

**괄호 인용 하나가 텍스트로 남았다.** (A, 2006; B, 2010) 중 하나라도 연결이 안 되면 그 괄호 전체를 텍스트로 둡니다. 실행 끝에 목록이 나오니, Zotero에 논문을 추가하고 다시 실행하면 그 것만 다시 묻습니다.

**지원 형식.** 저자-연도 인용(Human Reproduction, APA 등)만 대상입니다. \[1\] 같은 번호 인용과 각주(footnote) 안의 인용은 처리하지 않습니다.
