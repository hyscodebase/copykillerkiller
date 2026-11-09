# copykillerkiller

**과제 제출을 앞두고 카피킬러 AI 생성률이 높게 나왔다면, 어느 문단부터 다시 봐야 할까요?**

`copykillerkiller`는 AI 생성률을 낮추기 위해 과제 원고를 다시 살펴보는 대학생을 위한 Codex 스킬입니다. 카피킬러 상세 보고서의 표시를 Word 원고와 연결하고, 해당 문단에 출처나 설명이 빠진 곳을 짚어줍니다. 원고와 보고서의 판본도 대조해서 예전 초고의 검사 결과를 보며 현재 원고를 검토할 때 생기는 혼동을 줄입니다.

## 과제에서 이렇게 써보세요

Codex 대화에서 `$copykillerkiller`를 입력하고 Word 원고와 카피킬러 상세 보고서를 첨부하세요. 수정본이 여럿이면 지금 쓰는 원고를 지정하고, 분량이나 AI 활용에 관한 수업 지침이 있다면 함께 알려주세요.

> `$copykillerkiller` 상세 보고서와 지금 과제 원고를 대조해줘. 표시된 문단마다 내가 확인할 출처나 보완할 근거를 짚고, 원본을 보존한 Word 주석 사본으로 만들어줘.

가상의 과제에서 초안에 다음과 같이 적었다고 해보죠.

> 복습 기록을 쓰면 학생들의 성적이 오른다.

갖고 있는 자료가 수강생 12명 중 9명이 복습 기록을 제출했다는 기록이라면, 그 자료로 설명할 수 있는 것은 제출 결과입니다. 스킬은 해당 문단에 성적 변화까지 말하려면 어떤 자료가 더 필요한지 주석으로 남깁니다.

과제 지침이 AI 문장 수정을 제한하면 직접 고칠 수 있도록 확인 질문과 검토 의견을 제공합니다. 점수 변화를 보려면 수정한 판본을 같은 설정으로 검사하고 새 상세 보고서에서 확인하세요.

보고서 없이도 쓸 수 있습니다. 예를 들어 `$copykillerkiller 이 수업 메모로 보고서 개요를 잡아줘. 내가 확인할 출처와 수치는 따로 표시해줘.`처럼 요청하세요. 업무 보고, 제안서, 이메일 등 다른 문서 유형도 [작성 가이드](copykillerkiller/references/writing-playbook.md)에서 다룹니다.

## 설치

Git이 설치된 터미널에서 내려받으세요. GitHub 로그인 없이 설치할 수 있습니다.

```sh
git clone https://github.com/hyscodebase/copykillerkiller.git && cd copykillerkiller
```

내려받은 저장소에서 스킬 폴더를 복사합니다. 이미 설치되어 있으면 경로를 알려주고 종료합니다.

```sh
skill_dir="${CODEX_HOME:-$HOME/.codex}/skills/copykillerkiller"
if [ -e "$skill_dir" ] || [ -L "$skill_dir" ]; then
  printf '기존 설치를 확인하세요. %s\n' "$skill_dir"
else
  mkdir -p "$(dirname "$skill_dir")" &&
    mkdir "$skill_dir" &&
    cp -R copykillerkiller/. "$skill_dir/" &&
    printf '설치 완료. %s\n' "$skill_dir"
fi
```

자동 전처리에는 Python 3.9 이상, PDF에서 텍스트를 추출하려면 Poppler의 `pdftotext`가 필요합니다. `python3 --version`, `pdftotext -v`로 확인할 수 있습니다. 자동 전처리를 실행할 수 없는 환경에서는 PDF 화면과 원고를 직접 비교하는 절차를 사용합니다.

이미 설치했다면 내려받은 저장소에서 `git pull --ff-only`를 실행하고, 설치 폴더의 개인 수정이 있는지 확인한 뒤 새 `copykillerkiller` 폴더로 교체하세요. Codex에서 스킬을 찾지 못하면 새 대화를 열거나 설치한 `SKILL.md`를 지정해 주세요.

## 자동화 범위

| 단계 | 스킬이 하는 일 |
|---|---|
| 초안과 편집 | 학술 글, 업무·기술 보고, 제안서, 공지·이메일, 에세이의 독자와 목적에 맞춰 작성합니다. 제공한 자료로 뒷받침할 수 있는 범위를 지킵니다. |
| 검사본 대조 | 보고서의 표시를 원고 후보와 대조합니다. 검사 당시 원고와 현재 작업 대상은 따로 표시합니다. |
| 근거 검토 | 표시 구간의 경계를 PDF에서 확인하고, 위치·문제·확인할 자료를 정리합니다. |
| 수정 사본 검증 | 수정 요청이 있고 과제 규정이 허용되면 별도 사본을 만듭니다. 본문 텍스트 변화와 서식 차이를 비교해 주석 사본의 용도를 표시합니다. |

표시 범위의 정확한 끝, 페이지 사이 연결, 여러 판본의 공통 본문은 추가 확인이 필요합니다. 도구는 로컬에서 동작하며, 주석 사본을 만들 때는 지정된 작업 대상의 해시와 문단을 확인합니다.

저장소를 내려받은 폴더에서 직접 실행할 수도 있습니다.

```sh
python3 copykillerkiller/scripts/prepare_review.py --target-index 2 "상세 결과.pdf" "검사 당시 원고.docx" "현재 원고.docx"
python3 copykillerkiller/scripts/create_review_copy.py "현재 원고.docx" --expected-sha256 "<작업 대상 SHA-256>" --notes "주석.json" --out "검토용 사본.docx"
python3 copykillerkiller/scripts/verify_docx.py "원본.docx" "검토용 사본.docx"
```

자세한 작성 흐름과 결과 해석은 [문서 유형별 글쓰기](copykillerkiller/references/writing-playbook.md), [전처리 안내](copykillerkiller/references/preflight.md), [DOCX 검증 안내](copykillerkiller/references/docx-verification.md)를 참고하세요.

## 통합 테스트

`python3 -B -m unittest discover -s tests`는 임시 폴더에 가상 보고서와 이름이 같은 과거·현재 DOCX를 만듭니다. 과거 보고서가 있더라도 현재 원고로 주석 사본이 생성되는지, 원본과 문서 구조가 보존되는지 검증합니다.

`soffice`, `pdftotext`, `pdftoppm`을 사용할 수 있는 환경에서는 다음 한 줄로 본문 페이지의 픽셀 비교까지 실행할 수 있습니다.

```sh
COPYKILLER_RENDER_TEST=1 python3 -B -m unittest discover -s tests
```
