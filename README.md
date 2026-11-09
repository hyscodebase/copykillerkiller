# copykillerkiller

카피킬러 AI 생성검사 상세 결과를 Word 원고와 대조하고, 표시된 문단을 자료와 작성자의 판단에 근거해 검토하는 Codex 스킬입니다. 보고서 판본 확인, 문단별 수정 이유, 원본을 보존하는 DOCX 검증 절차를 담았습니다.

## 설치

Git이 설치된 터미널에서 아래 명령을 실행하세요. GitHub 로그인 없이 내려받을 수 있습니다. Codex 스킬 폴더에 `copykillerkiller`가 이미 있다면 내용을 확인한 뒤 갱신하세요.

```sh
skill_dir="${CODEX_HOME:-$HOME/.codex}/skills/copykillerkiller"
if [ ! -e "$skill_dir" ] && [ ! -L "$skill_dir" ]; then
  git clone https://github.com/hyscodebase/copykillerkiller.git &&
    mkdir -p "$(dirname "$skill_dir")" &&
    cp -R copykillerkiller/copykillerkiller "$skill_dir"
fi
```

## 사용 예시

원고와 상세 결과 PDF를 첨부하고 요청하세요.

> `$copykillerkiller` 최신 상세 보고서와 본문 DOCX를 대조해 표시 구간별 검토 목록을 만들어줘. 파일은 그대로 둬.

> `$copykillerkiller` 과제 지침에 맞춰 표시 구간에서 근거가 약한 문단을 검토용 사본으로 다듬어줘. 표와 인용, AI 활용 고지는 보존해줘.

스킬은 검사 결과를 읽는 데 활용합니다. 새 검사 점수는 해당 판본의 상세 결과에서 확인할 수 있습니다.
