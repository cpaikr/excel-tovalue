# Claude Code guide

- 제품 동작과 사용법은 [README.md](README.md), 현재 상태와 제약은
  [PLAN.md](PLAN.md)를 먼저 확인한다.
- 저장소의 기준 검증 명령은
  `pwsh -NoProfile -File .\scripts\check.ps1`이다. 코드나 Markdown을
  수정한 뒤에는 필요하면 `scripts/format.ps1`로 포맷하고 기준 검증을 실행한다.
- `ExcelFormulaRemover.Core`와 `ExcelFormulaRemover.Cli`는 Windows API에
  의존하지 않게 유지한다. WinForms와 Windows 전용 동작은
  `ExcelFormulaRemover.App`에만 둔다.
- 수식을 계산하지 않는다. Excel 파일에 마지막으로 저장된 계산값만 사용한다.
- 원본이나 기존 출력을 덮어쓰지 않는다. 파일별 임시 복사본을 검증한 뒤
  원자적으로 게시하며, 저장값 누락이나 안전하지 않은 범위 경계는 부분 출력 없는
  파일 실패로 처리한다.
- 동작이나 자동화가 바뀌면 README와 PLAN 계열 문서를 함께 최신 상태로 유지한다.
