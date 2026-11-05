# Excel Formula Remover

Excel 통합 문서의 수식을 **마지막으로 저장된 계산값**으로 바꾸는 도구입니다.
Windows용 GUI와 자동화용 CLI를 함께 제공합니다.

> 이 도구는 수식을 계산하지 않습니다. Excel에서 통합 문서를 다시 계산하고
> 저장한 뒤 사용하세요. 특히 수동 계산 모드인 파일은 저장된 값이 오래되었을 수
> 있습니다.

## 안전 보장

- 원본 `.xlsx`, `.xlsm`, `.xlsb` 파일은 수정하지 않습니다.
- 파일별로 임시 복사본을 변환하고 다시 열어 확인한 뒤, 같은 출력 폴더에 한 번에
  게시합니다.
- 대상 수식에 저장된 계산값이 없거나 안전하게 나눌 수 없는 수식 범위가 있으면
  해당 파일 전체를 실패 처리하며 부분 출력은 남기지 않습니다. 배치의 다른 파일은
  계속 처리합니다.
- 기존 파일과 앞서 만든 출력은 덮어쓰지 않습니다. `Book.xlsx`의 첫 출력은
  `Book_수식제거.xlsx`, 다음 충돌은 `Book_수식제거 (2).xlsx`가 됩니다.
- 취소 전에 완료된 출력은 유지하고, 처리 중이던 임시 파일은 정리합니다.
- `.xlsm`과 `.xlsb`의 VBA 프로젝트를 보존합니다.

지원 형식은 `.xlsx`, `.xlsm`, `.xlsb`(Excel 바이너리 통합 문서)입니다. 출력은
입력과 같은 형식으로 저장합니다. 이전 형식인 `.xls`, CSV 등은 지원하지 않습니다.

## 변환 범위

- 시작 행의 기본값은 `1`이며, 지정한 행부터 마지막 행까지 수식을 값으로
  바꿉니다. 시작 행 위의 수식은 유지합니다.
- 제외 패턴은 시트 이름에 대한 대소문자 구분 없는 부분 문자열입니다. 일치하는
  시트는 전체를 그대로 유지합니다.
- 시작 행이 여러 셀에 걸친 수식 정의를 가르면 공유 수식은 안전한 경우 위쪽
  범위로 줄입니다. 배열 수식, 데이터 테이블 수식 또는 Excel 표 수식 범위를
  안전하게 나눌 수 없으면 해당 파일을 실패 처리합니다. 이 경우 시작 행을 `1`로
  바꾸거나 해당 수식 범위 뒤로 옮기세요.
- 범위 안의 Excel 표는 계산된 열 수식과 사용자 지정 요약 행 수식을 제거합니다.
  합계·평균 같은 요약 행 함수 설정은 유지하며, 요약 셀에는 저장된 값만
  남습니다. Excel은 요약 셀에 값이 있는데 함수 설정이 없는 표를 열지 못합니다.
- 수식을 제거한 통합 문서는 계산 체인을 삭제하고 다음에 열 때 전체 다시 계산을
  요청합니다. `.xlsb`는 변경된 시트의 바이너리 인덱스(빠른 로드용 보조 데이터)도
  삭제하며, Excel이 필요할 때 다시 만듭니다.

## Windows GUI

릴리스의 `excel-formula-remover-gui-<version>-win-x64.zip`을 풀고
`ExcelFormulaRemover.exe`를 실행합니다. GUI 릴리스에는 .NET 런타임이 포함되어
있습니다.

1. **파일 추가**를 누르거나 `.xlsx`/`.xlsm`/`.xlsb` 파일을 목록에 끌어 놓습니다.
2. 출력 폴더를 선택합니다.
3. 필요하면 시작 행과 제외할 시트 이름 조각을 입력합니다. GUI에서는 여러 제외
   패턴을 쉼표(`,`)나 세미콜론(`;`)으로 구분합니다.
4. **값으로 변환**을 누릅니다.

처리가 끝나면 성공, 실패, 경고와 실제 저장 경로를 한 번에 보여 줍니다. 변환 중
취소하거나 창을 닫으면 현재 파일을 안전하게 정리한 뒤 종료합니다.

## CLI

포터블 릴리스의
`excel-formula-remover-cli-<version>-portable.zip`을 풀고 .NET 10 런타임이
설치된 환경에서 실행합니다.

```powershell
dotnet .\excel-formula-remover.dll --output .\converted .\Book.xlsx
```

여러 파일, 시작 행, 제외 패턴을 함께 지정할 수 있습니다.

```powershell
dotnet .\excel-formula-remover.dll `
  --output .\converted `
  --start-row 2 `
  --exclude TF `
  --exclude "변동성" `
  .\Book.xlsx .\MacroBook.xlsm
```

### 옵션

| 옵션 | 설명 |
| --- | --- |
| `-o`, `--output <directory>` | 필수 출력 폴더. 한 번만 지정합니다. |
| `-r`, `--start-row <number>` | 변환을 시작할 행. 기본값 `1`, 범위 `1`~`1048576`. |
| `-x`, `--exclude <pattern>` | 수식을 유지할 시트 이름 조각. 여러 번 지정할 수 있습니다. |
| `--json` | 기본 TOON 대신 상세 NDJSON을 출력합니다. |
| `-h`, `--help` | 도움말을 출력합니다. |
| `--` | 이후 인수를 옵션이 아닌 입력 경로로 처리합니다. |

긴 옵션은 `--output=converted`처럼 `=` 형식으로도 지정할 수 있습니다.

### 출력 형식

기본 표준 출력은 결정적인 순서의 간결한
[TOON 4.1](https://toonformat.dev/reference/spec) 문서입니다.

- `files`: 요청 순서의 입력, 성공 여부, 실제 출력 경로, 실패 종류
- `messages`: 파일별 오류와 경고
- `summary`: 전체, 성공, 실패, 취소 여부

`--json`은 파일마다 `result` 객체 한 줄을 쓰고 마지막에 `summary` 객체를 쓰는
NDJSON입니다. 각 결과에는 수식 및 표 수식 정의 제거 수, 처리/제외 시트 수와
경고가 포함됩니다. 두 형식 모두 시간이나 진행 로그를 섞지 않으므로 자동화에서
안정적으로 읽을 수 있습니다.

### 종료 코드

| 코드 | 의미 |
| ---: | --- |
| `0` | 모든 파일 성공 또는 도움말 출력 |
| `1` | 하나 이상의 파일 변환 실패 또는 실행 중 오류 |
| `2` | 잘못된 CLI 인수 |
| `130` | 사용자 취소 |

취소 시에도 이미 완료된 파일 결과와 `canceled: true`가 출력됩니다.

소스에서 CLI를 실행하려면 다음 명령을 사용합니다.

```powershell
dotnet run `
  --project .\src\ExcelFormulaRemover.Cli\ExcelFormulaRemover.Cli.csproj `
  -- `
  --output .\converted .\Book.xlsx
```

## 개발과 검증

필수 도구는 `global.json`에 지정된 .NET SDK `10.0.110`, PowerShell 7, 그리고
PATH에서 실행 가능한 `rumdl`입니다. 새 클론에서 전체 검증은 다음 한 명령으로
실행합니다.

```powershell
pwsh -NoProfile -File .\scripts\check.ps1
```

이 스크립트는 잠금 파일 기반 복원, 포맷 확인, Release 빌드, 전체 테스트,
직접/전이 패키지 취약점 검사와 Markdown 검사를 비대화식으로 수행합니다. 테스트
결과는 `artifacts/test-results/`에 기록됩니다.

코드와 Markdown을 자동 포맷하려면 다음을 실행합니다.

```powershell
pwsh -NoProfile -File .\scripts\format.ps1
```

검증을 통과한 뒤 새 SemVer 버전의 릴리스 산출물을 만들려면 다음을 실행합니다.

```powershell
pwsh -NoProfile -File .\scripts\publish.ps1 1.0.0
```

## 릴리스 산출물 계약

`publish.ps1`은 기존 버전 디렉터리를 변경하지 않으며 다음 파일만
`artifacts/publish/<version>/`에 게시합니다.

- `excel-formula-remover-gui-<version>-win-x64.zip`: Windows x64용 자체 포함,
  단일 파일 GUI
- `excel-formula-remover-cli-<version>-portable.zip`: .NET 10 런타임이 필요한
  프레임워크 종속 CLI
- `SHA256SUMS.txt`: 두 ZIP의 SHA-256 체크섬

현재 릴리스 바이너리는 코드 서명되지 않습니다. Windows에서 알 수 없는 게시자나
SmartScreen 경고가 표시될 수 있으므로 신뢰하는 배포 경로에서 받은 파일인지와
`SHA256SUMS.txt`를 확인하고, 조직의 보안 정책을 우회하지 마세요.

## 프로젝트 구조

```text
src/
  ExcelFormulaRemover.Core/       OpenXML·XLSB 변환 엔진과 결과 모델
  ExcelFormulaRemover.Cli/        비대화식 CLI 어댑터
  ExcelFormulaRemover.App/        Windows WinForms GUI
tests/
  ExcelFormulaRemover.Core.Tests/ 변환 안전성·충실도 테스트
  ExcelFormulaRemover.Cli.Tests/  파싱·출력·종료 코드 테스트
  ExcelFormulaRemover.App.Tests/  GUI 지원 로직·수명 주기 테스트
scripts/
  check.ps1                       전체 검증 게이트
  format.ps1                      코드와 Markdown 포맷
  publish.ps1                     버전 고정 릴리스 생성
PLAN.md                           현재 프로젝트 상태와 제약
plans/rewrite.md                  2026 재작성 설계 기록
```

`Core`와 `Cli`는 `net10.0`이며 .NET 10이 지원되는 운영체제에서 실행할 수
있습니다. `App`만 `net10.0-windows`와 WinForms에 의존하므로 GUI 실행은
Windows 전용입니다.
