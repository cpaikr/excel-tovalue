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

`excel-formula-remover`는 런타임 설치가 필요 없는 단일 실행 파일입니다. 사람과
AI 에이전트가 같은 명령과 같은 출력 계약을 사용합니다.

> 아래 설치 방법과 JSON 출력은 3.0.0부터 적용됩니다. 2.x 릴리스는 .NET 10
> 런타임이 필요한 `portable` ZIP과 TOON/NDJSON 출력을 사용했습니다.

### CLI 설치

**지원 환경.** Windows x64, Linux x64(glibc). 다른 플랫폼용 실행 파일은 게시하지
않습니다. .NET 런타임이나 Excel은 필요하지 않습니다. 저장소가 비공개이므로
릴리스를 받으려면 저장소 접근 권한으로 로그인된
[GitHub CLI](https://cli.github.com/)(`gh auth status`)가 필요합니다.

**Windows.** 빈 작업 폴더의 PowerShell에서 실행합니다.

```powershell
$repo = "sjunepark/excel-tovalue"
$tag = gh release view --repo $repo --json tagName --jq .tagName
$version = $tag.TrimStart("v")
$archive = "excel-formula-remover-cli-$version-win-x64.zip"
gh release download $tag --repo $repo `
  --pattern $archive --pattern SHA256SUMS.txt `
  --pattern install.ps1 --pattern verify-windows-install.ps1
.\install.ps1 -Archive ".\$archive" -Checksums ".\SHA256SUMS.txt"
```

**Linux.**

```sh
repo=sjunepark/excel-tovalue
tag=$(gh release view --repo "$repo" --json tagName --jq .tagName)
archive="excel-formula-remover-cli-${tag#v}-linux-x64.tar.gz"
gh release download "$tag" --repo "$repo" \
  --pattern "$archive" --pattern SHA256SUMS.txt --pattern install.sh
sh install.sh "$archive" SHA256SUMS.txt
```

설치 스크립트는 SHA-256 체크섬과 압축 파일 내용을 확인하고, 새 실행 파일이
릴리스 버전을 보고하는지 실행해 본 뒤에만 기존 설치를 교체합니다. 확인에
실패하면 기존 설치를 그대로 둡니다. 압축 파일, `SHA256SUMS.txt`, 설치
스크립트는 반드시 같은 릴리스의 것을 사용하세요. 조직의 스크립트 실행 정책을
따르세요.

**PATH.** 기본 설치 위치는 Windows `%LOCALAPPDATA%\excel-formula-remover\bin`,
Linux `~/.local/bin`입니다. 다른 위치는 `-BinDirectory <폴더>` 또는
`install.sh`의 세 번째 인수로 지정합니다. 설치 스크립트는 `PATH`를 바꾸지
않습니다. Windows에서 사용자 `PATH`에 영구 등록하려면 다음을 실행하고 새
터미널을 여세요.

```powershell
$bin = "$env:LOCALAPPDATA\excel-formula-remover\bin"
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
$newPath = $userPath.TrimEnd(";") + ";$bin"
[Environment]::SetEnvironmentVariable("Path", $newPath, "User")
```

**설치 확인.** 먼저 전체 경로로, 다음에 명령 이름으로 확인합니다. 출력된 버전이
받은 릴리스와 같아야 합니다.

```powershell
$bin = "$env:LOCALAPPDATA\excel-formula-remover\bin"
& "$bin\excel-formula-remover.exe" --version
excel-formula-remover --help
```

패키지형 데스크톱 앱(에이전트 앱 포함)이 띄운 Windows 셸은 `%LOCALAPPDATA%`
쓰기가 앱 전용 저장소로 리디렉션될 수 있습니다. 설치 스크립트는 이런 위치를
거부하며, 이 경우 독립적으로 연 PowerShell에서 다시 실행하거나
`-BinDirectory "$env:USERPROFILE\.local\bin"`을 지정하세요. 호출한 앱 밖의
프로세스에서도 실행 파일과 `PATH` 등록이 보이는지는 다음으로 확인합니다.

```powershell
.\verify-windows-install.ps1 -ExpectedVersion $version
```

**업그레이드.** 자동 업데이트나 업데이트 확인 기능은 없습니다. 새 릴리스의
파일로 위 설치 절차를 다시 실행하면 기존 실행 파일을 교체합니다.

### 사용

```powershell
excel-formula-remover --output .\converted .\Book.xlsx
```

여러 파일, 시작 행, 제외 패턴을 함께 지정할 수 있습니다.

```powershell
excel-formula-remover `
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
| `-h`, `--help` | 도움말을 출력합니다. |
| `--version` | `excel-formula-remover <version>`을 출력합니다. |
| `--` | 이후 인수를 옵션이 아닌 입력 경로로 처리합니다. |

긴 옵션은 `--output=converted`처럼 `=` 형식으로도 지정할 수 있습니다. 인수 없이
실행하면 `--help`와 바이트 단위로 같은 도움말을 출력하고 `0`으로 종료합니다.
알 수 없는 옵션은 사용 오류입니다.

### 출력 형식

변환을 시도한 모든 실행은 표준 출력에 JSON 문서 **하나**와 줄바꿈 하나를 쓰고,
표준 오류에는 아무것도 쓰지 않습니다. 사용 오류와 취소도 같은 형식입니다.
도움말과 버전만 일반 텍스트입니다.

```json
{
  "ok": false,
  "version": "3.0.0",
  "files": [
    {
      "input": "Book.xlsx",
      "status": "converted",
      "output": "C:\\work\\converted\\Book_수식제거.xlsx",
      "failure": null,
      "removedFormulas": 128,
      "removedTableFormulaDefinitions": 2,
      "processedSheets": 3,
      "excludedSheets": 1,
      "warnings": []
    },
    {
      "input": "Draft.xlsx",
      "status": "failed",
      "output": null,
      "failure": { "kind": "missing-cached-value", "message": "..." },
      "removedFormulas": 0,
      "removedTableFormulaDefinitions": 0,
      "processedSheets": 0,
      "excludedSheets": 0,
      "warnings": []
    }
  ],
  "summary": { "total": 2, "succeeded": 1, "failed": 1, "canceled": false },
  "error": null
}
```

실제 출력은 한 줄입니다. 위 예시는 읽기 쉽게 줄을 나눈 것입니다.

- `ok`: 모든 파일이 변환되어 종료 코드가 `0`일 때만 `true`입니다.
- `files`: 요청 순서의 파일별 결과. `status`는 `converted` 또는 `failed`입니다.
  출력 파일 이름은 충돌을 피해 정해지므로 항상 `output`에서 읽으세요.
- `failure.kind`: `invalid-input`, `missing-cached-value`,
  `formula-range-crosses-boundary`, `invalid-workbook`, `file-access`,
  `unexpected` 중 하나입니다. `message`는 진단용 문장이며 식별자가 아닙니다.
- `warnings`: 수동 계산 모드처럼 저장값이 오래되었을 수 있다는 경고 문장입니다.
- `summary`: 전체, 성공, 실패 수와 취소 여부입니다.
- `error`: 명령 자체가 실패했을 때만 값이 있습니다. `kind`는 `usage` 또는
  `unexpected`이며, 이때 `files`는 비어 있습니다.

### 종료 코드

| 코드 | 의미 |
| ---: | --- |
| `0` | 모든 파일 성공, 또는 도움말·버전 출력 |
| `1` | 하나 이상의 파일 변환 실패 또는 실행 중 오류(`error`) |
| `2` | 잘못된 CLI 인수 |
| `130` | 사용자 취소 |

취소 시에도 이미 완료된 파일 결과와 `summary.canceled: true`가 출력됩니다.
이때 `files`에는 취소 전에 끝난 파일만 들어 있습니다. Ctrl+C와 SIGTERM을 같은
취소로 처리합니다.

소스에서 CLI를 실행하려면 다음 명령을 사용합니다.

```powershell
dotnet run `
  --project .\src\ExcelFormulaRemover.Cli\ExcelFormulaRemover.Cli.csproj `
  -- `
  --output .\converted .\Book.xlsx
```

## 에이전트 스킬

AI 에이전트용 스킬 원본은
[`skill/excel-formula-remover`](skill/excel-formula-remover/SKILL.md)에 있습니다.
명령 도움말이 알려 주지 못하는 판단을 담습니다. 저장값이 없는 파일(예:
openpyxl로 만든 파일)의 처리, 오래된 저장값 경고, 실패 종류별 복구, 결과 보고
범위입니다. 설치 안내는 실행 파일이 없을 때만 읽도록 분리되어 있습니다.

`skill/excel-formula-remover` 폴더 전체(`references/`, `agents/` 포함)를
사용하는 에이전트의 스킬 위치에 복사해 설치합니다. 스킬은 명시적으로 호출할 때만
동작하도록 설정되어 있습니다. 실행 파일은 위 절차로 따로 설치합니다.

## 개발과 검증

필수 도구는 `global.json`에 지정된 .NET SDK `10.0.110`, PowerShell 7, 그리고
PATH에서 실행 가능한 `rumdl`입니다. 새 클론에서 전체 검증은 다음 한 명령으로
실행합니다.

```powershell
pwsh -NoProfile -File .\scripts\check.ps1
```

이 스크립트는 잠금 파일 기반 복원, 포맷 확인, Release 빌드, 전체 테스트,
직접/전이 패키지 취약점 검사, Markdown 검사와 스킬 패키지 검사를 비대화식으로
수행합니다. 테스트에는 CLI를 외부 프로세스로 실행하는 계약 테스트가 포함됩니다.
테스트 결과는 `artifacts/test-results/`에 기록됩니다.

코드와 Markdown을 자동 포맷하려면 다음을 실행합니다.

```powershell
pwsh -NoProfile -File .\scripts\format.ps1
```

검증을 통과한 뒤 새 SemVer 버전의 릴리스 산출물을 로컬에서 만들려면 다음을
실행합니다. Windows에서 실행하며, Windows CLI 압축 파일을 설치 스크립트로 설치한
뒤 설치된 실행 파일에 프로세스 계약 테스트를 실행합니다.

```powershell
pwsh -NoProfile -File .\scripts\publish.ps1 1.0.0
```

Linux 압축 파일은 Linux에서 같은 검증을 실행합니다.

```sh
pwsh -NoProfile -File ./scripts/test-artifact.ps1 artifacts/publish/1.0.0 1.0.0
```

정식 릴리스는 `v<version>` 태그를 푸시하면 `.github/workflows/release.yml`이
전체 검증, 게시, Windows·Linux 산출물 검증을 거쳐 GitHub 릴리스를 만듭니다.
게시된 태그와 자산은 바꾸지 않으며, 수정은 새 버전으로 합니다.

## 릴리스 산출물 계약

`publish.ps1`은 기존 버전 디렉터리를 변경하지 않으며 다음 파일만
`artifacts/publish/<version>/`에 게시합니다.

- `excel-formula-remover-gui-<version>-win-x64.zip`: Windows x64용 자체 포함,
  단일 파일 GUI
- `excel-formula-remover-cli-<version>-win-x64.zip`: Windows x64용 자체 포함,
  단일 파일 CLI(`excel-formula-remover.exe` 하나)
- `excel-formula-remover-cli-<version>-linux-x64.tar.gz`: Linux x64(glibc)용
  자체 포함, 단일 파일 CLI(`excel-formula-remover` 하나)
- `SHA256SUMS.txt`: 세 압축 파일의 SHA-256 체크섬
- `install.ps1`, `install.sh`, `verify-windows-install.ps1`: CLI 설치와 설치
  확인 스크립트

CLI는 트리밍하지 않은 자체 포함 빌드이며 압축 파일 하나가 약 34MB입니다.

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
  ExcelFormulaRemover.Cli.Tests/  파싱·출력·종료 코드와 프로세스 계약 테스트
  ExcelFormulaRemover.App.Tests/  GUI 지원 로직·수명 주기 테스트
skill/
  excel-formula-remover/          AI 에이전트용 CLI 소비자 스킬
scripts/
  check.ps1                       전체 검증 게이트
  format.ps1                      코드와 Markdown 포맷
  publish.ps1                     버전 고정 릴리스 생성
  test-artifact.ps1               게시된 CLI 압축 파일 설치·계약 검증
  install/                        릴리스에 포함되는 CLI 설치 스크립트
PLAN.md                           현재 프로젝트 상태와 제약
plans/rewrite.md                  2026 재작성 설계 기록
plans/agent-cli.md                에이전트용 CLI와 스킬 설계 기록
```

`Core`와 `Cli`는 `net10.0`이며 Windows API에 의존하지 않습니다. 게시하는 CLI
실행 파일은 Windows x64와 Linux x64용입니다.
`App`만 `net10.0-windows`와 WinForms에 의존하므로 GUI 실행은
Windows 전용입니다.
