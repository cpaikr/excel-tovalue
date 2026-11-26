# xlclean

Excel 통합 문서를 원본은 그대로 두고 정리된 새 파일로 만드는 도구입니다.
Windows용 GUI와 자동화용 CLI를 함께 제공합니다.

- **수식을 값으로**(`xlclean values`): 수식을 Excel이 **마지막으로 저장한
  계산값**으로 바꿉니다.
- **스타일 정리**(`xlclean styles`): XLStylesTool처럼 쓰이지 않는 사용자 지정 셀
  스타일과 셀 서식을 지우고, 똑같아 보이는 셀 서식을 합칩니다. "서로 다른 셀
  서식이 너무 많습니다" 오류나 다른 통합 문서에서 복사되어 쌓인 수천 개의
  스타일을 정리할 때 씁니다.

> `values`는 수식을 계산하지 않습니다. Excel에서 통합 문서를 다시 계산하고
> 저장한 뒤 사용하세요. 특히 수동 계산 모드인 파일은 저장된 값이 오래되었을 수
> 있습니다.

## 안전 보장

- 원본 `.xlsx`, `.xlsm`, `.xlsb` 파일은 수정하지 않습니다.
- 파일별로 임시 복사본을 처리하고 다시 열어 확인한 뒤, 같은 출력 폴더에 한 번에
  게시합니다.
- 결과를 증명할 수 없으면 해당 파일 전체를 실패 처리하며 부분 출력은 남기지
  않습니다. 배치의 다른 파일은 계속 처리합니다.
- 기존 파일과 앞서 만든 출력은 덮어쓰지 않습니다. `Book.xlsx`의 첫 출력은
  `Book_수식제거.xlsx`(`values`) 또는 `Book_스타일정리.xlsx`(`styles`), 다음
  충돌은 `Book_수식제거 (2).xlsx`처럼 번호가 붙습니다.
- 취소 전에 완료된 출력은 유지하고, 처리 중이던 임시 파일은 정리합니다.
- `.xlsm`과 `.xlsb`의 VBA 프로젝트를 보존합니다.

지원 형식은 `.xlsx`, `.xlsm`, `.xlsb`(Excel 바이너리 통합 문서)입니다. 출력은
입력과 같은 형식으로 저장합니다. 이전 형식인 `.xls`, CSV 등은 지원하지 않습니다.

## 수식을 값으로 변환하는 범위

- 시작 행의 기본값은 `1`이며, 지정한 행부터 마지막 행까지 수식을 값으로
  바꿉니다. 시작 행 위의 수식은 유지합니다.
- 제외 패턴은 시트 이름에 대한 대소문자 구분 없는 부분 문자열입니다. 일치하는
  시트는 전체를 그대로 유지합니다.
- 대상 수식에 저장된 계산값이 없으면 해당 파일을 실패 처리합니다.
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

## 스타일 정리 범위

- **지우는 것**: 어떤 셀·행·열도 쓰지 않는 셀 서식, 그리고 쓰는 셀 서식이 없는
  사용자 지정 셀 스타일.
- **합치는 것**: 표시 형식·글꼴·채우기·테두리·맞춤·보호와 소속 셀 스타일까지
  모두 같은 셀 서식. 셀이 가리키는 번호만 바뀌고 모양은 바뀌지 않습니다.
- **유지하는 것**: 기본 제공 스타일(표준, 좋음, 제목 1 등), 표가 이름으로 쓰는
  스타일, 셀이 쓰는 스타일(서로 똑같아 보여도 이름이 다르면 그대로 둠), 그리고
  표시 형식·글꼴·채우기·테두리·조건부 서식 목록.
- 게시 전에 원본과 결과를 따로 다시 읽어 모든 셀·행·열의 서식이 똑같은지, 스타일
  정보 외의 내용이 원본과 같은지 확인합니다. 다르면 출력하지 않습니다.
- Excel 4.0 매크로 시트, 공유 통합 문서 변경 기록, Strict Open XML 형식, 이
  도구가 모르는 `.xlsb` 셀 레코드가 있으면 `unsupported-workbook-content`로
  실패합니다.
- 경고: 매크로가 있는 통합 문서에서 사용자 지정 스타일을 지웠을 때(매크로가
  스타일 이름을 쓸 수 있음), 정리 후에도 Excel의 서식 개수 한도(65,430)를 넘을
  때.
- 정리할 것이 없으면 원본과 같은 사본을 게시합니다.

## Windows GUI

릴리스의 `xlclean-gui-<version>-win-x64.zip`을 풀고 `XlClean.exe`를 실행합니다.
GUI 릴리스에는 .NET 런타임이 포함되어 있습니다. GUI는 CLI와 같은 엔진으로 두
작업을 제공합니다.

1. **작업**에서 **수식을 값으로 바꾸기**(`xlclean values`) 또는 **스타일
   정리**(`xlclean styles`)를 고릅니다.
2. **파일 추가**를 누르거나 `.xlsx`/`.xlsm`/`.xlsb` 파일을 목록에 끌어 놓습니다.
3. 출력 폴더를 선택합니다.
4. 수식 변환이면 필요에 따라 시작 행과 제외할 시트 이름 조각을 입력합니다.
   GUI에서는 여러 제외 패턴을 쉼표(`,`)나 세미콜론(`;`)으로 구분합니다. 스타일
   정리에는 범위 설정이 없어 이 항목이 숨겨집니다.
5. **값으로 변환** 또는 **스타일 정리**를 누릅니다.

처리가 끝나면 성공, 실패, 경고와 실제 저장 경로를 한 번에 보여 줍니다. 스타일
정리는 파일마다 셀 스타일과 셀 서식의 정리 전후 개수, 합친 서식 수도 보여
줍니다. 처리 중 취소하거나 창을 닫으면 현재 파일을 안전하게 정리한 뒤
종료합니다.

## CLI

`xlclean`은 런타임 설치가 필요 없는 단일 실행 파일입니다. 사람과 AI 에이전트가
같은 명령과 같은 출력 계약을 사용합니다.

> 4.0.0에서 도구 이름이 `excel-formula-remover`에서 `xlclean`으로 바뀌었고,
> 수식 변환은 `xlclean values` 하위 명령이 되었습니다. 3.x를 쓰던 경우 아래
> 절차로 새로 설치하세요. 새 설치는 이전 설치를 지우지 않습니다. 필요 없으면
> Windows에서는 `%LOCALAPPDATA%\excel-formula-remover` 폴더와 사용자 `PATH`의 그
> 항목을, Linux에서는 `~/.local/bin/excel-formula-remover`를 지우세요.

### CLI 설치

**지원 환경.** Windows x64, Linux x64(glibc). 다른 플랫폼용 실행 파일은 게시하지
않습니다. .NET 런타임이나 Excel은 필요하지 않습니다. 릴리스는
[공개 GitHub 릴리스](https://github.com/cpaikr/xlclean/releases)에서 로그인
없이 받습니다. Linux에서는 `curl`과 `tar`가 필요합니다.

**Windows.** 빈 작업 폴더의 PowerShell에서 실행합니다.

```powershell
$repo = "cpaikr/xlclean"
$api = "https://api.github.com/repos/$repo/releases/latest"
$tag = (Invoke-RestMethod $api -UseBasicParsing).tag_name
$version = $tag.TrimStart("v")
$base = "https://github.com/$repo/releases/download/$tag"
$archive = "xlclean-cli-$version-win-x64.zip"
$files = $archive, "SHA256SUMS.txt", "install.ps1", "verify-windows-install.ps1"
foreach ($name in $files) {
  Invoke-WebRequest "$base/$name" -OutFile $name -UseBasicParsing
}
.\install.ps1 -Archive ".\$archive" -Checksums ".\SHA256SUMS.txt"
```

**Linux.**

```sh
repo=cpaikr/xlclean
tag=$(curl -fsSL "https://api.github.com/repos/$repo/releases/latest" |
  sed -n 's/^ *"tag_name": "\(.*\)",$/\1/p')
base="https://github.com/$repo/releases/download/$tag"
archive="xlclean-cli-${tag#v}-linux-x64.tar.gz"
for name in "$archive" SHA256SUMS.txt install.sh; do
  curl -fsSLO "$base/$name"
done
sh install.sh "$archive" SHA256SUMS.txt
```

설치 스크립트는 SHA-256 체크섬과 압축 파일 내용을 확인하고, 새 실행 파일이
릴리스 버전을 보고하는지 실행해 본 뒤에만 기존 설치를 교체합니다. 확인에
실패하면 기존 설치를 그대로 둡니다. 압축 파일, `SHA256SUMS.txt`, 설치
스크립트는 반드시 같은 릴리스의 것을 사용하세요. 조직의 스크립트 실행 정책을
따르세요.

**PATH.** 기본 설치 위치는 Windows `%LOCALAPPDATA%\xlclean\bin`, Linux
`~/.local/bin`입니다. 다른 위치는 `-BinDirectory <폴더>` 또는 `install.sh`의 세
번째 인수로 지정합니다. 설치 스크립트는 `PATH`를 바꾸지 않습니다. Windows에서
사용자 `PATH`에 영구 등록하려면 다음을 실행하고 새 터미널을 여세요.

```powershell
$bin = "$env:LOCALAPPDATA\xlclean\bin"
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
$newPath = $userPath.TrimEnd(";") + ";$bin"
[Environment]::SetEnvironmentVariable("Path", $newPath, "User")
```

**설치 확인.** 먼저 전체 경로로, 다음에 명령 이름으로 확인합니다. 출력된 버전이
받은 릴리스와 같아야 합니다.

```powershell
$bin = "$env:LOCALAPPDATA\xlclean\bin"
& "$bin\xlclean.exe" --version
xlclean --help
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
xlclean values --output .\converted .\Book.xlsx
xlclean styles --output .\cleaned .\Bloated.xlsx
```

`values`는 여러 파일, 시작 행, 제외 패턴을 함께 지정할 수 있습니다.

```powershell
xlclean values `
  --output .\converted `
  --start-row 2 `
  --exclude TF `
  --exclude "변동성" `
  .\Book.xlsx .\MacroBook.xlsm
```

두 작업을 모두 하려면 `values`의 출력에 `styles`를 실행합니다.

### 옵션

| 옵션 | 명령 | 설명 |
| --- | --- | --- |
| `-o`, `--output <directory>` | 모두 | 필수 출력 폴더. 한 번만 지정합니다. |
| `-r`, `--start-row <number>` | `values` | 변환을 시작할 행. 기본값 `1`, 범위 `1`~`1048576`. |
| `-x`, `--exclude <pattern>` | `values` | 수식을 유지할 시트 이름 조각. 여러 번 지정할 수 있습니다. |
| `-h`, `--help` | 모두 | 도움말. 명령 이름 뒤에 쓰면 그 명령의 도움말입니다. |
| `--version` | 모두 | `xlclean <version>`을 출력합니다. |
| `--` | 모두 | 이후 인수를 옵션이 아닌 입력 경로로 처리합니다. |

긴 옵션은 `--output=converted`처럼 `=` 형식으로도 지정할 수 있습니다. 인수 없이
실행하면 `--help`와 바이트 단위로 같은 도움말을 출력하고 `0`으로 종료합니다.
명령 없이 옵션을 쓰거나(3.x 형식) 알 수 없는 옵션을 쓰면 사용 오류입니다.

### 출력 형식

명령을 실행한 모든 경우 표준 출력에 JSON 문서 **하나**와 줄바꿈 하나를 쓰고,
표준 오류에는 아무것도 쓰지 않습니다. 사용 오류와 취소도 같은 형식입니다.
도움말과 버전만 일반 텍스트입니다.

```json
{
  "ok": false,
  "version": "4.0.0",
  "command": "styles",
  "files": [
    {
      "input": "Bloated.xlsx",
      "status": "converted",
      "output": "C:\\work\\cleaned\\Bloated_스타일정리.xlsx",
      "failure": null,
      "changes": {
        "cellStyles": { "before": 5003, "after": 4 },
        "cellFormats": { "before": 812, "after": 37 },
        "mergedCellFormats": 120
      },
      "warnings": []
    },
    {
      "input": "Macro.xlsm",
      "status": "failed",
      "output": null,
      "failure": { "kind": "unsupported-workbook-content", "message": "..." },
      "changes": null,
      "warnings": []
    }
  ],
  "summary": { "total": 2, "succeeded": 1, "failed": 1, "canceled": false },
  "error": null
}
```

실제 출력은 한 줄입니다. 위 예시는 읽기 쉽게 줄을 나눈 것입니다.

- `ok`: 모든 파일이 성공해 종료 코드가 `0`일 때만 `true`입니다.
- `command`: 실행한 명령(`values` 또는 `styles`). 올바른 명령이 없으면 `null`.
- `files`: 요청 순서의 파일별 결과. `status`는 `converted` 또는 `failed`입니다.
  출력 파일 이름은 충돌을 피해 정해지므로 항상 `output`에서 읽으세요.
- `changes`: 명령별 개수. 실패한 파일은 `null`입니다.
  - `values`: `removedFormulas`, `removedTableFormulaDefinitions`,
    `processedSheets`, `excludedSheets`.
  - `styles`: `cellStyles`와 `cellFormats`의 전후 개수, `mergedCellFormats`.
- `failure.kind`: `invalid-input`, `missing-cached-value`,
  `formula-range-crosses-boundary`, `invalid-workbook`, `file-access`,
  `unsupported-workbook-content`, `unexpected` 중 하나입니다. `message`는
  진단용 문장이며 식별자가 아닙니다.
- `warnings`: 수동 계산 모드, 매크로가 있는 파일의 스타일 제거, 서식 한도 초과
  같은 경고 문장입니다.
- `summary`: 전체, 성공, 실패 수와 취소 여부입니다.
- `error`: 명령 자체가 실패했을 때만 값이 있습니다. `kind`는 `usage` 또는
  `unexpected`이며, 이때 `files`는 비어 있습니다.

### 종료 코드

| 코드 | 의미 |
| ---: | --- |
| `0` | 모든 파일 성공, 또는 도움말·버전 출력 |
| `1` | 하나 이상의 파일 실패 또는 실행 중 오류(`error`) |
| `2` | 잘못된 CLI 인수 |
| `130` | 사용자 취소 |

취소 시에도 이미 완료된 파일 결과와 `summary.canceled: true`가 출력됩니다.
이때 `files`에는 취소 전에 끝난 파일만 들어 있습니다. Ctrl+C와 SIGTERM을 같은
취소로 처리합니다.

소스에서 CLI를 실행하려면 다음 명령을 사용합니다.

```powershell
dotnet run `
  --project .\src\XlClean.Cli\XlClean.Cli.csproj `
  -- `
  values --output .\converted .\Book.xlsx
```

## 에이전트 스킬

AI 에이전트용 스킬 원본은 [`skills/xlclean`](skills/xlclean/SKILL.md)에
있습니다. 명령 도움말이 알려 주지 못하는 판단을 담습니다. 어떤 명령을 쓸지,
저장값이 없는 파일(예: openpyxl로 만든 파일)의 처리, 오래된 저장값 경고, 스타일
정리의 경고, 실패 종류별 복구, 결과 보고 범위입니다. 설치 안내는 실행 파일이
없을 때만 읽도록 분리되어 있습니다.

[`skills` CLI](https://github.com/vercel-labs/skills)로 설치하거나, 저장소
클론에서 `skills/xlclean` 폴더 전체(`references/`, `agents/` 포함)를 사용하는
에이전트의 스킬 위치에 복사합니다.

```powershell
npx skills add cpaikr/xlclean --skill xlclean
```

```powershell
New-Item -ItemType Directory -Force "$HOME\.agents\skills" | Out-Null
Copy-Item -Recurse .\skills\xlclean "$HOME\.agents\skills\"
```

스킬은 명시적으로 호출할 때만 동작하도록 설정되어 있습니다. 실행 파일은 위
절차로 따로 설치합니다. 실행 파일이 없으면 스킬이 에이전트에게 같은 설치
절차를 따르게 합니다. 3.x의 `excel-formula-remover` 스킬을 설치했다면 지우고
`xlclean`을 설치하세요.

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
`artifacts/publish/<version>/`에 게시합니다. 3.x 릴리스의 자산 이름은
`xlclean` 대신 `excel-formula-remover`, GUI 실행 파일은
`ExcelFormulaRemover.exe`입니다.

- `xlclean-gui-<version>-win-x64.zip`: Windows x64용 자체 포함, 단일 파일
  GUI(`XlClean.exe`)
- `xlclean-cli-<version>-win-x64.zip`: Windows x64용 자체 포함, 단일 파일
  CLI(`xlclean.exe`, `LICENSE.md`, `THIRD-PARTY-NOTICES.txt`)
- `xlclean-cli-<version>-linux-x64.tar.gz`: Linux x64(glibc)용 자체 포함, 단일
  파일 CLI(`xlclean`, `LICENSE.md`, `THIRD-PARTY-NOTICES.txt`)
- `SHA256SUMS.txt`: 세 압축 파일의 SHA-256 체크섬
- `install.ps1`, `install.sh`, `verify-windows-install.ps1`: CLI 설치와 설치
  확인 스크립트

세 압축 파일의 `THIRD-PARTY-NOTICES.txt`는 저장소의
[THIRD-PARTY-NOTICES.txt](THIRD-PARTY-NOTICES.txt)에 그 빌드가 포함한 .NET
런타임 팩의 라이선스와 고지를 이어 붙여 만듭니다. 설치 스크립트는 실행 파일만
설치합니다.

CLI는 트리밍하지 않은 자체 포함 빌드이며 압축 파일 하나가 약 34MB입니다.

현재 릴리스 바이너리는 코드 서명되지 않습니다. Windows에서 알 수 없는 게시자나
SmartScreen 경고가 표시될 수 있으므로 신뢰하는 배포 경로에서 받은 파일인지와
`SHA256SUMS.txt`를 확인하고, 조직의 보안 정책을 우회하지 마세요.

## 프로젝트 구조

```text
src/
  XlClean.Core/                   OpenXML·XLSB 변환 엔진과 결과 모델
  XlClean.Cli/                    비대화식 CLI 어댑터
  XlClean.App/                    Windows WinForms GUI
tests/
  XlClean.Core.Tests/             변환 안전성·충실도 테스트
  XlClean.Cli.Tests/              파싱·출력·종료 코드와 프로세스 계약 테스트
  XlClean.App.Tests/              GUI 지원 로직·수명 주기 테스트
skills/
  xlclean/                        AI 에이전트용 CLI 소비자 스킬
scripts/
  check.ps1                       전체 검증 게이트
  format.ps1                      코드와 Markdown 포맷
  publish.ps1                     버전 고정 릴리스 생성
  test-artifact.ps1               게시된 CLI 압축 파일 설치·계약 검증
  install/                        릴리스에 포함되는 CLI 설치 스크립트
PLAN.md                           현재 프로젝트 상태와 제약
plans/rewrite.md                  2026 재작성 설계 기록
plans/agent-cli.md                에이전트용 CLI와 스킬 설계 기록
plans/style-cleanup.md            xlclean 이름 변경과 스타일 정리 설계
```

`Core`와 `Cli`는 `net10.0`이며 Windows API에 의존하지 않습니다. 게시하는 CLI
실행 파일은 Windows x64와 Linux x64용입니다.
`App`만 `net10.0-windows`와 WinForms에 의존하므로 GUI 실행은
Windows 전용입니다.

## 라이선스

Elastic License 2.0. [LICENSE.md](LICENSE.md)를 참고하세요. 릴리스 실행
파일에 포함된 Open XML SDK와 .NET 런타임의 라이선스 고지는
[THIRD-PARTY-NOTICES.txt](THIRD-PARTY-NOTICES.txt)와 각 릴리스 압축 파일에
있습니다.
