# outlook-mcp

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](./LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18-339933?logo=node.js&logoColor=white)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](#요구사항)
[![MCP](https://img.shields.io/badge/MCP-Model%20Context%20Protocol-8A2BE2)](https://modelcontextprotocol.io)
[![GitHub Repo stars](https://img.shields.io/github/stars/hayein-bit/outlook-mcp?style=social)](https://github.com/hayein-bit/outlook-mcp)
[![한국어](https://img.shields.io/badge/lang-한국어-lightgrey?style=social&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiMwMDAwMDAiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cGF0aCBzdHJva2U9Im5vbmUiIGQ9Ik0wIDBoMjR2MjRIMHoiIGZpbGw9Im5vbmUiIC8+PHBhdGggZD0iTTMgMTJhOSA5IDAgMSAwIDE4IDBhOSA5IDAgMCAwIC0xOCAwIiAvPjxwYXRoIGQ9Ik0zLjYgOWgxNi44IiAvPjxwYXRoIGQ9Ik0zLjYgMTVoMTYuOCIgLz48cGF0aCBkPSJNMTEuNSAzYTE3IDE3IDAgMCAwIDAgMTgiIC8+PHBhdGggZD0iTTEyLjUgM2ExNyAxNyAwIDAgMSAwIDE4IiAvPjwvc3ZnPg==)](#outlook-mcp)
[![English](https://img.shields.io/badge/lang-English-lightgrey?style=social&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiMwMDAwMDAiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cGF0aCBzdHJva2U9Im5vbmUiIGQ9Ik0wIDBoMjR2MjRIMHoiIGZpbGw9Im5vbmUiIC8+PHBhdGggZD0iTTMgMTJhOSA5IDAgMSAwIDE4IDBhOSA5IDAgMCAwIC0xOCAwIiAvPjxwYXRoIGQ9Ik0zLjYgOWgxNi44IiAvPjxwYXRoIGQ9Ik0zLjYgMTVoMTYuOCIgLz48cGF0aCBkPSJNMTEuNSAzYTE3IDE3IDAgMCAwIDAgMTgiIC8+PHBhdGggZD0iTTEyLjUgM2ExNyAxNyAwIDAgMSAwIDE4IiAvPjwvc3ZnPg==)](#english)

Windows **Classic Outlook**(데스크톱 앱)과 연동되는 [MCP](https://modelcontextprotocol.io)(Model Context Protocol) 서버입니다.
Microsoft Graph API 대신 **Outlook COM**(Object Model)을 직접 사용하므로, 별도의 Azure AD 앱 등록이나
클라우드 인증 없이 로컬에 설치된 Outlook을 그대로 자동화합니다.

받은 편지함을 회사/카테고리별로 자동 정리하고, 메일을 검색·조회하며, 사용자의 문체로 답장 초안을
Outlook 임시보관함(Drafts)에 준비해 두는 것을 목표로 합니다. **메일은 항상 사용자가 직접 검토한 후
발송**하며, 이 서버는 메일을 임의로 발송하는 기능을 제공하지 않습니다.

> 특정 회사에 종속되지 않도록, 고객사·카테고리·답장 스타일은 모두 `config/` 안의 파일만 수정하면
> 코드 변경 없이 다른 조직에서도 그대로 사용할 수 있도록 설계했습니다.

## 동작 원리

Node.js는 Windows COM 객체를 직접 다룰 수 없으므로, 이 프로젝트는 Outlook과의 모든 상호작용을
**PowerShell 자식 프로세스**에 위임합니다.

```mermaid
flowchart TB
    accTitle: Outlook MCP 요청 처리 상세 흐름
    accDescr: MCP 클라이언트 요청이 임시 JSON 파일을 통해 outlook-mcp 서버에서 PowerShell 자식 프로세스로 전달되고, Outlook 실행 여부에 따라 기존 인스턴스에 접속하거나 새로 실행한 뒤 COM 호출 결과가 같은 경로로 되돌아오는 과정을 보여준다.

    subgraph client_layer ["MCP 클라이언트"]
        client["MCP Client<br/>(Claude 등)"]:::input_style
    end

    subgraph server_layer ["outlook-mcp 서버 (Node.js/TypeScript)"]
        write_request[("임시 요청 JSON<br/>호출별 임시 폴더 생성")]:::process_style
        spawn_ps["PowerShell 자식 프로세스 실행"]:::process_style
        read_response[("임시 응답 JSON")]:::process_style
        cleanup["임시 폴더 정리"]:::aux_style
    end

    subgraph ps_layer ["PowerShell → Outlook COM"]
        check_running{"Outlook<br/>실행 중?"}
        attach_instance["기존 인스턴스에 접속<br/>(새 창·프로필 프롬프트 없음)"]:::process_style
        launch_instance["새 인스턴스 실행"]:::process_style
        outlook["Classic Outlook<br/>실행 중인 인스턴스"]:::output_style
        com_call["Outlook.Application<br/>COM 호출"]:::process_style
    end

    client -->|"stdio 요청(JSON-RPC)"| write_request
    write_request --> spawn_ps
    spawn_ps --> check_running
    check_running -->|예| attach_instance
    check_running -->|아니오| launch_instance
    attach_instance --> outlook
    launch_instance --> outlook
    outlook --> com_call
    com_call --> read_response
    read_response -->|"stdio 응답(JSON-RPC)"| client
    read_response -.->|"호출 종료 시"| cleanup

    classDef input_style fill:#e9eeec,stroke:#4c5b60,stroke-width:2px,color:#2c3a3d
    classDef process_style fill:#232522,stroke:#111111,stroke-width:2px,color:#f5f3ec
    classDef output_style fill:#f6efde,stroke:#8a6f45,stroke-width:2px,color:#4a3b22
    classDef aux_style fill:none,stroke:#a85e1a,stroke-width:2px,stroke-dasharray:4 3,color:#7a4712
```

요청 파라미터와 결과를 stdout이 아니라 임시 JSON 파일로 주고받는 이유는 로케일·인코딩 문제를 피하기 위해서입니다.

## 요구사항

- Windows 10/11
- Classic Outlook(데스크톱 앱, Microsoft 365 / Exchange / POP/IMAP 계정 무관) 설치 및 프로필 설정
- Node.js 18 이상
- Windows PowerShell 5.1 (Windows 기본 내장, 별도 설치 불필요)

> New Outlook(웹 기반 신규 Outlook)은 COM Object Model을 지원하지 않으므로 이 프로젝트로 제어할 수
> 없습니다. Classic Outlook으로 전환되어 있어야 합니다.

## 설치

### 1. 사전 준비

- Windows 10/11
- Classic Outlook 데스크톱 앱이 설치되어 있고, 한 번 이상 실행해서 계정 설정을 마친 상태
- [Node.js](https://nodejs.org) 18 이상 (PowerShell에서 `node --version`으로 확인)
- Windows PowerShell 5.1 (Windows에 기본 내장되어 있어 별도 설치가 필요 없습니다)

### 2. 저장소 내려받기

```bash
git clone https://github.com/hayein-bit/outlook-mcp.git
cd outlook-mcp
```

git이 없다면 저장소 페이지의 **Code → Download ZIP**으로 내려받아 압축을 풀어도 됩니다.

### 3. 설정 파일 준비

`config` 폴더의 `.example` 파일을 복사해 `.example`을 제거한 파일을 생성합니다.

PowerShell에서는 다음과 같이 복사합니다:

```powershell
Copy-Item config\customers.example.yml config\customers.yml
Copy-Item config\categories.example.yml config\categories.yml
Copy-Item config\reply_style.example.md config\reply_style.md
```

(macOS/Linux나 bash 환경이라면 `cp config/customers.example.yml config/customers.yml` 형태로
바꿔 사용합니다.) 탐색기에서 파일을 복사·붙여넣기한 뒤 이름에서 `.example`만 지워도 동일하게 동작합니다.

방금 생성한 세 파일을 열어 자신의 조직(고객사 목록, 카테고리, 답장 문체)에 맞게 채워 넣습니다.
당장 채우지 않고 예시 값 그대로 빌드해 동작만 먼저 확인해도 됩니다.

### 4. 의존성 설치 및 빌드

```bash
npm install
npm run build
```

`npm run build`는 TypeScript를 `dist/`로 컴파일하고, `src/outlook/scripts/*.ps1`을
`dist/outlook/scripts/`로 복사합니다. 다음과 비슷한 줄이 마지막에 출력되면 성공입니다:

```
PowerShell 스크립트 복사 완료: ...\src\outlook\scripts -> ...\dist\outlook\scripts (9개 파일)
```

### 5. 빌드 확인 (선택)

Outlook을 실행한 상태에서 서버가 정상적으로 기동되는지만 빠르게 확인하려면 다음을 실행합니다:

```bash
npm start
```

`outlook-mcp 서버가 시작되었습니다 (stdio).` 로그가 출력되면 정상입니다(`Ctrl+C`로 종료). 이 명령만으로는
어떤 도구도 호출되지 않으며, 실제로 사용하려면 아래와 같이 MCP 클라이언트에 연결해야 합니다.

## 설정 (config/)

각 설정 파일은 `*.example.*` 템플릿(공개, git 추적)과 실제 사용 파일(비공개, git 무시)로 나뉩니다.
실제 고객사명·사내 도메인 등이 담기는 쪽은 항상 `.example`이 없는 파일입니다.

| 템플릿 (git 추적) | 실사용 파일 (git 무시) | 역할 |
|---|---|---|
| `config/customers.example.yml` | `config/customers.yml` | 고객사 목록과 별칭. 메일 제목/본문/참조에서 별칭을 검색해 고객사를 판별합니다. |
| `config/categories.example.yml` | `config/categories.yml` | 고객사가 아닌 메일을 분류할 카테고리 규칙(키워드/발신자 도메인)과 이동할 폴더 경로. |
| `config/reply_style.example.md` | `config/reply_style.md` | 답장 작성 시 항상 참고하는 문체/톤 가이드. 자유 형식의 Markdown. |

세 파일 모두 **코드를 건드리지 않고** 자유롭게 수정·추가·삭제할 수 있습니다.

### customers.yml 예시

```yaml
customers:
  - name: 고객사A       # 이동할 폴더 이름 (받은 편지함 > 2. 고객사 > 고객사A)
    aliases:
      - 고객사A
      - A사

  - name: 고객사B
    aliases:
      - 고객사B
      - B사
```

> 고객사 판별은 **발신자 이메일 도메인이 아니라** 제목·본문·참조(CC)에 별칭이 등장하는지로
> 판단합니다. 내부 직원이 대신 보낸 메일이라도 본문에 "고객사B"가 있으면 해당 폴더로 이동합니다.

### categories.yml 예시

```yaml
categories:
  - name: 사내공지
    folder: "3. 사내공지"
    keywords: ["[공지]", "[안내]", 전사공지, 인사발령]
    sender_domains: []

default_folder: "7. 기타"
customer_root_folder: "2. 고객사"
calendar_folder: "6. 일정알림"
```

규칙은 위에서부터 순서대로 검사하며, 고객사 판별 → 캘린더성 항목(`calendar_folder`) → 카테고리
판별 순으로 적용합니다.

### 기본 폴더 구조

```
받은 편지함
├── 1. 참고자료
├── 2. 고객사
│   ├── 고객사A
│   ├── 고객사B
│   └── 고객사C
├── 3. 사내공지
├── 4. 뉴스레터
├── 5. 업무논의
├── 6. 일정알림
└── 7. 기타
```

존재하지 않는 폴더는 `organize_mail` 실행 시 자동으로 생성됩니다. 폴더 이름 앞의 번호는 예시일
뿐이므로, 제거하고 싶다면 `folder`/`customer_root_folder`/`calendar_folder`/`default_folder`
값에서 번호만 빼면 됩니다.

### reply_style.md

정중함의 정도, 인사말·맺음말 형식, 서명 등 답장 문체를 자유롭게 정의합니다. `create_reply` 도구는
이 파일 전체를 항상 함께 읽어, AI가 답장을 작성할 때 참고하도록 전달합니다.

### config 위치 변경

기본적으로 서버를 실행한 디렉터리(`cwd`) 기준 `./config`를 사용합니다. 다른 위치를 사용하려면
`OUTLOOK_MCP_CONFIG_DIR` 환경변수를 지정합니다.

## MCP 클라이언트에 연결하기

이 서버는 로컬 stdio 프로세스로 동작합니다. **Outlook이 설치된 PC에서 실행되는 MCP 클라이언트**에만
연결할 수 있으며, 원격·클라우드 세션에서는 Outlook에 접근할 수 없어 동작하지 않습니다.

아래 예시의 경로(`C:/outlook-mcp`)는 실제로 저장소를 내려받은 위치로 바꿔서 사용합니다. Windows
경로의 백슬래시(`\`)는 JSON·커맨드라인 안에서 슬래시(`/`)로 쓰거나 이스케이프(`\\`)해야 합니다.

### Claude Desktop에 연결하기

1. 설정 파일을 엽니다: `%APPDATA%\Claude\claude_desktop_config.json`
   (탐색기 주소창에 `%APPDATA%\Claude`를 붙여넣으면 해당 폴더로 바로 이동합니다. 파일이 없으면 새로 만듭니다.)
2. `mcpServers` 항목에 다음 내용을 추가합니다(파일이 비어 있다면 전체를 그대로 붙여넣습니다):

   ```json
   {
     "mcpServers": {
       "outlook-mcp": {
         "command": "node",
         "args": ["C:/outlook-mcp/dist/server/index.js"],
         "env": {
           "OUTLOOK_MCP_CONFIG_DIR": "C:/outlook-mcp/config"
         }
       }
     }
   }
   ```

3. 파일을 저장한 뒤 **Claude Desktop을 완전히 종료했다가 다시 실행합니다**(트레이 아이콘까지 종료해야
   반영됩니다).
4. 새 대화를 열고 도구(🔨) 아이콘을 눌렀을 때 `outlook-mcp`가 목록에 보이면 연결이 완료된 것입니다.

### Claude Code CLI에 연결하기

`claude mcp add`는 Claude Code 대화 안이 아니라 일반 터미널(PowerShell/cmd)에서 실행하는
명령입니다. 아래처럼 한 번만 등록해 두면 됩니다.

```bash
claude mcp add outlook-mcp --scope local -e OUTLOOK_MCP_CONFIG_DIR="C:/outlook-mcp/config" -- node "C:/outlook-mcp/dist/server/index.js"
```

scope 옵션:

- `--scope local` (기본값, 권장): 명령을 실행한 디렉터리에서만 도구가 보입니다. `outlook-mcp`
  폴더 안에서 실행하면, 이후 그 폴더에서 여는 세션에만 적용됩니다.
- `--scope user`: 이 PC의 동일 계정이면 어느 디렉터리에서 세션을 열어도 도구가 보입니다.
- `--scope project`: 등록 정보가 프로젝트 폴더 안에 저장되어, 같은 프로젝트를 쓰는 다른
  사람에게도 함께 적용됩니다. 팀에서 설정을 공유할 때만 사용을 권장합니다.

기타 옵션:

- `-e KEY=VALUE`: 서버 프로세스에 전달할 환경변수입니다. `OUTLOOK_MCP_CONFIG_DIR`는 필수는
  아니지만, Claude Code 실행 디렉터리가 매번 달라질 수 있으므로 명시해 두기를 권장합니다.
- `--` 뒤: 실제로 실행할 명령과 인자입니다.

등록은 한 번만 하면 됩니다. 이후에는 등록해 둔 범위 안에서 평소처럼 `claude`로 새 세션을
열기만 하면 도구가 자동으로 보입니다.

등록 후 확인:

```bash
claude mcp list              # outlook-mcp 가 목록에 보이는지
claude mcp get outlook-mcp   # 등록된 command/env 상세 확인
```

정확한 플래그·명령은 Claude Code 버전에 따라 달라질 수 있으므로 `claude mcp add --help`로 다시
확인합니다. 등록 이후에는 아무 디렉터리에서나 새 세션을 열고 "메일 정리해줘"처럼 자연어로 요청하면
`organize_mail` 등의 도구가 자동으로 호출됩니다.

### 연결이 잘 됐는지 확인하기

가장 안전한 첫 호출은 `list_folders`입니다 (아무것도 이동시키지 않고 폴더 목록만 읽습니다). 대화에서
"outlook-mcp로 폴더 목록 보여줘"처럼 요청했을 때 실제 Outlook 폴더 구조가 출력되면 정상적으로
연결된 것입니다.

## 제공하는 Tool

| Tool | 설명 |
|---|---|
| `organize_mail` | 받은/보낸 편지함(`rootFolder`)을 검사해 고객사·카테고리 폴더로 자동 이동합니다. 읽지 않은 메일은 이동하지 않으며, 애매하게 판별된 메일은 옮기지 않고 "확인 필요" 목록으로만 보고합니다. `scope`(all/unread/today), `dryRun` 지원 |
| `read_mail` | `entryId`로 메일 한 건의 본문/제목/발신자/수신자/참조/날짜/첨부파일을 조회 |
| `search_mail` | 제목/본문/발신자/날짜/고객사/키워드로 메일 검색 (하위 폴더 포함) |
| `create_reply` | 원본 메일과 `reply_style.md`를 함께 조회하여 AI가 답장 본문을 작성할 수 있도록 준비 |
| `save_draft` | 작성된 답장(또는 새 메일)을 Outlook 임시보관함에 저장. `mode: "update"`로 기존 초안을 그 자리에서 덮어쓸 수도 있음. **발송하지 않음** |
| `list_drafts` | 임시보관함(Drafts)에 저장된 미완성 초안 목록 조회 (제목/키워드 필터, 최근 수정순). "쓰던 메일 이어서 완성해줘" 같은 요청에 사용 |
| `list_folders` | 받은 편지함 하위 폴더 트리와 메일/안읽음 수 조회 (디버깅용 보조 도구) |
| `learn_reply_style` | 보낸 편지함에서 실제 작성한 메일을 모아옵니다. AI가 이를 분석해 `reply_style.md` 초안을 작성하는 데 사용 (초기 설정용) |

### 사용 흐름 예시: 초기 설정 시 "내 메일 문체 학습해서 답장 스타일 만들어줘"

1. AI가 `learn_reply_style({ maxCount: 40 })`을 호출해 보낸 편지함에서 실제 작성한 메일을 가져옵니다.
2. 인사말/맺음말/서명/톤 패턴과 자주 쓰는 답장 유형을 분석합니다.
3. 분석 결과를 바탕으로 `config/reply_style.md`를 새로 작성합니다.
4. 이후 `create_reply`/`save_draft` 는 이 파일을 참고해 사용자의 실제 문체로 답장을 작성합니다.

### 사용 흐름 예시: "메일 정리해"

1. 사용자: "오늘 온 메일 정리해줘"
2. AI가 `organize_mail({ scope: "today" })`를 호출합니다 → 고객사·카테고리 폴더로 이동 후 결과
   요약을 반환합니다.

### 사용 흐름 예시: "고객사B 메일에 답장 초안 써줘"

1. AI가 `search_mail({ customer: "고객사B" })`로 관련 메일을 찾습니다.
2. `read_mail` 또는 `create_reply({ entryId })`로 원문과 답장 스타일 가이드를 확인합니다.
3. `reply_style.md` 기준으로 답장 본문(HTML)을 작성합니다.
4. `save_draft({ mode: "reply", sourceEntryId, bodyHtml })`를 호출해 임시보관함에 저장합니다.
5. 사용자가 Outlook에서 직접 검토한 후 발송합니다.

### 사용 흐름 예시: "임시보관함에 쓰던 메일 이어서 완성해줘"

1. AI가 `list_drafts()`를 호출해 임시보관함의 미완성 초안 목록(제목/받는사람/entryId)을 확인합니다.
2. 여러 건이면 사용자에게 어떤 초안인지 확인하거나 제목/키워드로 좁혀서 특정합니다.
3. `read_mail({ entryId })`로 지금까지 쓴 본문을 읽습니다.
4. 나머지 내용을 이어서 작성한 뒤, `save_draft({ mode: "update", draftEntryId, bodyHtml })`를
   호출해 새 초안을 만들지 않고 같은 초안에 덮어씁니다.
5. 사용자가 Outlook 임시보관함에서 검토한 후 직접 발송합니다.

## 프로젝트 구조

```
outlook-mcp/
├── config/
│   ├── customers.yml       # 고객사/별칭
│   ├── categories.yml      # 카테고리 규칙 + 기본 폴더
│   └── reply_style.md      # 답장 문체 가이드
├── src/
│   ├── server/index.ts     # MCP 서버 부트스트랩 (stdio transport)
│   ├── tools/               # MCP tool 8종 정의 (organize_mail, read_mail, search_mail,
│   │                         #   create_reply, save_draft, list_drafts, list_folders, learn_reply_style)
│   ├── outlook/
│   │   ├── client.ts        # OutlookClient (상위 레벨 API)
│   │   ├── powershellBridge.ts  # Node <-> PowerShell 프로세스 브리지
│   │   ├── types.ts         # zod 스키마 (COM 응답 검증)
│   │   └── scripts/*.ps1    # 실제 Outlook COM 호출 (PowerShell)
│   ├── config/               # customers.yml/categories.yml 로더 + zod 스키마
│   ├── classification/       # 제목/본문/참조 기반 분류 로직 (순수 함수)
│   └── utils/                 # 로거, 텍스트 정리 유틸
└── scripts/copy-assets.mjs   # 빌드 시 .ps1 스크립트를 dist/ 로 복사
```

## 개발

```bash
npm run dev         # tsx 로 src/server/index.ts 바로 실행 (컴파일 없이)
npm run typecheck   # tsc --noEmit
npm run build        # dist/ 로 빌드
npm start             # dist/server/index.js 실행
```

`LOG_LEVEL` 환경변수(`debug`/`info`/`warn`/`error`, 기본값 `info`)로 로그 상세도를 조절할 수 있습니다.
모든 로그는 MCP stdio 프로토콜과 충돌하지 않도록 **stderr**로만 출력합니다.

## 트러블슈팅

- **"폴더를 찾을 수 없습니다" 오류 없이도 메일이 이동하지 않는 경우**: Outlook이 실행 중인지, 그리고
  실행 중인 Outlook 프로필이 대상 메일함을 포함하는지 확인합니다.
- **답장 저장 시 수신자가 비어 있는 경우**: `mode: "new"`는 `to`가 필수입니다. `reply`/`replyAll`은
  원본 메일에서 자동으로 채워집니다.
- **한글이 깨져 보이는 경우**: PowerShell 콘솔에 직접 로그를 출력할 때 코드페이지 설정에 따라 콘솔
  표시가 깨질 수 있으나, 실제 stderr·파일 데이터는 UTF-8로 정상 기록됩니다. MCP 클라이언트는 파이프를
  통해 바이트를 직접 읽으므로 영향을 받지 않습니다.
- **다른 PowerShell(pwsh.exe)을 사용하려는 경우**: `OUTLOOK_MCP_POWERSHELL_EXE` 환경변수로 실행
  파일 경로를 지정할 수 있습니다.
- **폴더 이름에 "/"가 포함된 경우**: `folderPath`/`folder` 값은 "/"를 하위 폴더 구분자로 해석하므로,
  Outlook 폴더 이름 자체에 "/"가 들어 있으면(예: `공지/알림`) 경로 해석이 꼬입니다. 이런 폴더는 이름에서
  "/"를 빼거나, 해당 폴더를 `folder`/`folderPath` 값으로 직접 참조하지 않습니다.

## 라이선스

[MIT](./LICENSE)

---

<a id="english"></a>

## English

[![한국어](https://img.shields.io/badge/lang-한국어-lightgrey?style=social&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiMwMDAwMDAiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cGF0aCBzdHJva2U9Im5vbmUiIGQ9Ik0wIDBoMjR2MjRIMHoiIGZpbGw9Im5vbmUiIC8+PHBhdGggZD0iTTMgMTJhOSA5IDAgMSAwIDE4IDBhOSA5IDAgMCAwIC0xOCAwIiAvPjxwYXRoIGQ9Ik0zLjYgOWgxNi44IiAvPjxwYXRoIGQ9Ik0zLjYgMTVoMTYuOCIgLz48cGF0aCBkPSJNMTEuNSAzYTE3IDE3IDAgMCAwIDAgMTgiIC8+PHBhdGggZD0iTTEyLjUgM2ExNyAxNyAwIDAgMSAwIDE4IiAvPjwvc3ZnPg==)](#outlook-mcp)
[![English](https://img.shields.io/badge/lang-English-lightgrey?style=social&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9IiMwMDAwMDAiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cGF0aCBzdHJva2U9Im5vbmUiIGQ9Ik0wIDBoMjR2MjRIMHoiIGZpbGw9Im5vbmUiIC8+PHBhdGggZD0iTTMgMTJhOSA5IDAgMSAwIDE4IDBhOSA5IDAgMCAwIC0xOCAwIiAvPjxwYXRoIGQ9Ik0zLjYgOWgxNi44IiAvPjxwYXRoIGQ9Ik0zLjYgMTVoMTYuOCIgLz48cGF0aCBkPSJNMTEuNSAzYTE3IDE3IDAgMCAwIDAgMTgiIC8+PHBhdGggZD0iTTEyLjUgM2ExNyAxNyAwIDAgMSAwIDE4IiAvPjwvc3ZnPg==)](#english)

An [MCP](https://modelcontextprotocol.io) (Model Context Protocol) server that integrates with
Windows **Classic Outlook** (the desktop app). It uses **Outlook COM** (Object Model) directly
instead of the Microsoft Graph API, so it works without registering an Azure AD app or any cloud
authentication. It automates whatever Outlook is already installed locally.

The goal is to automatically sort the inbox by customer/category, search and read mail, and
prepare reply drafts in the user's own writing style inside Outlook's Drafts folder. **Mail is
always reviewed and sent by the user directly.** This server never sends mail on its own.

> To avoid being tied to any specific company, the customer list, categories, and reply style are
> all defined by files under `config/`. Editing those files is enough to reuse this project in a
> different organization, with no code changes required.

### How it works

Node.js cannot work with Windows COM objects directly, so this project delegates every
interaction with Outlook to a **PowerShell child process**.

```mermaid
flowchart TB
    accTitle: Outlook MCP Request Handling In Detail
    accDescr: Shows an MCP client request passing through temp JSON files from the outlook-mcp server to a PowerShell child process, which either attaches to a running Outlook instance or launches a new one, calls the COM API, and returns the result along the same path.

    subgraph client_layer ["MCP Client"]
        client["MCP Client<br/>(e.g. Claude)"]:::input_style
    end

    subgraph server_layer ["outlook-mcp server (Node.js/TypeScript)"]
        write_request[("Temp request JSON<br/>new temp folder per call")]:::process_style
        spawn_ps["Spawn PowerShell<br/>child process"]:::process_style
        read_response[("Temp response JSON")]:::process_style
        cleanup["Clean up temp folder"]:::aux_style
    end

    subgraph ps_layer ["PowerShell → Outlook COM"]
        check_running{"Outlook<br/>already running?"}
        attach_instance["Attach to running instance<br/>(no new window/profile prompt)"]:::process_style
        launch_instance["Launch new instance"]:::process_style
        outlook["Classic Outlook<br/>running instance"]:::output_style
        com_call["Outlook.Application<br/>COM call"]:::process_style
    end

    client -->|"stdio request (JSON-RPC)"| write_request
    write_request --> spawn_ps
    spawn_ps --> check_running
    check_running -->|yes| attach_instance
    check_running -->|no| launch_instance
    attach_instance --> outlook
    launch_instance --> outlook
    outlook --> com_call
    com_call --> read_response
    read_response -->|"stdio response (JSON-RPC)"| client
    read_response -.->|"on call end"| cleanup

    classDef input_style fill:#e9eeec,stroke:#4c5b60,stroke-width:2px,color:#2c3a3d
    classDef process_style fill:#232522,stroke:#111111,stroke-width:2px,color:#f5f3ec
    classDef output_style fill:#f6efde,stroke:#8a6f45,stroke-width:2px,color:#4a3b22
    classDef aux_style fill:none,stroke:#a85e1a,stroke-width:2px,stroke-dasharray:4 3,color:#7a4712
```

Request parameters and results are exchanged through **temporary JSON files** rather than stdout, to avoid locale/encoding issues.

### Requirements

- Windows 10/11
- Classic Outlook (desktop app; Microsoft 365 / Exchange / POP/IMAP accounts are all fine)
  installed with a profile configured
- Node.js 18 or later
- Windows PowerShell 5.1 (built into Windows; no separate install needed)

> New Outlook (the web-based rebuild) does not support the COM Object Model, so it cannot be
> controlled by this project. Your account needs to be on Classic Outlook.

### Installation

#### 1. Prerequisites

- Windows 10/11
- Classic Outlook desktop app installed and run at least once, with the account already set up
- [Node.js](https://nodejs.org) 18 or later (check with `node --version` in PowerShell)
- Windows PowerShell 5.1 (built into Windows; no separate install needed)

#### 2. Get the repository

```bash
git clone https://github.com/hayein-bit/outlook-mcp.git
cd outlook-mcp
```

If you don't have git, you can also download it as a ZIP from the repository page (**Code →
Download ZIP**) and extract it.

#### 3. Prepare the config files

Copy each `.example` file in the `config` folder to the same name without `.example`.

In PowerShell:

```powershell
Copy-Item config\customers.example.yml config\customers.yml
Copy-Item config\categories.example.yml config\categories.yml
Copy-Item config\reply_style.example.md config\reply_style.md
```

(On macOS/Linux or bash, use `cp config/customers.example.yml config/customers.yml` instead.)
Copying the files in File Explorer and removing `.example` from the name works the same way.

Open the three new files and fill them in for your own organization (customer list, categories,
reply style). You can also leave the example values as-is at first and just build to confirm
everything works.

#### 4. Install dependencies and build

```bash
npm install
npm run build
```

`npm run build` compiles TypeScript into `dist/` and copies `src/outlook/scripts/*.ps1` into
`dist/outlook/scripts/`. A successful build ends with a line similar to:

```
PowerShell script copy complete: ...\src\outlook\scripts -> ...\dist\outlook\scripts (9 files)
```

#### 5. Verify the build (optional)

With Outlook running, to quickly check that the server itself starts correctly:

```bash
npm start
```

A `outlook-mcp server started (stdio).` log line means it's working (`Ctrl+C` to stop). This
command alone doesn't invoke any tools. To actually use the server, connect it to an MCP client
as described below.

### Configuration (config/)

Each config file is split into a `*.example.*` template (public, tracked by git) and the real
file actually used (private, ignored by git). Anything containing real customer names or internal
domains always lives in the file without `.example` in its name.

| Template (tracked by git) | Real file (git-ignored) | Purpose |
|---|---|---|
| `config/customers.example.yml` | `config/customers.yml` | Customer list and aliases. Mail subject/body/CC are searched for these aliases to identify the customer. |
| `config/categories.example.yml` | `config/categories.yml` | Category rules (keywords/sender domains) for mail that isn't matched to a customer, and the folder each one moves to. |
| `config/reply_style.example.md` | `config/reply_style.md` | Tone/style guide always consulted when drafting a reply. Free-form Markdown. |

All three files can be freely edited, added to, or removed **without touching any code**.

#### customers.yml example

```yaml
customers:
  - name: CustomerA       # destination folder name (Inbox > 2. Customers > CustomerA)
    aliases:
      - CustomerA
      - Customer A Inc.

  - name: CustomerB
    aliases:
      - CustomerB
      - Customer B Inc.
```

> Customer detection is based on whether an alias appears in the subject, body, or CC, **not**
> on the sender's email domain. Even if an internal colleague sends the mail on the customer's
> behalf, it still moves to that customer's folder as long as "CustomerB" appears in the body.

#### categories.yml example

```yaml
categories:
  - name: Internal Notice
    folder: "3. Internal Notices"
    keywords: ["[Notice]", "[Announcement]", "company-wide notice", "HR notice"]
    sender_domains: []

default_folder: "7. Other"
customer_root_folder: "2. Customers"
calendar_folder: "6. Calendar & Meetings"
```

Rules are checked in order from top to bottom: customer detection first, then calendar-type items
(`calendar_folder`), then the categories below.

#### Default folder structure

```
Inbox
├── 1. Reference Material
├── 2. Customers
│   ├── CustomerA
│   ├── CustomerB
│   └── CustomerC
├── 3. Internal Notices
├── 4. Newsletters
├── 5. Work Discussion
├── 6. Calendar & Meetings
└── 7. Other
```

Folders that don't exist yet are created automatically when `organize_mail` runs. The numeric
prefixes in the example are just that: an example. Drop them from the `folder` /
`customer_root_folder` / `calendar_folder` / `default_folder` values if you don't want them.

#### reply_style.md

Freely define the level of formality, greeting/closing conventions, and signature format. The
`create_reply` tool always reads this entire file and passes it along for the AI to follow when
drafting a reply.

#### Changing the config location

By default, the server uses `./config` relative to the directory it was launched from (`cwd`). To
use a different location, set the `OUTLOOK_MCP_CONFIG_DIR` environment variable.

### Connecting an MCP client

This server runs as a local stdio process. It can only be reached by **an MCP client running on
the same PC where Outlook is installed.** A remote or cloud session has no way to reach Outlook
and won't work.

In the examples below, replace the path (`C:/outlook-mcp`) with wherever you actually cloned the
repository. On Windows, write backslashes (`\`) as forward slashes (`/`) or escape them (`\\`)
inside JSON or command lines.

#### Connecting Claude Desktop

1. Open the config file: `%APPDATA%\Claude\claude_desktop_config.json`
   (Paste `%APPDATA%\Claude` into File Explorer's address bar to jump straight to the folder. If
   the file doesn't exist, create it.)
2. Add the following under `mcpServers` (if the file is empty, you can paste this whole block in):

   ```json
   {
     "mcpServers": {
       "outlook-mcp": {
         "command": "node",
         "args": ["C:/outlook-mcp/dist/server/index.js"],
         "env": {
           "OUTLOOK_MCP_CONFIG_DIR": "C:/outlook-mcp/config"
         }
       }
     }
   }
   ```

3. Save the file, then **fully quit and restart Claude Desktop** (it needs to be closed from the
   system tray too for the change to take effect).
4. Open a new conversation and click the tools icon (🔨). If `outlook-mcp` shows up in the list,
   the connection is working.

#### Connecting the Claude Code CLI

`claude mcp add` is a command you run in a regular terminal (PowerShell/cmd), not inside a Claude
Code conversation. Register it once:

```bash
claude mcp add outlook-mcp --scope local -e OUTLOOK_MCP_CONFIG_DIR="C:/outlook-mcp/config" -- node "C:/outlook-mcp/dist/server/index.js"
```

Scope options:

- `--scope local` (default, recommended): the tool is only visible in the directory where you ran
  the command. Run it from inside the `outlook-mcp` folder to scope it to that project only.
- `--scope user`: visible from any directory, as long as it's the same Windows account on this PC.
- `--scope project`: the registration is stored inside the project folder itself and applies to
  anyone else working on that same project. Only use this if you're sharing the setup with a team.

Other options:

- `-e KEY=VALUE`: an environment variable passed to the server process. `OUTLOOK_MCP_CONFIG_DIR`
  isn't strictly required, but is recommended since Claude Code's working directory can vary.
- Everything after `--`: the actual command and arguments to run.

You only need to register once. After that, just open a new `claude` session as usual within the
registered scope, and the tools show up automatically.

Verify the registration:

```bash
claude mcp list              # check that outlook-mcp is listed
claude mcp get outlook-mcp   # inspect the registered command/env
```

Exact flags and behavior may vary by Claude Code version. Check `claude mcp add --help` if
something doesn't match. Once registered, open any new session and ask in natural language (e.g.
"organize my mail") to have `organize_mail` and the other tools invoked automatically.

#### Confirming the connection works

The safest first call is `list_folders` (it only reads the folder list and moves nothing). Ask
something like "show me the Outlook folder list via outlook-mcp". If the real folder structure
comes back, the connection is working.

### Available tools

| Tool | Description |
|---|---|
| `organize_mail` | Scans the inbox or sent folder (`rootFolder`) and automatically moves mail into customer/category folders. Unread mail is never moved, and mail that can't be confidently classified is left in place and reported as a "needs review" list instead of being moved. Supports `scope` (all/unread/today) and `dryRun`. |
| `read_mail` | Looks up a single mail's body/subject/sender/recipients/CC/date/attachments by `entryId`. |
| `search_mail` | Searches mail by subject/body/sender/date/customer/keyword (including subfolders). |
| `create_reply` | Fetches the original mail together with `reply_style.md` so the AI can draft a reply body. |
| `save_draft` | Saves a composed reply (or new mail) to Outlook's Drafts folder. `mode: "update"` overwrites an existing draft in place instead of creating a new one. **Never sends it.** |
| `list_drafts` | Lists unfinished drafts saved in the Drafts folder (subject/keyword filter, most recently modified first). Use this for requests like "finish the mail I was drafting." |
| `list_folders` | Lists the inbox's folder tree with mail/unread counts (debugging helper). |
| `learn_reply_style` | Gathers mail actually sent by the user, for the AI to analyze and draft `reply_style.md` from (used during initial setup). |

#### Example flow: initial setup, "Learn my writing style and build a reply style guide"

1. The AI calls `learn_reply_style({ maxCount: 40 })`, which pulls mail actually sent by the user.
2. It analyzes greeting/closing/signature/tone patterns and any recurring reply types.
3. Based on that analysis, it writes a new `config/reply_style.md`.
4. From then on, `create_reply` / `save_draft` follow that file to write replies in the user's
   actual voice.

#### Example flow: "Organize my mail"

1. User: "Organize the mail I got today."
2. The AI calls `organize_mail({ scope: "today" })`, which moves mail into customer/category
   folders and returns a summary.

#### Example flow: "Draft a reply to CustomerB's mail"

1. The AI calls `search_mail({ customer: "CustomerB" })` to find the relevant mail.
2. It calls `read_mail` or `create_reply({ entryId })` to see the original mail and the reply
   style guide.
3. It writes the reply body (HTML) following `reply_style.md`.
4. It calls `save_draft({ mode: "reply", sourceEntryId, bodyHtml })` to save it to Drafts.
5. The user reviews it in Outlook and sends it themselves.

#### Example flow: "Finish the mail I was drafting in the Drafts folder"

1. The AI calls `list_drafts()` to see the unfinished drafts (subject/recipient/entryId).
2. If there are several, it asks the user which one, or narrows it down by subject/keyword.
3. It calls `read_mail({ entryId })` to read what's been written so far.
4. It writes the rest of the body, then calls `save_draft({ mode: "update", draftEntryId, bodyHtml })`
   to overwrite that same draft instead of creating a new one.
5. The user reviews it in the Drafts folder and sends it themselves.

### Project structure

```
outlook-mcp/
├── config/
│   ├── customers.yml       # customers and aliases
│   ├── categories.yml      # category rules + default folder
│   └── reply_style.md      # reply style guide
├── src/
│   ├── server/index.ts     # MCP server bootstrap (stdio transport)
│   ├── tools/               # 8 MCP tool definitions (organize_mail, read_mail, search_mail,
│   │                         #   create_reply, save_draft, list_drafts, list_folders, learn_reply_style)
│   ├── outlook/
│   │   ├── client.ts        # OutlookClient (high-level API)
│   │   ├── powershellBridge.ts  # Node <-> PowerShell process bridge
│   │   ├── types.ts         # zod schemas (validate COM responses)
│   │   └── scripts/*.ps1    # the actual Outlook COM calls (PowerShell)
│   ├── config/               # customers.yml/categories.yml loader + zod schemas
│   ├── classification/       # subject/body/CC-based classification logic (pure functions)
│   └── utils/                 # logger, text cleanup utilities
└── scripts/copy-assets.mjs   # copies the .ps1 scripts into dist/ during build
```

### Development

```bash
npm run dev         # run src/server/index.ts directly with tsx (no compile step)
npm run typecheck   # tsc --noEmit
npm run build        # build into dist/
npm start             # run dist/server/index.js
```

Log verbosity is controlled by the `LOG_LEVEL` environment variable (`debug`/`info`/`warn`/`error`,
default `info`). All logs are written to **stderr only**, so they never interfere with the MCP
stdio protocol.

### Troubleshooting

- **Mail isn't moving, with no "folder not found" error**: Check that Outlook is running, and that
  the running Outlook profile actually contains the target mailbox.
- **Recipient is empty when saving a draft**: `mode: "new"` requires `to`. `reply`/`replyAll` fill
  it in automatically from the original mail.
- **Korean/other text looks garbled**: When logs are printed directly to a PowerShell console, the
  console's code page setting can make the display look garbled, but the actual stderr/file data
  is correctly written as UTF-8. MCP clients read the bytes directly through a pipe, so they are
  not affected.
- **Want to use a different PowerShell (pwsh.exe)**: set the `OUTLOOK_MCP_POWERSHELL_EXE`
  environment variable to that executable's path.
- **A folder name contains "/"**: `folderPath`/`folder` values treat "/" as a subfolder separator,
  so an actual Outlook folder whose name contains "/" (e.g. `Notices/Alerts`) will confuse the path
  resolution. Either remove the "/" from that folder's name, or avoid referencing it directly via
  `folder`/`folderPath`.

### License

[MIT](./LICENSE)


