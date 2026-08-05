# outlook-mcp

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](./LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18-339933?logo=node.js&logoColor=white)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](#요구사항)
[![MCP](https://img.shields.io/badge/MCP-Model%20Context%20Protocol-8A2BE2)](https://modelcontextprotocol.io)
[![GitHub Repo stars](https://img.shields.io/github/stars/hayein-bit/outlook-mcp?style=social)](https://github.com/hayein-bit/outlook-mcp)

Windows **Classic Outlook**(데스크톱 앱)과 연동되는 [MCP](https://modelcontextprotocol.io)(Model Context Protocol) 서버입니다.
Microsoft Graph API 대신 **Outlook COM(Object Model)**을 직접 사용하므로, 별도의 Azure AD 앱 등록이나
클라우드 인증 없이 로컬에 설치된 Outlook 을 그대로 자동화합니다.

받은 편지함을 회사/카테고리별로 자동 정리하고, 메일을 검색/조회하고, 나만의 문체로 답장 초안을
Outlook 임시보관함(Drafts)에 준비해 두는 것을 목표로 합니다. **메일은 항상 사용자가 직접 검토 후
발송**하며, 이 서버가 임의로 메일을 발송하는 기능은 제공하지 않습니다.

> 특정 회사에 종속되지 않도록, 고객사/카테고리/답장 스타일은 모두 `config/` 안의 파일만 수정하면
> 코드 변경 없이 다른 조직에서도 그대로 사용할 수 있게 설계되어 있습니다.

## 동작 원리

Node.js 에서는 Windows COM 객체를 직접 다룰 수 없기 때문에, 이 프로젝트는 Outlook 과의 모든 상호작용을
**PowerShell 자식 프로세스**에 위임합니다.

```
MCP Client (Claude 등)
      │  stdio (JSON-RPC)
      ▼
outlook-mcp (Node.js / TypeScript)
      │  임시 JSON 파일로 요청/응답 교환
      ▼
PowerShell (powershell.exe)
      │  Outlook.Application COM 객체
      ▼
Classic Outlook (실행 중인 인스턴스에 접속)
```

- 이미 실행 중인 Outlook 이 있으면 그 인스턴스에 그대로 붙습니다(새 창/프로필 선택 프롬프트 없음).
- 실행 중이 아니면 새로 실행합니다.
- 요청 파라미터와 결과는 로케일/인코딩 문제를 피하기 위해 stdout 이 아닌 **임시 JSON 파일**을 통해
  주고받습니다. 매 호출마다 임시 폴더를 생성하고 종료 시 정리합니다.

## 요구사항

- Windows 10/11
- Classic Outlook(데스크톱 앱, Microsoft 365 / Exchange / POP/IMAP 계정 무관) 설치 및 프로필 설정
- Node.js 18 이상
- Windows PowerShell 5.1 (Windows 기본 내장, 별도 설치 불필요)

> New Outlook(웹 기반 신규 Outlook)은 COM Object Model 을 지원하지 않으므로 이 프로젝트로 제어할 수
> 없습니다. Classic Outlook 으로 전환되어 있어야 합니다.

## 설치

### 1. 사전 준비물 확인

- Windows 10/11
- Classic Outlook 데스크톱 앱이 설치되어 있고, 한 번 이상 실행해서 계정 설정을 마친 상태
- [Node.js](https://nodejs.org) 18 이상 — PowerShell에서 `node --version` 으로 확인
- Windows PowerShell 5.1 — Windows에 기본 내장되어 있어 별도 설치가 필요 없습니다

### 2. 저장소 내려받기

```bash
git clone https://github.com/hayein-bit/outlook-mcp.git
cd outlook-mcp
```

git이 없다면 저장소 페이지의 **Code → Download ZIP** 으로 받아 압축을 풀어도 됩니다.

### 3. 설정 파일 준비

`config/` 안의 실제 사용 파일(`customers.yml`, `categories.yml`, `reply_style.md`)은 `.gitignore`
에 등록되어 있어 저장소에는 `*.example.*` 템플릿만 들어있습니다. PowerShell에서 아래처럼 복사해
실제 파일을 만드세요:

```powershell
Copy-Item config\customers.example.yml config\customers.yml
Copy-Item config\categories.example.yml config\categories.yml
Copy-Item config\reply_style.example.md config\reply_style.md
```

(macOS/Linux 나 bash 라면 `cp config/customers.example.yml config/customers.yml` 형태로 바꿔 쓰세요.)

방금 만든 3개 파일을 열어 자신의 조직(고객사 목록, 카테고리, 답장 문체)에 맞게 채워 넣으세요.
당장 안 채워도 예시(판다/수달/여우) 그대로 빌드해서 동작 확인만 먼저 해도 됩니다.

### 4. 의존성 설치 및 빌드

```bash
npm install
npm run build
```

`npm run build` 는 TypeScript 를 `dist/` 로 컴파일하고, `src/outlook/scripts/*.ps1` 을
`dist/outlook/scripts/` 로 복사합니다. 아래와 비슷한 줄이 마지막에 출력되면 성공입니다:

```
PowerShell 스크립트 복사 완료: ...\src\outlook\scripts -> ...\dist\outlook\scripts (8개 파일)
```

### 5. 빌드 확인 (선택)

Outlook을 실행해 둔 상태에서 서버 자체가 뜨는지만 빠르게 확인하려면:

```bash
npm start
```

`outlook-mcp 서버가 시작되었습니다 (stdio).` 로그가 뜨면 정상입니다(`Ctrl+C` 로 종료). 이 명령만으로는
어떤 도구도 호출되지 않습니다 — 실제로 쓰려면 아래처럼 MCP 클라이언트에 연결해야 합니다.

## 설정 (config/)

각 설정 파일은 `*.example.*` 템플릿(공개, git 추적)과 실제 사용 파일(비공개, git 무시)로
나뉘어 있습니다. 실제 고객사명/사내 도메인 등이 담기는 쪽은 항상 `.example` 이 없는 파일입니다.

| 템플릿 (git 추적) | 실사용 파일 (git 무시) | 역할 |
|---|---|---|
| `config/customers.example.yml` | `config/customers.yml` | 고객사 목록과 별칭. 메일 제목/본문/참조에서 별칭을 검색해 고객사를 판별합니다. |
| `config/categories.example.yml` | `config/categories.yml` | 고객사가 아닌 메일을 분류할 카테고리 규칙(키워드/발신자 도메인)과 이동할 폴더 경로. |
| `config/reply_style.example.md` | `config/reply_style.md` | 답장 작성 시 항상 참고하는 문체/톤 가이드. 자유 형식의 Markdown. |

세 파일 모두 **코드를 건드리지 않고** 그대로 수정/추가/삭제하면 됩니다.

### customers.yml 예시

```yaml
customers:
  - name: 판다       # 이동할 폴더 이름 (받은 편지함 > 고객사 > 판다)
    aliases:
      - 판다
      - PANDA

  - name: 수달
    aliases:
      - 수달
      - OTTER
```

> 고객사 판별은 **발신자 이메일 도메인이 아니라** 제목/본문/참조(CC)에 별칭이 등장하는지로
> 판단합니다. 내부 직원이 대신 보낸 메일이라도 본문에 "수달" 이 있으면 그 폴더로 이동합니다.

### categories.yml 예시

```yaml
categories:
  - name: 사내공지
    folder: 사내공지
    keywords: [사내공지, 전사공지, 인사발령]
    sender_domains: []

default_folder: 미분류
customer_root_folder: 고객사
```

규칙은 위에서부터 순서대로 검사되며, 고객사 판별이 카테고리 판별보다 항상 먼저 적용됩니다.

### 기본 폴더 구조

```
받은 편지함
├── 참고자료
├── 고객사
│   ├── 판다
│   ├── 수달
│   └── 여우
├── 사내공지
├── 뉴스레터
└── 미분류
```

존재하지 않는 폴더는 `organize_mail` 실행 시 자동으로 생성됩니다.

### reply_style.md

정중함의 정도, 인사말/맺음말 형식, 서명 등 답장 문체를 자유롭게 정의하세요. `create_reply` 도구는
이 파일 전체를 항상 함께 읽어 AI 가 답장을 작성할 때 참고하도록 전달합니다.

### config 위치 변경

기본적으로 서버를 실행한 디렉터리(`cwd`) 기준 `./config` 를 사용합니다. 다른 위치를 쓰려면
`OUTLOOK_MCP_CONFIG_DIR` 환경변수를 지정하세요.

## MCP 클라이언트에 연결하기

이 서버는 로컬 stdio 프로세스로 동작합니다. **Outlook이 설치된 이 PC에서 실행되는 MCP 클라이언트**
에만 연결할 수 있고, 원격/클라우드 세션에서는 Outlook에 접근할 수 없어 동작하지 않습니다.

아래 예시의 경로(`C:/outlook-mcp`)는 실제로 저장소를 내려받은 위치로 바꿔서 사용하세요. Windows
경로의 백슬래시(`\`)는 JSON/커맨드라인 안에서 슬래시(`/`)로 쓰거나 이스케이프(`\\`)해야 합니다.

### Claude Desktop 에 연결하기

1. 설정 파일을 엽니다: `%APPDATA%\Claude\claude_desktop_config.json`
   (탐색기 주소창에 `%APPDATA%\Claude` 를 붙여넣으면 폴더로 바로 이동합니다. 파일이 없으면 새로 만드세요.)
2. `mcpServers` 항목에 아래 내용을 추가합니다 (파일이 비어있다면 통째로 붙여넣으면 됩니다):

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

3. 파일을 저장한 뒤 **Claude Desktop을 완전히 종료했다가 다시 실행**합니다 (트레이 아이콘까지 종료해야
   반영됩니다).
4. 새 대화를 열고 도구(🔨) 아이콘을 눌러 `outlook-mcp` 가 목록에 보이면 연결된 것입니다.

### Claude Code CLI 에 연결하기

터미널에서 `claude mcp add` 명령으로 한 번만 등록하면 됩니다:

```bash
claude mcp add outlook-mcp --scope user -e OUTLOOK_MCP_CONFIG_DIR="C:/outlook-mcp/config" -- node "C:/outlook-mcp/dist/server/index.js"
```

옵션 설명:

- `--scope user` — 이 PC의 이 계정이라면 **어느 디렉터리에서 세션을 열어도** 도구가 보입니다 (권장).
- `--scope project` — `.mcp.json` 파일로 프로젝트 디렉터리에 저장되어, 그 프로젝트를 공유하는
  다른 사람에게도 함께 적용됩니다 (팀에서 같이 쓸 때).
- `-e KEY=VALUE` — 서버 프로세스에 넘길 환경변수. `OUTLOOK_MCP_CONFIG_DIR` 는 필수는 아니지만,
  Claude Code 실행 디렉터리가 매번 달라질 수 있으므로 명시해 두는 것을 권장합니다.
- `--` 뒤: 실제로 실행할 명령과 인자.

등록 후 확인:

```bash
claude mcp list              # outlook-mcp 가 목록에 보이는지
claude mcp get outlook-mcp   # 등록된 command/env 상세 확인
```

정확한 플래그/명령은 Claude Code 버전에 따라 달라질 수 있으니 `claude mcp add --help` 로 다시
확인하세요. 등록 이후에는 아무 디렉터리에서나 새 세션을 열고 "메일 정리해줘" 처럼 자연어로
요청하면 `organize_mail` 등의 도구가 자동으로 호출됩니다.

### 연결이 잘 됐는지 확인하기

가장 안전한 첫 호출은 `list_folders` 입니다 (아무것도 옮기지 않고 폴더 목록만 읽습니다). 대화에서
"outlook-mcp로 폴더 목록 보여줘" 처럼 요청했을 때 실제 Outlook 폴더 구조가 나오면 정상 연결된 것입니다.

## 제공하는 Tool

| Tool | 설명 |
|---|---|
| `organize_mail` | 받은/보낸 편지함(`rootFolder`)을 검사하여 고객사/카테고리 폴더로 자동 이동. 읽지 않은 메일은 이동하지 않고, 애매하게 판별된 메일은 옮기지 않고 "확인 필요" 목록으로만 보고합니다. `scope`(all/unread/today), `dryRun` 지원 |
| `read_mail` | `entryId` 로 메일 한 건의 본문/제목/발신자/수신자/참조/날짜/첨부파일을 조회 |
| `search_mail` | 제목/본문/발신자/날짜/고객사/키워드로 메일 검색 (하위 폴더 포함) |
| `create_reply` | 원본 메일과 `reply_style.md` 를 함께 조회하여 AI 가 답장 본문을 작성할 수 있도록 준비 |
| `save_draft` | 작성된 답장(또는 새 메일)을 Outlook 임시보관함에 저장. **발송하지 않음** |
| `list_folders` | 받은 편지함 하위 폴더 트리와 메일/안읽음 수 조회 (디버깅용 보조 도구) |

### 사용 흐름 예시: "메일 정리해"

1. 사용자: "오늘 온 메일 정리해줘"
2. AI 가 `organize_mail({ scope: "today" })` 호출 → 고객사/카테고리 폴더로 이동, 결과 요약 반환

### 사용 흐름 예시: "수달 메일에 답장 초안 써줘"

1. AI 가 `search_mail({ customer: "수달" })` 로 관련 메일을 찾음
2. `read_mail` 또는 `create_reply({ entryId })` 로 원문 + 답장 스타일 가이드를 확인
3. `reply_style.md` 기준으로 답장 본문(HTML)을 작성
4. `save_draft({ mode: "reply", sourceEntryId, bodyHtml })` 호출 → 임시보관함에 저장
5. 사용자가 Outlook 에서 직접 검토 후 발송

## 프로젝트 구조

```
outlook-mcp/
├── config/
│   ├── customers.yml       # 고객사/별칭
│   ├── categories.yml      # 카테고리 규칙 + 기본 폴더
│   └── reply_style.md      # 답장 문체 가이드
├── src/
│   ├── server/index.ts     # MCP 서버 부트스트랩 (stdio transport)
│   ├── tools/               # MCP tool 5+1종 정의
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

`LOG_LEVEL` 환경변수(`debug`/`info`/`warn`/`error`, 기본 `info`)로 로그 상세도를 조절할 수 있습니다.
모든 로그는 MCP stdio 프로토콜과 충돌하지 않도록 **stderr** 로만 출력됩니다.

## 트러블슈팅

- **"폴더를 찾을 수 없습니다" 오류 없이도 메일이 이동하지 않음**: Outlook 이 실행 중인지, 그리고
  실행 중인 Outlook 프로필이 대상 메일함을 포함하는지 확인하세요.
- **답장 저장 시 수신자가 비어 있음**: `mode: "new"` 는 `to` 가 필수입니다. `reply`/`replyAll` 은
  원본 메일에서 자동으로 채워집니다.
- **한글이 깨져 보임**: PowerShell 콘솔에 직접 로그를 출력해 볼 때 코드페이지 설정에 따라 콘솔
  표시가 깨질 수 있으나, 실제 stderr/파일 데이터는 UTF-8 로 정상 기록됩니다. MCP 클라이언트는
  파이프를 통해 바이트를 직접 읽으므로 영향을 받지 않습니다.
- **다른 PowerShell(pwsh.exe)을 쓰고 싶다면**: `OUTLOOK_MCP_POWERSHELL_EXE` 환경변수로 실행 파일
  경로를 지정할 수 있습니다.
- **폴더 이름에 "/" 가 포함된 경우**: `folderPath`/`folder` 값은 "/" 를 하위 폴더 구분자로
  해석하므로, Outlook 폴더 이름 자체에 "/" 가 들어 있으면(예: `공지/알림`) 경로 해석이 꼬입니다.
  이런 폴더는 이름에서 "/" 를 빼거나, 해당 폴더를 `folder`/`folderPath` 값으로 직접 참조하지 마세요.

## 라이선스

[MIT](./LICENSE)
