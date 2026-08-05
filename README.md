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

```bash
git clone <this-repo>
cd outlook-mcp
npm install
npm run build
```

`npm run build` 는 TypeScript 를 `dist/` 로 컴파일하고, `src/outlook/scripts/*.ps1` 을
`dist/outlook/scripts/` 로 복사합니다.

## 설정 (config/)

| 파일 | 역할 |
|---|---|
| `config/customers.yml` | 고객사 목록과 별칭. 메일 제목/본문/참조에서 별칭을 검색해 고객사를 판별합니다. |
| `config/categories.yml` | 고객사가 아닌 메일을 분류할 카테고리 규칙(키워드/발신자 도메인)과 이동할 폴더 경로. |
| `config/reply_style.md` | 답장 작성 시 항상 참고하는 문체/톤 가이드. 자유 형식의 Markdown. |

세 파일 모두 **코드를 건드리지 않고** 그대로 수정/추가/삭제하면 됩니다.

### customers.yml 예시

```yaml
customers:
  - name: 아크미시청       # 이동할 폴더 이름 (받은 편지함 > 고객사 > 아크미시청)
    aliases:
      - 아크미시청
      - 아크미

  - name: 우리동네주택공사
    aliases:
      - 우리동네주택공사
      - ONHC
```

> 고객사 판별은 **발신자 이메일 도메인이 아니라** 제목/본문/참조(CC)에 별칭이 등장하는지로
> 판단합니다. 내부 직원이 대신 보낸 메일이라도 본문에 "우리동네주택공사" 가 있으면 그 폴더로 이동합니다.

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
│   ├── 아크미시청
│   ├── 우리동네주택공사
│   └── 별빛지도연구원
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

## MCP 클라이언트에 등록하기

Claude Desktop 설정 파일(`claude_desktop_config.json`) 예시:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "node",
      "args": ["C:/path/to/outlook-mcp/dist/server/index.js"],
      "env": {
        "OUTLOOK_MCP_CONFIG_DIR": "C:/path/to/outlook-mcp/config"
      }
    }
  }
}
```

경로의 백슬래시(`\`)는 슬래시(`/`)로 쓰거나 이스케이프(`\\`)해야 합니다.

## 제공하는 Tool

| Tool | 설명 |
|---|---|
| `organize_mail` | 받은 편지함을 검사하여 고객사/카테고리 폴더로 자동 이동. `scope`(all/unread/today), `dryRun` 지원 |
| `read_mail` | `entryId` 로 메일 한 건의 본문/제목/발신자/수신자/참조/날짜/첨부파일을 조회 |
| `search_mail` | 제목/본문/발신자/날짜/고객사/키워드로 메일 검색 (하위 폴더 포함) |
| `create_reply` | 원본 메일과 `reply_style.md` 를 함께 조회하여 AI 가 답장 본문을 작성할 수 있도록 준비 |
| `save_draft` | 작성된 답장(또는 새 메일)을 Outlook 임시보관함에 저장. **발송하지 않음** |
| `list_folders` | 받은 편지함 하위 폴더 트리와 메일/안읽음 수 조회 (디버깅용 보조 도구) |

### 사용 흐름 예시 — "메일 정리해"

1. 사용자: "오늘 온 메일 정리해줘"
2. AI 가 `organize_mail({ scope: "today" })` 호출 → 고객사/카테고리 폴더로 이동, 결과 요약 반환

### 사용 흐름 예시 — "우리동네주택공사 메일에 답장 초안 써줘"

1. AI 가 `search_mail({ customer: "우리동네주택공사" })` 로 관련 메일을 찾음
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

## 라이선스

[MIT](./LICENSE)
