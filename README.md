# mail-watch

클래식 Outlook으로 새 메일을 읽고 [Claude Code](https://claude.com/claude-code) CLI로 분류한 뒤 Windows 토스트로 알려 주는 PowerShell 스크립트입니다.

## 하는 일

- 받은 편지함과 번호가 붙은 규칙 폴더(`00.`, `10.` 등과 그 하위 폴더, `9x`로 시작하는 폴더는 제외)에서 새로 온 **안 읽은 메일**을 모읍니다.
- 메일을 한 번에 Claude에 보내 다음 중 하나로 분류합니다.

  | 분류 | 의미 | 처리 |
  |------|------|------|
  | `PHISHING` | 피싱·사기 의심 | 지운 편지함으로 이동하고 알림 |
  | `ACTION` | 내가 처리해야 할 일 | 열린 항목으로 추적 |
  | `CHECK` | 할 일은 없지만 읽어야 할 메일 | 열린 항목으로 추적 |
  | `INFO` | 업무 관련이지만 건너뛰어도 되는 메일 | 열린 항목으로 추적 |
  | `NONE` | 뉴스레터, 광고, 자동 알림 | 무시 |

- 분류별 건수를 토스트로 띄웁니다. 토스트의 *Open status*를 누르면 `status.md`가 열립니다.
- 열린 항목은 `items.json`에 남고 `status.md`로 정리됩니다. 메일을 읽거나, 답장·전달하거나, 삭제·이동하거나, 같은 스레드에 이후 메일이 오면 목록에서 빠집니다.
- 모든 분류 결과는 `log\triage-<날짜>.md`에, 오류는 `log\error.md`에 남습니다.

### 안전장치

- 메일 내용은 신뢰하지 않는 데이터로 다룹니다. Claude는 도구를 하나도 쓰지 못하는 상태(`--tools ""`, `--strict-mcp-config`)로 실행되므로, 악성 메일이 지시문을 넣어도 텍스트 답변 말고는 아무 일도 일으킬 수 없습니다.
- 분석이 실패하면 메일을 확인한 것으로 표시하지 않습니다. 다음 실행 때 다시 처리합니다.
- Outlook이 켜지는 중이라 COM 호출을 거부하면(`RPC_E_CALL_REJECTED`) 10초 간격으로 최대 2분 동안 다시 시도합니다.

## 요구사항

- Windows 10/11, Windows PowerShell 5.1
- **클래식 Outlook** 데스크톱 앱과 로그인된 프로필 (새 Outlook과 웹 Outlook은 COM을 지원하지 않습니다)
- npm으로 설치하고 로그인까지 마친 Claude Code CLI (`%APPDATA%\npm\claude.cmd`)

> 메일 본문(앞부분 최대 3000자)과 헤더가 Claude API로 전송됩니다. 회사 메일에 쓰기 전에 사내 보안 정책을 확인하세요.

## 설치

1. 이 저장소를 원하는 폴더에 받습니다.
2. `config.example.txt`를 `config.txt`로 복사하고 내용을 채웁니다. 이 정보는 분류 프롬프트에 들어갑니다.
   ```
   OWNER=you@example.com
   ROLE=IT developer maintaining the XYZ system
   ```
3. 필요하면 `labels.txt`에서 토스트 문구를 바꿉니다(UTF-8).
4. 테스트: Outlook에서 메일 하나를 `.msg`로 저장한 뒤 실행합니다. 결과는 `log\test.md`에 남고 상태 파일은 바뀌지 않습니다.
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\watch.ps1 -MsgFile .\sample.msg
   ```
5. 작업 스케줄러에 등록합니다(예: 15분마다). `run-hidden.vbs`를 거쳐 실행하면 콘솔 창이 뜨지 않습니다.
   ```powershell
   schtasks /Create /TN MailWatch /SC MINUTE /MO 15 /TR "wscript.exe \"C:\path\to\mail-watch\run-hidden.vbs\""
   ```

처음 실행하면 오늘 0시 이후 메일부터 훑습니다. 그다음부터는 마지막 실행 30분 전부터 훑고, 이미 본 메일은 EntryID로 걸러 냅니다.

## 파일

| 파일 | 설명 |
|------|------|
| `watch.ps1` | 본체. ASCII로만 작성해야 합니다(PowerShell 5.1은 BOM 없는 스크립트를 ANSI로 읽음). |
| `run-hidden.vbs` | 작업 스케줄러용 실행기. 창 없이 `watch.ps1`을 실행합니다. |
| `labels.txt` | 토스트·상태 문구 |
| `config.txt` | 메일함 주인 정보 (git 제외) |
| `seen.txt`, `items.json`, `status.md`, `last-run.txt`, `log\` | 실행 중에 생기는 상태와 로그. 실제 메일 정보가 들어 있으므로 git에서 제외합니다. |
