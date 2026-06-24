# daily_briefing 도구 설계

- **날짜**: 2026-06-24
- **대상 프로젝트**: `ms-teams-mcp`
- **파일**: `ms_teams_mcp/server.py` (단일 파일 MCP 서버)

## 목표

회사에서 지금 무슨 일이 돌아가는지를 한 번의 호출로 파악하게 한다. 흩어진 조회 도구를
사람이 일일이 돌려 종합하던 일을, 여러 소스를 가로질러 종합하는 단일 읽기 도구로 대체한다.

`daily_briefing()` — 오늘 일정 + 안 읽은 메일 + 최근 채팅 + 채널 활동 + 응답 대기 항목을
하나의 브리핑 문자열로 반환한다.

## 설계 결정 (사용자 확정)

| 항목 | 결정 |
|------|------|
| 포함 섹션 | 오늘 일정 / 안 읽은 메일 / 최근 채팅 / 채널 활동 / 응답 대기 (5종) |
| 채널 비용 제한 | 시간창(기본 24h) + 하드캡(순회 채널 수, 채널당 메시지 수) |
| 응답 대기 정의 | 플래그된 메일 + 시간창 내 내가 @멘션된 채널 글 |
| 구현 방식 | 단일 도구 + 섹션별 private 헬퍼. 채널은 1회만 스캔해 활동·멘션을 동시 도출 |
| 권한 | 신규 스코프 불필요 (기존 스코프로 충분) |
| 확인 정책 | 읽기 전용 → 사용자 확인 불필요 |

## 아키텍처

기존 `get_unread_summary`의 패턴을 따른다: 섹션별로 `parts[]`에 문자열을 쌓고, 각 섹션을
`try/except`로 감싸 한 소스가 실패해도 나머지는 정상 출력(graceful degradation)한다.
모든 Graph 호출은 기존 `graph_get` 헬퍼를 거친다.

### 도구

```python
@mcp.tool()
def daily_briefing(hours: int = 24, max_channels: int = 20, max_messages_per_channel: int = 10) -> str:
    """Aggregate today's schedule, unread mail, recent chats, channel activity,
    and items needing your response into a single briefing."""
```

- `hours`: "recent" 창. 채널 글·@멘션 대상. 기본 24.
- `max_channels`: 전체 순회 채널 상한(레이트리밋 방지). 기본 20.
- `max_messages_per_channel`: 채널당 조회 메시지 상한. 기본 10.

### 섹션 헬퍼 (private, `@mcp.tool()` 아님)

| # | 섹션 헤더 | 헬퍼 | Graph 엔드포인트 |
|---|-----------|------|------------------|
| 1 | `Today's Schedule` | `_briefing_calendar_today()` | `/me/calendarView` (오늘 00:00~23:59Z) |
| 2 | `Unread Mail` | `_briefing_unread_email()` | `/me/mailFolders/inbox` ($select unreadItemCount,totalItemCount) |
| 3 | `Recent Chats` | `_briefing_recent_chats()` | `/me/chats` ($expand members, orderby lastMessagePreview/createdDateTime desc) |
| 4 | `Channel Activity` | `_scan_channels(...)` 결과의 활동 부분 | joinedTeams → channels → messages |
| 5 | `Needs Your Response` | `_briefing_flagged_email()` + `_scan_channels(...)`의 멘션 부분 | `/me/messages?$filter=flag/flagStatus eq 'flagged'` + 채널 스캔 |

각 헬퍼는 한 섹션의 문자열(또는 리스트)을 만들고, 내부에서 실패를 잡아
`"<Section>: Failed to retrieve ({e})"`로 degrade한다.

### 채널 1회 스캔

```python
def _scan_channels(hours: int, max_channels: int, max_messages_per_channel: int, me_id: str) -> tuple[list, list]:
    """Walk joined teams -> channels -> recent messages once.
    Returns (activity_lines, mention_lines): recent posts within the window, and
    posts where the current user (me_id) is @mentioned. Bounded by max_channels.
    """
```

- `/me/joinedTeams` → 각 팀 `/teams/{tid}/channels` → 각 채널
  `/teams/{tid}/channels/{cid}/messages` ($top = max_messages_per_channel).
- 순회한 채널 누적이 `max_channels`에 도달하면 중단하고, 그 사실을 호출부가 출력에 표기한다
  (silent truncation 금지: `"… scanned {K}/{total} channels (capped)"`).
- `createdDateTime`이 `hours` 창 밖이면 제외.
- 멘션 판정: 메시지 `mentions[].mentioned.user.id == me_id`.
- 팀 루프·채널 루프 내부에도 `try/except` → 한 채널의 403/404가 전체 스캔을 죽이지 않는다.
- 활동 섹션(4)과 멘션(5)이 **이 단일 반환값을 공유**한다(중복 순회 없음).

## 데이터 흐름 (`daily_briefing` 본문)

```
me_id = graph_get("/me")["id"]        # 멘션 판정용 1회. 실패 시 me_id=None → 멘션 판정만 생략
parts = [f"Daily Briefing — {today} (last {hours}h)"]
parts += _briefing_calendar_today()
parts += _briefing_unread_email()
parts += _briefing_recent_chats()
activity, mentions = _scan_channels(hours, max_channels, max_messages_per_channel, me_id)
parts += render Channel Activity (activity, + capped note if any)
parts += render Needs Your Response (_briefing_flagged_email() + mentions)
return "\n".join(parts)
```

- 캘린더는 "오늘"(날짜 기준), `hours` 창과 무관.
- 플래그 메일은 시간창과 무관(모든 플래그).

## 에러 처리

- 섹션별 `try/except`로 부분 실패 허용. 실패 섹션만 `"Failed to retrieve ({e})"`.
- `_scan_channels`는 팀/채널 단위 `try/except`로 격리.
- `/me` 조회 실패 → 멘션 판정 불가 → 멘션 부분 `"(mention detection unavailable)"`, 활동 섹션은 계속.
- 상한 도달 시 채널 섹션 끝에 capped 표기.

## 출력 형식

단일 문자열, 영어 섹션 헤더. 빈 섹션은 명시(`"No events today."`, `"None"`).

```
Daily Briefing — 2026-06-24 (last 24h)
── Today's Schedule ──
1. ...
── Unread Mail ──
Inbox: 7 unread / 213 total
── Recent Chats ──
1. ...
── Channel Activity ──
1. [Team / Channel] ...
   (scanned 18/41 channels, capped)
── Needs Your Response ──
Flagged emails (2): ...
Channel @mentions (3): ...
```

## 권한 (스코프)

신규 스코프 불필요. 사용하는 기존 스코프: `Calendars.ReadWrite`, `Mail.Read`,
`Chat.Read`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `User.Read`.

## 테스트 — pytest (기존 mock 패턴)

`tests/test_server.py`에 `graph_get`를 `unittest.mock.patch`로 가로채는 패턴으로 신규
테스트 클래스를 추가한다. `graph_get`은 엔드포인트(path)별로 staged 응답을 돌려주는
side_effect로 모킹한다. assert는 신규 코드가 반환하는 영어 출력 기준.

검증 대상:
- 정상 경로: 5개 섹션 헤더가 모두 등장, 각 섹션에 staged 데이터 반영.
- 부분 실패: 한 섹션의 `graph_get`이 raise → 그 섹션만 "Failed to retrieve", 나머지 정상.
- `_scan_channels`: 시간창 밖 메시지 제외 / `max_channels` 상한 준수(+capped 표기) /
  `me_id`와 일치하는 멘션만 추출 / 한 채널 예외가 스캔 전체를 죽이지 않음.
- `/me` 실패 시 멘션 degrade 문구.
- 빈 데이터 시 빈-섹션 문구.

신규 테스트만 추가하며 기존 stale 테스트(한국어 assert, 사전 13개 실패)는 손대지 않는다.
실행은 venv(`.venv/bin/python -m pytest`).

## 범위 밖 (이번에 하지 않음)

- 채널 글/이메일 본문 LLM 요약
- "since last briefing" 상태 저장(증분 브리핑)
- 우선순위 스코어링/정렬
- 지정 채널 우선 화이트리스트(설정 파일)
