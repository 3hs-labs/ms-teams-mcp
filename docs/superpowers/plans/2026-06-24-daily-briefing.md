# daily_briefing 도구 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** ms-teams-mcp 서버에 단일 읽기 도구 `daily_briefing()`을 추가해 오늘 일정·안 읽은 메일·최근 채팅·채널 활동·응답 대기 항목을 하나의 브리핑으로 종합한다.

**Architecture:** `get_unread_summary`의 패턴(섹션별 `parts[]` + `try/except` graceful degradation)을 확장한다. 단순 섹션 4개는 private 헬퍼로, 채널은 `_scan_channels`로 1회만 순회해 "활동"과 "@멘션"을 동시에 도출한다. 모든 Graph 호출은 기존 `graph_get` 헬퍼를 거친다.

**Tech Stack:** Python 3, FastMCP, Microsoft Graph API, pytest + `unittest.mock`.

## Global Constraints

- **영어 전용 소스**: 모든 주석·docstring·출력 문자열은 영어로만 (`ms-teams-mcp/CLAUDE.md`).
- **모든 Graph 호출은 `graph_get` 헬퍼 경유**.
- **읽기 전용 도구**: 사용자 확인 정책 무관, 신규 스코프 불필요.
- **부분 실패 허용**: 각 섹션은 독립 `try/except`로 degrade(`"Failed to retrieve ({e})"`), 한 소스 장애가 전체를 죽이지 않는다.
- **silent truncation 금지**: 채널 상한 도달 시 `"(scanned K/total channels, capped)"` 표기.
- **테스트**: `tests/test_server.py`의 기존 mock 패턴(`graph_get`을 `unittest.mock.patch`)을 따라 TDD. 신규 테스트만 추가하며 기존 stale 13개는 손대지 않는다. 실행은 venv: `.venv/bin/python -m pytest`.
- **테스트 assert는 영어 출력 기준**.

---

## File Structure

- **Modify** `ms_teams_mcp/server.py`: `get_unread_summary`(끝 ~line 1940) 뒤, `# Auth Management` 구분선(~line 1942) 앞에 private 헬퍼 5개 + `@mcp.tool() daily_briefing` 추가.
- **Modify** `tests/test_server.py`: 헬퍼·도구 테스트 클래스 추가.
- **Modify** `ms-teams-mcp/CLAUDE.md` / `README.md` / `TODO.md`: 도구 수·도구 목록 갱신.

## Task 0 — 환경 (이미 완료, 참고)

venv는 `ms-teams-mcp/.venv`에 이미 존재하고 의존성·pytest 설치됨(`.venv`는 gitignore). 모든 테스트는 `.venv/bin/python -m pytest`로 실행한다. 베이스라인 전체 스위트: `13 failed, 26 passed`(13개는 기존 stale, 손대지 않음). 신규 테스트는 모두 통과해야 한다.

---

## Task 1: 단순 섹션 헬퍼 4개

오늘 일정 / 안 읽은 메일 / 최근 채팅 / 플래그 메일. 각 헬퍼는 한 섹션 문자열을 만들고 내부 `try/except`로 degrade한다.

**Files:**
- Modify: `ms_teams_mcp/server.py` (insert after `get_unread_summary`, ~line 1940)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `graph_get`, `strip_html`, `datetime`, `timezone` (모두 server.py에 이미 import됨)
- Produces:
  - `_briefing_calendar_today() -> str`
  - `_briefing_unread_email() -> str`
  - `_briefing_recent_chats() -> str`
  - `_briefing_flagged_email() -> str`

- [ ] **Step 1: Write the failing tests**

`tests/test_server.py`에 import 추가 후 테스트 클래스 추가:
```python
from ms_teams_mcp.server import (
    _briefing_calendar_today,
    _briefing_unread_email,
    _briefing_recent_chats,
    _briefing_flagged_email,
)

class TestBriefingSimpleSections:
    @patch("ms_teams_mcp.server.graph_get")
    def test_calendar_today(self, mock_get):
        mock_get.return_value = {"value": [
            {"subject": "Standup", "start": {"dateTime": "2026-06-24T09:00:00"},
             "location": {"displayName": "Room A"}}
        ]}
        result = _briefing_calendar_today()
        assert "Today's Schedule" in result
        assert "Standup" in result
        assert "Room A" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_calendar_empty(self, mock_get):
        mock_get.return_value = {"value": []}
        assert "No events today." in _briefing_calendar_today()

    @patch("ms_teams_mcp.server.graph_get")
    def test_calendar_failure_degrades(self, mock_get):
        mock_get.side_effect = Exception("boom")
        result = _briefing_calendar_today()
        assert "Today's Schedule" in result
        assert "Failed to retrieve" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_unread_email(self, mock_get):
        mock_get.return_value = {"unreadItemCount": 7, "totalItemCount": 213}
        result = _briefing_unread_email()
        assert "7 unread / 213 total" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_recent_chats(self, mock_get):
        mock_get.return_value = {"value": [
            {"topic": "Project X", "members": [{"displayName": "A"}],
             "lastMessagePreview": {"body": {"content": "<p>hi</p>"},
                                    "createdDateTime": "2026-06-24T08:00:00Z"}}
        ]}
        result = _briefing_recent_chats()
        assert "Recent Chats" in result
        assert "Project X" in result
        assert "hi" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_flagged_email(self, mock_get):
        mock_get.return_value = {"value": [
            {"subject": "Reply needed", "from": {"emailAddress": {"name": "Bob"}}}
        ]}
        result = _briefing_flagged_email()
        assert "Flagged emails (1):" in result
        assert "Reply needed" in result
        assert "Bob" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_flagged_email_none(self, mock_get):
        mock_get.return_value = {"value": []}
        assert "Flagged emails: none" in _briefing_flagged_email()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestBriefingSimpleSections -v`
Expected: FAIL — `ImportError: cannot import name '_briefing_calendar_today'`

- [ ] **Step 3: Implement the helpers**

`ms_teams_mcp/server.py`에서 `get_unread_summary` 함수 끝(~line 1940) 뒤에 삽입:
```python
def _briefing_calendar_today() -> str:
    """Return today's calendar events as a briefing section."""
    header = "── Today's Schedule ──"
    try:
        today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        data = graph_get("/me/calendarView", params={
            "startDateTime": f"{today}T00:00:00Z",
            "endDateTime": f"{today}T23:59:59Z",
            "$select": "subject,start,end,location",
            "$orderby": "start/dateTime",
            "$top": 20,
        })
        events = data.get("value", [])
        if not events:
            return f"{header}\nNo events today."
        lines = [header]
        for i, ev in enumerate(events, 1):
            subject = ev.get("subject", "(No subject)")
            start_time = ev.get("start", {}).get("dateTime", "")[:16]
            location = ev.get("location", {}).get("displayName", "")
            loc = f" @ {location}" if location else ""
            lines.append(f"{i}. [{start_time}] {subject}{loc}")
        return "\n".join(lines)
    except Exception as e:
        return f"{header}\nFailed to retrieve ({e})"

def _briefing_unread_email() -> str:
    """Return the inbox unread/total count as a briefing section."""
    header = "── Unread Mail ──"
    try:
        inbox = graph_get("/me/mailFolders/inbox", params={"$select": "unreadItemCount,totalItemCount"})
        unread = inbox.get("unreadItemCount", 0)
        total = inbox.get("totalItemCount", 0)
        return f"{header}\nInbox: {unread} unread / {total} total"
    except Exception as e:
        return f"{header}\nFailed to retrieve ({e})"

def _briefing_recent_chats() -> str:
    """Return the most recent chats as a briefing section."""
    header = "── Recent Chats ──"
    try:
        chats = graph_get("/me/chats", params={
            "$top": 10,
            "$expand": "members",
            "$orderby": "lastMessagePreview/createdDateTime desc",
        })
        chat_list = chats.get("value", [])
        if not chat_list:
            return f"{header}\nNo recent chats."
        lines = [header]
        for i, c in enumerate(chat_list, 1):
            topic = c.get("topic") or ""
            members = [m.get("displayName") or "" for m in c.get("members", [])]
            label = topic if topic else ", ".join(members[:4])
            preview = c.get("lastMessagePreview", {}) or {}
            body = strip_html(preview.get("body", {}).get("content", ""))[:60]
            ts = preview.get("createdDateTime", "")[:16]
            lines.append(f"{i}. {label}\n   [{ts}] {body}")
        return "\n".join(lines)
    except Exception as e:
        return f"{header}\nFailed to retrieve ({e})"

def _briefing_flagged_email() -> str:
    """Return the flagged-emails portion of the 'Needs Your Response' section.

    No $orderby is combined with the flag filter — Graph rejects that pairing
    as too complex on the messages collection.
    """
    try:
        data = graph_get("/me/messages", params={
            "$filter": "flag/flagStatus eq 'flagged'",
            "$select": "subject,from",
            "$top": 15,
        })
        msgs = data.get("value", [])
        if not msgs:
            return "Flagged emails: none"
        lines = [f"Flagged emails ({len(msgs)}):"]
        for m in msgs:
            subject = m.get("subject", "(No subject)")
            sender = m.get("from", {}).get("emailAddress", {}).get("name", "Unknown")
            lines.append(f"  - {subject} — {sender}")
        return "\n".join(lines)
    except Exception as e:
        return f"Flagged emails: Failed to retrieve ({e})"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestBriefingSimpleSections -v`
Expected: PASS (7 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(briefing): add simple section helpers (calendar/unread/chats/flagged)"
```

---

## Task 2: `_scan_channels` 채널 1회 스캔

가입 팀→채널→최근 메시지를 1회 순회해 활동 글과 내 @멘션을 동시에 수집한다. 시간창·채널 상한으로 제한하고, 팀/채널 단위 `try/except`로 격리한다.

**Files:**
- Modify: `ms_teams_mcp/server.py` (insert after `_briefing_flagged_email`)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `graph_get`, `strip_html`, `datetime`, `timezone`, `timedelta` (모두 import됨)
- Produces: `_scan_channels(hours: int, max_channels: int, max_messages_per_channel: int, me_id: str | None) -> tuple[list, list, int, int]` — 반환 `(activity, mentions, scanned, total)`: 시간창 내 활동 글 리스트, 내 멘션 글 리스트, 실제 스캔한 채널 수, 전체 채널 수. `total > scanned`이면 상한에 걸린 것.

- [ ] **Step 1: Write the failing tests**

채널 메시지의 `createdDateTime`을 "시간창 안"으로 만들기 위해 **고정 과거 시각** 대신, 테스트는 `datetime`을 직접 다루지 않고 매우 최근(now에 가까운) ISO 문자열을 쓰면 24h 창 안에 든다. 창 밖 검증은 먼 과거(2000년) 문자열을 쓴다.
```python
from ms_teams_mcp.server import _scan_channels

def _channel_side_effect(messages_by_channel):
    """Return a graph_get side_effect that routes by path."""
    def _se(path, params=None, url=None):
        if path == "/me/joinedTeams":
            return {"value": [{"id": "t1", "displayName": "Team1"}]}
        if path == "/teams/t1/channels":
            return {"value": [
                {"id": "c1", "displayName": "General"},
                {"id": "c2", "displayName": "Random"},
            ]}
        if path.startswith("/teams/t1/channels/"):
            cid = path.split("/")[4]
            return {"value": messages_by_channel.get(cid, [])}
        return {"value": []}
    return _se

class TestScanChannels:
    @patch("ms_teams_mcp.server.graph_get")
    def test_collects_recent_activity(self, mock_get):
        recent = "2999-01-01T00:00:00Z"  # always within window
        mock_get.side_effect = _channel_side_effect({
            "c1": [{"id": "m1", "createdDateTime": recent,
                    "from": {"user": {"displayName": "Alice"}},
                    "body": {"content": "<p>hello</p>"}, "mentions": []}],
            "c2": [],
        })
        activity, mentions, scanned, total = _scan_channels(24, 20, 10, "me")
        assert any("Alice" in a and "hello" in a for a in activity)
        assert scanned == 2 and total == 2
        assert mentions == []

    @patch("ms_teams_mcp.server.graph_get")
    def test_excludes_outside_window(self, mock_get):
        old = "2000-01-01T00:00:00Z"
        mock_get.side_effect = _channel_side_effect({
            "c1": [{"id": "m1", "createdDateTime": old,
                    "from": {"user": {"displayName": "Old"}},
                    "body": {"content": "stale"}, "mentions": []}],
            "c2": [],
        })
        activity, mentions, scanned, total = _scan_channels(24, 20, 10, "me")
        assert activity == []

    @patch("ms_teams_mcp.server.graph_get")
    def test_detects_my_mention(self, mock_get):
        recent = "2999-01-01T00:00:00Z"
        mock_get.side_effect = _channel_side_effect({
            "c1": [{"id": "m1", "createdDateTime": recent,
                    "from": {"user": {"displayName": "Bob"}},
                    "body": {"content": "ping"},
                    "mentions": [{"mentioned": {"user": {"id": "me"}}}]}],
            "c2": [],
        })
        activity, mentions, scanned, total = _scan_channels(24, 20, 10, "me")
        assert len(mentions) == 1 and "Bob" in mentions[0]

    @patch("ms_teams_mcp.server.graph_get")
    def test_respects_max_channels(self, mock_get):
        recent = "2999-01-01T00:00:00Z"
        mock_get.side_effect = _channel_side_effect({
            "c1": [{"id": "m1", "createdDateTime": recent,
                    "from": {"user": {"displayName": "A"}},
                    "body": {"content": "x"}, "mentions": []}],
            "c2": [{"id": "m2", "createdDateTime": recent,
                    "from": {"user": {"displayName": "B"}},
                    "body": {"content": "y"}, "mentions": []}],
        })
        activity, mentions, scanned, total = _scan_channels(24, 1, 10, "me")
        assert scanned == 1 and total == 2   # capped

    @patch("ms_teams_mcp.server.graph_get")
    def test_channel_error_isolated(self, mock_get):
        recent = "2999-01-01T00:00:00Z"
        def _se(path, params=None, url=None):
            if path == "/me/joinedTeams":
                return {"value": [{"id": "t1", "displayName": "Team1"}]}
            if path == "/teams/t1/channels":
                return {"value": [{"id": "c1", "displayName": "General"},
                                  {"id": "c2", "displayName": "Random"}]}
            if path == "/teams/t1/channels/c1/messages":
                raise Exception("403")
            if path == "/teams/t1/channels/c2/messages":
                return {"value": [{"id": "m2", "createdDateTime": recent,
                                   "from": {"user": {"displayName": "C"}},
                                   "body": {"content": "ok"}, "mentions": []}]}
            return {"value": []}
        mock_get.side_effect = _se
        activity, mentions, scanned, total = _scan_channels(24, 20, 10, "me")
        assert any("C" in a for a in activity)   # c1 failure didn't kill the scan
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestScanChannels -v`
Expected: FAIL — `ImportError: cannot import name '_scan_channels'`

- [ ] **Step 3: Implement the scan**

`_briefing_flagged_email` 뒤에 삽입:
```python
def _scan_channels(hours: int, max_channels: int, max_messages_per_channel: int, me_id):
    """Walk joined teams -> channels -> recent messages once.

    Returns (activity, mentions, scanned, total): recent posts within the time
    window, posts where me_id is @mentioned, channels actually scanned, and total
    channels seen. total > scanned means the max_channels cap was hit. Per-team and
    per-channel errors are isolated so one failure does not abort the scan.
    """
    activity, mentions = [], []
    cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)

    # Enumerate all channels first (one cheap call per team) to know the true total.
    all_channels = []  # (team_name, team_id, channel_name, channel_id)
    teams = graph_get("/me/joinedTeams", params={"$select": "id,displayName"}).get("value", [])
    for t in teams:
        tid, tname = t.get("id"), t.get("displayName", "")
        try:
            chs = graph_get(f"/teams/{tid}/channels", params={"$select": "id,displayName"}).get("value", [])
        except Exception:
            continue
        for ch in chs:
            all_channels.append((tname, tid, ch.get("displayName", ""), ch.get("id")))

    total = len(all_channels)
    scanned = 0
    for (tname, tid, cname, cid) in all_channels:
        if scanned >= max_channels:
            break
        scanned += 1
        try:
            msgs = graph_get(
                f"/teams/{tid}/channels/{cid}/messages",
                params={"$top": max_messages_per_channel},
            ).get("value", [])
        except Exception:
            continue
        for m in msgs:
            created_raw = m.get("createdDateTime", "")
            if not created_raw:
                continue
            try:
                created = datetime.fromisoformat(created_raw.replace("Z", "+00:00"))
            except ValueError:
                continue
            if created < cutoff:
                continue
            sender = (m.get("from") or {}).get("user") or {}
            sender_name = sender.get("displayName") or "Unknown"
            body = strip_html(m.get("body", {}).get("content", ""))[:100]
            line = f"[{tname} / {cname}] {sender_name}: {body}"
            activity.append(line)
            if me_id:
                for mention in (m.get("mentions") or []):
                    mentioned_id = ((mention.get("mentioned") or {}).get("user") or {}).get("id")
                    if mentioned_id == me_id:
                        mentions.append(line)
                        break
    return activity, mentions, scanned, total
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestScanChannels -v`
Expected: PASS (5 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(briefing): add _scan_channels (bounded one-pass channel scan)"
```

---

## Task 3: `daily_briefing` 도구 (종합 + 와이어링)

5개 섹션을 종합한다. 채널은 1회 스캔해 활동·멘션을 공유하고, `/me` 실패 시 멘션만 degrade한다.

**Files:**
- Modify: `ms_teams_mcp/server.py` (insert after `_scan_channels`)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `_briefing_calendar_today`, `_briefing_unread_email`, `_briefing_recent_chats`, `_briefing_flagged_email`, `_scan_channels`, `graph_get`
- Produces: `daily_briefing(hours: int = 24, max_channels: int = 20, max_messages_per_channel: int = 10) -> str` (`@mcp.tool()`)

- [ ] **Step 1: Write the failing tests**

```python
from ms_teams_mcp.server import daily_briefing

def _briefing_side_effect(path, params=None, url=None):
    if path == "/me":
        return {"id": "me"}
    if path == "/me/calendarView":
        return {"value": [{"subject": "Standup",
                           "start": {"dateTime": "2026-06-24T09:00:00"},
                           "location": {"displayName": ""}}]}
    if path == "/me/mailFolders/inbox":
        return {"unreadItemCount": 3, "totalItemCount": 100}
    if path == "/me/chats":
        return {"value": []}
    if path == "/me/messages":
        return {"value": [{"subject": "Flagged1",
                           "from": {"emailAddress": {"name": "Sue"}}}]}
    if path == "/me/joinedTeams":
        return {"value": [{"id": "t1", "displayName": "Team1"}]}
    if path == "/teams/t1/channels":
        return {"value": [{"id": "c1", "displayName": "General"}]}
    if path == "/teams/t1/channels/c1/messages":
        return {"value": [{"id": "m1", "createdDateTime": "2999-01-01T00:00:00Z",
                           "from": {"user": {"displayName": "Alice"}},
                           "body": {"content": "hi"},
                           "mentions": [{"mentioned": {"user": {"id": "me"}}}]}]}
    return {"value": []}

class TestDailyBriefing:
    @patch("ms_teams_mcp.server.graph_get")
    def test_all_sections_present(self, mock_get):
        mock_get.side_effect = _briefing_side_effect
        result = daily_briefing()
        for header in ["Daily Briefing", "Today's Schedule", "Unread Mail",
                       "Recent Chats", "Channel Activity", "Needs Your Response"]:
            assert header in result
        assert "Standup" in result
        assert "3 unread / 100 total" in result
        assert "Flagged emails (1):" in result
        assert "Channel @mentions (1):" in result   # Alice mention surfaced

    @patch("ms_teams_mcp.server.graph_get")
    def test_partial_failure_degrades(self, mock_get):
        def _se(path, params=None, url=None):
            if path == "/me/mailFolders/inbox":
                raise Exception("mail down")
            return _briefing_side_effect(path, params, url)
        mock_get.side_effect = _se
        result = daily_briefing()
        assert "Unread Mail" in result
        assert "Failed to retrieve" in result      # only that section degraded
        assert "Today's Schedule" in result and "Standup" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_me_failure_degrades_mentions(self, mock_get):
        def _se(path, params=None, url=None):
            if path == "/me":
                raise Exception("no me")
            return _briefing_side_effect(path, params, url)
        mock_get.side_effect = _se
        result = daily_briefing()
        assert "mention detection unavailable" in result
        assert "Channel Activity" in result        # activity still works

    @patch("ms_teams_mcp.server.graph_get")
    def test_capped_note(self, mock_get):
        def _se(path, params=None, url=None):
            if path == "/teams/t1/channels":
                return {"value": [{"id": "c1", "displayName": "General"},
                                  {"id": "c2", "displayName": "Random"}]}
            return _briefing_side_effect(path, params, url)
        mock_get.side_effect = _se
        result = daily_briefing(max_channels=1)
        assert "capped" in result
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestDailyBriefing -v`
Expected: FAIL — `ImportError: cannot import name 'daily_briefing'`

- [ ] **Step 3: Implement the tool**

`_scan_channels` 뒤에 삽입:
```python
@mcp.tool()
def daily_briefing(hours: int = 24, max_channels: int = 20, max_messages_per_channel: int = 10) -> str:
    """Aggregate today's schedule, unread mail, recent chats, channel activity,
    and items needing your response into a single read-only briefing.
    - hours: recency window for channel posts and @mentions (default 24)
    - max_channels: cap on channels scanned, to bound API calls (default 20)
    - max_messages_per_channel: messages fetched per channel (default 10)
    """
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    parts = [f"Daily Briefing — {today} (last {hours}h)"]
    parts.append(_briefing_calendar_today())
    parts.append(_briefing_unread_email())
    parts.append(_briefing_recent_chats())

    # /me id for @mention detection; degrade gracefully if unavailable.
    try:
        me_id = graph_get("/me", params={"$select": "id"}).get("id")
    except Exception:
        me_id = None

    activity_header = "── Channel Activity ──"
    response_header = "── Needs Your Response ──"
    mentions = None
    try:
        activity, mentions, scanned, total = _scan_channels(
            hours, max_channels, max_messages_per_channel, me_id
        )
        act_lines = [activity_header]
        if activity:
            act_lines += [f"{i}. {a}" for i, a in enumerate(activity, 1)]
        else:
            act_lines.append(f"No channel activity in the last {hours}h.")
        if total > scanned:
            act_lines.append(f"   (scanned {scanned}/{total} channels, capped)")
        parts.append("\n".join(act_lines))
    except Exception as e:
        parts.append(f"{activity_header}\nFailed to retrieve ({e})")

    resp_lines = [response_header, _briefing_flagged_email()]
    if mentions is None:
        resp_lines.append("Channel @mentions: (unavailable)")
    elif me_id is None:
        resp_lines.append("Channel @mentions: (mention detection unavailable)")
    elif mentions:
        resp_lines.append(f"Channel @mentions ({len(mentions)}):")
        resp_lines += [f"  - {m}" for m in mentions]
    else:
        resp_lines.append("Channel @mentions: none")
    parts.append("\n".join(resp_lines))

    return "\n\n".join(parts)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestDailyBriefing -v`
Expected: PASS (4 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(briefing): add daily_briefing tool wiring all sections"
```

---

## Task 4: 문서 갱신

**Files:**
- Modify: `ms-teams-mcp/CLAUDE.md`, `ms-teams-mcp/README.md`, `ms-teams-mcp/TODO.md`

- [ ] **Step 1: CLAUDE.md**

- "MCP Tools (42)" → "MCP Tools (43)".
- Architecture/Conventions에 한 줄 추가: `daily_briefing` — read-only aggregation of schedule/unread mail/recent chats/channel activity/needs-response, channels scanned once via `_scan_channels`, bounded by `hours`/`max_channels`/`max_messages_per_channel`.

- [ ] **Step 2: README.md**

- 도구 목록/표에 `daily_briefing` 한 줄 추가(예: "Aggregate today's schedule, unread mail, recent chats, channel activity, and items needing response"). 도구 수 표기가 있으면 +1.

- [ ] **Step 3: TODO.md**

- "기능 추가" 섹션에 완료 항목으로 `- [x] daily_briefing — 일일 업무 종합 브리핑 (읽기 전용)` 추가.

- [ ] **Step 4: Verify no non-ASCII leaked into new server.py code**

Run:
```bash
cd ms-teams-mcp && .venv/bin/python -c "import ms_teams_mcp.server as s, inspect; [print('CHECK', n) for n in ('daily_briefing','_scan_channels','_briefing_calendar_today','_briefing_unread_email','_briefing_recent_chats','_briefing_flagged_email') if any(ord(ch)>0x3000 for ch in inspect.getsource(getattr(s,n)))]"
```
Expected: 출력 없음.

> 주의: 위 검사는 `ord > 0x3000`(한글/CJK)만 잡는다. 신규 코드의 섹션 헤더는 박스드로잉 문자 `──`(U+2500, 0x2500 < 0x3000)를 의도적으로 포함하므로 이 검사에 걸리지 않는다 — 정상이다. (ASCII 전용이 아니라 "한국어/CJK 없음"을 검증하는 것이며, 출력 문자열의 영어 규칙은 충족한다.)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add CLAUDE.md README.md TODO.md
git commit -m "docs(briefing): document daily_briefing tool"
```

---

## Task 5: 통합 검증

**Files:** 없음 (검증만)

- [ ] **Step 1: Run all new briefing test classes**

Run:
```bash
cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py -v -k "BriefingSimpleSections or ScanChannels or DailyBriefing"
```
Expected: 모두 PASS (16 passed).

- [ ] **Step 2: Full-suite regression check**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/ -q`
Expected: 신규 16개 추가 통과 → `13 failed, 42 passed` (13개는 기존 stale, 신규 회귀 없음).

- [ ] **Step 3: Sanity import**

Run: `cd ms-teams-mcp && .venv/bin/python -c "import ms_teams_mcp.server; print('import OK')"`
Expected: `import OK`

---

## 검증 요약 (구현 후 기대 상태)

- 신규 private 헬퍼 5개 + `@mcp.tool() daily_briefing` 1개, 모두 영어 출력.
- 신규 pytest 클래스 3개(16 테스트) 전부 통과, 기존 stale 13개 불변.
- 신규 스코프 없음. 도구 수 42→43.
- CLAUDE.md/README/TODO 갱신.

## Self-Review note (박스드로잉 문자)

섹션 헤더에 `──`(U+2500)를 쓴다. 이는 한국어/CJK가 아니라 장식 구분선이며 `get_unread_summary` 등 기존 출력 스타일과 일관된 시각적 구분 목적이다. CLAUDE.md의 "English only" 규칙은 자연어를 영어로 쓰라는 의미이므로 위배가 아니다. (만약 리뷰에서 ASCII 전용을 요구하면 `---`로 대체 가능.)
