# 메일 관리 4종 도구 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** ms-teams-mcp 서버에 메일 상태 변경 도구 4종(`mark_email_read`, `flag_email`, `move_email`, `delete_email`)을 추가한다.

**Architecture:** 단일 파일 MCP 서버(`ms_teams_mcp/server.py`)에 "동작당 `@mcp.tool()` 1개" 컨벤션을 따라 도구 4개를 추가하고, 쉼표 구분 일괄 처리 반복 로직은 공유 헬퍼 `_apply_to_messages`로, 폴더명 해석은 `_resolve_folder_id`로 모은다. 모든 Graph 호출은 기존 `graph_patch`/`graph_post`/`graph_delete` 헬퍼를 거친다.

**Tech Stack:** Python 3, FastMCP, Microsoft Graph API, pytest + `unittest.mock`.

## Global Constraints

- **영어 전용 소스**: 모든 주석·docstring·사용자 노출 출력 문자열은 영어로만 작성 (`ms-teams-mcp/CLAUDE.md` 규칙).
- **모든 Graph 호출은 `graph_*` 헬퍼 경유** (`graph_get`/`graph_post`/`graph_patch`/`graph_delete`/`graph_post_action`).
- **확인 정책**: `delete_email`만 호출 전 명시적 사용자 확인 필수 — docstring에 `IMPORTANT: Always show the target emails to the user and get explicit confirmation before calling this tool.` 명시. 나머지 3종은 즉시 실행.
- **일괄 처리**: 4종 모두 `message_ids`를 쉼표 구분 문자열로 받아 1개 이상 처리, 건별 성공/실패 요약 반환.
- **삭제 의미**: Graph `DELETE`의 기본 동작(휴지통 이동, 복구 가능). 영구삭제 아님.
- **테스트**: `tests/test_server.py`의 기존 mock 패턴을 따라 TDD. 신규 테스트만 추가하며 기존 stale 테스트(한국어 assert)는 손대지 않는다.
- **테스트 출력 assert는 영어 기준** (신규 코드가 반환하는 문자열).

---

## File Structure

- **Modify** `ms_teams_mcp/server.py`:
  - `SCOPES` 리스트 (line 28-38)에 `"Mail.ReadWrite"` 추가
  - 공유 헬퍼 `_apply_to_messages` 추가 — `_parse_recipients`(line 200) 근처
  - `WELL_KNOWN_FOLDERS` 상수 + `_resolve_folder_id` 헬퍼 추가 — `_apply_to_messages` 뒤
  - 도구 4개 추가 — `forward_email`/`list_mail_folders`(line 721 근처) 뒤, 이메일 도구 영역
- **Modify** `tests/test_server.py`: 헬퍼·도구 테스트 클래스 추가
- **Modify** `ms-teams-mcp/CLAUDE.md`: 도구 수 38→42, 확인 필수 목록에 `delete_email`, 신규 스코프
- **Modify** `ms-teams-mcp/README.md`: 도구 표·스코프에 4종 반영
- **Modify** `ms-teams-mcp/TODO.md`: 메일 관리 항목 체크

---

## Task 0: 개발 환경 셋업 (완료됨 — 참고)

> **이미 완료:** PEP 668(externally-managed) 때문에 시스템 pip 직접 설치가 막혀 있어,
> `ms-teams-mcp/.venv` 가상환경을 만들고 `pip install -e . pytest`를 설치해 두었다(`.venv`는
> gitignore됨). **이후 모든 테스트는 `.venv/bin/python -m pytest ...` 로 실행한다**
> (`.venv/bin/python -m pytest`가 아님 — 시스템 파이썬엔 의존성이 없다).
>
> 베이스라인 확인 결과: `13 failed, 10 passed`. 13개 실패는 전부 **기존 stale 테스트**(한국어
> assert가 영어 출력과 불일치)로 본 작업 범위 밖이다. 신규 추가 테스트는 모두 통과해야 한다.

작업 디렉토리는 `ms-teams-mcp/`. 재현이 필요할 때만:
```bash
cd ms-teams-mcp
python3 -m venv .venv
.venv/bin/python -m pip install -e . pytest
.venv/bin/python -m pytest tests/ -q   # baseline: 13 failed, 10 passed
```

---

## Task 1: `_apply_to_messages` 공유 배치 헬퍼

쉼표 구분 ID를 각각 `action`에 적용하고 건별 성공/실패를 집계한다.

**Files:**
- Modify: `ms_teams_mcp/server.py` (insert after `_parse_recipients`, line 202)
- Test: `tests/test_server.py`

**Interfaces:**
- Produces: `_apply_to_messages(message_ids: str, action) -> str` — `action`은 단일 message id를 받아 Graph 호출 1회를 수행하고 실패 시 raise하는 callable. 반환은 `"N succeeded, M failed."` 요약(실패 시 상세 줄 추가).

- [ ] **Step 1: Write the failing tests**

`tests/test_server.py` import 블록에 추가: `_apply_to_messages` (그리고 이후 태스크에서 쓸 `_resolve_folder_id`, `mark_email_read`, `flag_email`, `move_email`, `delete_email`는 해당 태스크에서 추가).

```python
from ms_teams_mcp.server import _apply_to_messages

class TestApplyToMessages:
    def test_all_succeed(self):
        calls = []
        result = _apply_to_messages("a, b, c", lambda mid: calls.append(mid))
        assert calls == ["a", "b", "c"]
        assert "3 succeeded, 0 failed." in result

    def test_partial_failure(self):
        def action(mid):
            if mid == "bad":
                raise Exception("Resource not found (404)")
        result = _apply_to_messages("ok,bad", action)
        assert "1 succeeded, 1 failed." in result
        assert "bad: Resource not found (404)" in result

    def test_empty_input(self):
        result = _apply_to_messages("  ", lambda mid: None)
        assert result == "No message IDs provided."
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestApplyToMessages -v`
Expected: FAIL — `ImportError: cannot import name '_apply_to_messages'`

- [ ] **Step 3: Implement the helper**

`ms_teams_mcp/server.py`에서 `_parse_recipients` 함수 정의 바로 뒤(line 202 다음, `_parse_attendees` 앞)에 삽입:

```python
def _apply_to_messages(message_ids: str, action) -> str:
    """Apply `action(message_id)` to each comma-separated message ID and aggregate results.

    `action` performs one Graph call for a single message and may raise on failure.
    Returns a summary like "N succeeded, M failed." with failure details when M > 0.
    """
    ids = [m.strip() for m in message_ids.split(",") if m.strip()]
    if not ids:
        return "No message IDs provided."
    succeeded, failed = [], []
    for mid in ids:
        try:
            action(mid)
            succeeded.append(mid)
        except Exception as e:
            failed.append(f"  {mid}: {e}")
    parts = [f"{len(succeeded)} succeeded, {len(failed)} failed."]
    if failed:
        parts.append("\n".join(failed))
    return "\n".join(parts)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestApplyToMessages -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(mail): add _apply_to_messages batch helper"
```

---

## Task 2: `_resolve_folder_id` 폴더명 해석 헬퍼

목적지 폴더명을 Graph 폴더 id로 해석한다.

**Files:**
- Modify: `ms_teams_mcp/server.py` (insert after `_apply_to_messages`)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `graph_get` (기존)
- Produces:
  - `WELL_KNOWN_FOLDERS: set[str]`
  - `_resolve_folder_id(destination: str) -> str` — well-known명(소문자)은 그대로 반환(API 호출 없음), 그 외엔 `/me/mailFolders`의 `displayName` 대소문자 무시 매칭. 0건/2건+이면 raise.

- [ ] **Step 1: Write the failing tests**

```python
from ms_teams_mcp.server import _resolve_folder_id

class TestResolveFolderId:
    def test_well_known_passthrough(self):
        # No API call needed for well-known names
        assert _resolve_folder_id("Archive") == "archive"
        assert _resolve_folder_id("inbox") == "inbox"

    @patch("ms_teams_mcp.server.graph_get")
    def test_display_name_match(self, mock_graph_get):
        mock_graph_get.return_value = {"value": [
            {"id": "AAA", "displayName": "Projects"},
            {"id": "BBB", "displayName": "Receipts"},
        ]}
        assert _resolve_folder_id("projects") == "AAA"

    @patch("ms_teams_mcp.server.graph_get")
    def test_not_found_raises(self, mock_graph_get):
        mock_graph_get.return_value = {"value": [{"id": "AAA", "displayName": "Projects"}]}
        with pytest.raises(Exception, match="not found"):
            _resolve_folder_id("Nope")

    @patch("ms_teams_mcp.server.graph_get")
    def test_ambiguous_raises(self, mock_graph_get):
        mock_graph_get.return_value = {"value": [
            {"id": "AAA", "displayName": "Work"},
            {"id": "BBB", "displayName": "work"},
        ]}
        with pytest.raises(Exception, match="ambiguous"):
            _resolve_folder_id("work")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestResolveFolderId -v`
Expected: FAIL — `ImportError: cannot import name '_resolve_folder_id'`

- [ ] **Step 3: Implement the helper**

`_apply_to_messages` 뒤에 삽입:

```python
WELL_KNOWN_FOLDERS = {
    "inbox", "archive", "drafts", "sentitems",
    "deleteditems", "junkemail", "outbox",
}

def _resolve_folder_id(destination: str) -> str:
    """Resolve a destination folder name to a Graph folder id.

    Well-known names (case-insensitive) pass through unchanged with no API call.
    Otherwise look up by displayName via /me/mailFolders; raise on no match
    (lists available folders) or ambiguous match (suggest using the folder ID).
    """
    key = destination.strip().lower()
    if key in WELL_KNOWN_FOLDERS:
        return key
    data = graph_get("/me/mailFolders", params={"$select": "id,displayName"})
    folders = data.get("value", [])
    matches = [f for f in folders if f.get("displayName", "").lower() == key]
    if not matches:
        names = ", ".join(f.get("displayName", "") for f in folders)
        raise Exception(f"Folder '{destination}' not found. Available: {names}")
    if len(matches) > 1:
        raise Exception(
            f"Folder '{destination}' is ambiguous ({len(matches)} matches). "
            f"Use the folder ID instead."
        )
    return matches[0]["id"]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestResolveFolderId -v`
Expected: PASS (4 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(mail): add _resolve_folder_id folder-name resolver"
```

---

## Task 3: `mark_email_read` 도구

**Files:**
- Modify: `ms_teams_mcp/server.py` (add after `list_mail_folders`, ~line 730)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `_apply_to_messages`, `graph_patch`
- Produces: `mark_email_read(message_ids: str, is_read: bool = True) -> str`

- [ ] **Step 1: Write the failing tests**

```python
from ms_teams_mcp.server import mark_email_read

class TestMarkEmailRead:
    @patch("ms_teams_mcp.server.graph_patch")
    def test_mark_read(self, mock_patch):
        result = mark_email_read("m1,m2")
        assert "2 succeeded, 0 failed." in result
        assert mock_patch.call_count == 2
        mock_patch.assert_any_call("/me/messages/m1", {"isRead": True})

    @patch("ms_teams_mcp.server.graph_patch")
    def test_mark_unread(self, mock_patch):
        mark_email_read("m1", is_read=False)
        mock_patch.assert_called_once_with("/me/messages/m1", {"isRead": False})
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestMarkEmailRead -v`
Expected: FAIL — `ImportError: cannot import name 'mark_email_read'`

- [ ] **Step 3: Implement the tool**

`list_mail_folders` 함수 정의가 끝나는 지점 뒤에 삽입:

```python
@mcp.tool()
def mark_email_read(message_ids: str, is_read: bool = True) -> str:
    """
    Mark one or more emails as read or unread.
    - message_ids: One or more message IDs, comma-separated
    - is_read: True to mark as read (default), False to mark as unread
    """
    return _apply_to_messages(
        message_ids,
        lambda mid: graph_patch(f"/me/messages/{mid}", {"isRead": is_read}),
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestMarkEmailRead -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(mail): add mark_email_read tool"
```

---

## Task 4: `flag_email` 도구

**Files:**
- Modify: `ms_teams_mcp/server.py` (add after `mark_email_read`)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `_apply_to_messages`, `graph_patch`
- Produces: `flag_email(message_ids: str, flag_status: str = "flagged") -> str`

- [ ] **Step 1: Write the failing tests**

```python
from ms_teams_mcp.server import flag_email

class TestFlagEmail:
    @patch("ms_teams_mcp.server.graph_patch")
    def test_flag_default(self, mock_patch):
        flag_email("m1")
        mock_patch.assert_called_once_with(
            "/me/messages/m1", {"flag": {"flagStatus": "flagged"}}
        )

    @patch("ms_teams_mcp.server.graph_patch")
    def test_flag_complete(self, mock_patch):
        flag_email("m1", flag_status="complete")
        mock_patch.assert_called_once_with(
            "/me/messages/m1", {"flag": {"flagStatus": "complete"}}
        )

    @patch("ms_teams_mcp.server.graph_patch")
    def test_invalid_status_no_call(self, mock_patch):
        result = flag_email("m1", flag_status="bogus")
        assert "Invalid flag_status" in result
        mock_patch.assert_not_called()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestFlagEmail -v`
Expected: FAIL — `ImportError: cannot import name 'flag_email'`

- [ ] **Step 3: Implement the tool**

`mark_email_read` 뒤에 삽입:

```python
@mcp.tool()
def flag_email(message_ids: str, flag_status: str = "flagged") -> str:
    """
    Flag, complete, or clear the flag on one or more emails.
    - message_ids: One or more message IDs, comma-separated
    - flag_status: "flagged" (default), "complete", or "notFlagged" (clear)
    """
    valid = {"flagged", "complete", "notFlagged"}
    if flag_status not in valid:
        return f"Invalid flag_status '{flag_status}'. Use one of: flagged, complete, notFlagged."
    return _apply_to_messages(
        message_ids,
        lambda mid: graph_patch(f"/me/messages/{mid}", {"flag": {"flagStatus": flag_status}}),
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestFlagEmail -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(mail): add flag_email tool"
```

---

## Task 5: `move_email` 도구

**Files:**
- Modify: `ms_teams_mcp/server.py` (add after `flag_email`)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `_resolve_folder_id`, `_apply_to_messages`, `graph_post`
- Produces: `move_email(message_ids: str, destination: str) -> str`

- [ ] **Step 1: Write the failing tests**

```python
from ms_teams_mcp.server import move_email

class TestMoveEmail:
    @patch("ms_teams_mcp.server.graph_post")
    def test_move_well_known(self, mock_post):
        # 'archive' is well-known -> no folder lookup, destinationId == "archive"
        result = move_email("m1,m2", destination="archive")
        assert "2 succeeded, 0 failed." in result
        mock_post.assert_any_call("/me/messages/m1/move", {"destinationId": "archive"})

    @patch("ms_teams_mcp.server.graph_post")
    @patch("ms_teams_mcp.server.graph_get")
    def test_move_display_name(self, mock_get, mock_post):
        mock_get.return_value = {"value": [{"id": "FID", "displayName": "Projects"}]}
        move_email("m1", destination="Projects")
        # folder resolved exactly once, then used for the move
        mock_post.assert_called_once_with("/me/messages/m1/move", {"destinationId": "FID"})
        assert mock_get.call_count == 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestMoveEmail -v`
Expected: FAIL — `ImportError: cannot import name 'move_email'`

- [ ] **Step 3: Implement the tool**

`flag_email` 뒤에 삽입. 목적지는 루프 진입 전 1회만 해석한다:

```python
@mcp.tool()
def move_email(message_ids: str, destination: str) -> str:
    """
    Move one or more emails to another mail folder.
    - message_ids: One or more message IDs, comma-separated
    - destination: A well-known folder name (inbox, archive, deleteditems, ...),
      a custom folder display name, or a folder ID
    """
    destination_id = _resolve_folder_id(destination)
    return _apply_to_messages(
        message_ids,
        lambda mid: graph_post(f"/me/messages/{mid}/move", {"destinationId": destination_id}),
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestMoveEmail -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(mail): add move_email tool with folder-name resolution"
```

---

## Task 6: `delete_email` 도구 (확인 필수)

**Files:**
- Modify: `ms_teams_mcp/server.py` (add after `move_email`)
- Test: `tests/test_server.py`

**Interfaces:**
- Consumes: `_apply_to_messages`, `graph_delete`
- Produces: `delete_email(message_ids: str) -> str`

- [ ] **Step 1: Write the failing tests**

```python
from ms_teams_mcp.server import delete_email

class TestDeleteEmail:
    @patch("ms_teams_mcp.server.graph_delete")
    def test_delete_multiple(self, mock_delete):
        result = delete_email("m1,m2,m3")
        assert "3 succeeded, 0 failed." in result
        assert mock_delete.call_count == 3
        mock_delete.assert_any_call("/me/messages/m1")

    @patch("ms_teams_mcp.server.graph_delete")
    def test_delete_confirmation_note_in_docstring(self, mock_delete):
        # delete is destructive -> docstring must instruct explicit confirmation
        assert "confirmation" in delete_email.__doc__.lower()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestDeleteEmail -v`
Expected: FAIL — `ImportError: cannot import name 'delete_email'`

- [ ] **Step 3: Implement the tool**

`move_email` 뒤에 삽입:

```python
@mcp.tool()
def delete_email(message_ids: str) -> str:
    """
    Delete one or more emails (moved to Deleted Items; recoverable).
    IMPORTANT: Always show the target emails to the user and get explicit confirmation before calling this tool.
    - message_ids: One or more message IDs, comma-separated
    """
    return _apply_to_messages(
        message_ids,
        lambda mid: graph_delete(f"/me/messages/{mid}"),
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py::TestDeleteEmail -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py tests/test_server.py
git commit -m "feat(mail): add delete_email tool (confirmation required)"
```

---

## Task 7: `Mail.ReadWrite` 스코프 추가

쓰기 작업에 필요한 권한을 추가한다.

**Files:**
- Modify: `ms_teams_mcp/server.py:29` (SCOPES 리스트)

- [ ] **Step 1: Add the scope**

`ms_teams_mcp/server.py`의 SCOPES 첫 줄을 수정:

```python
    "Mail.Read", "Mail.Send", "Mail.ReadWrite", "User.Read",
```

(line 29: `"Mail.Read", "Mail.Send", "User.Read",` → 위와 같이 `"Mail.ReadWrite"` 추가)

- [ ] **Step 2: Verify import still succeeds**

Run: `cd ms-teams-mcp && .venv/bin/python -c "import ms_teams_mcp.server as s; assert 'Mail.ReadWrite' in s.SCOPES; print('scope OK')"`
Expected: `scope OK`

- [ ] **Step 3: Commit**

```bash
cd ms-teams-mcp
git add ms_teams_mcp/server.py
git commit -m "feat(mail): add Mail.ReadWrite scope for mail management"
```

---

## Task 8: 문서 갱신 (CLAUDE.md / README.md / TODO.md)

**Files:**
- Modify: `ms-teams-mcp/CLAUDE.md`
- Modify: `ms-teams-mcp/README.md`
- Modify: `ms-teams-mcp/TODO.md`

- [ ] **Step 1: CLAUDE.md 갱신**

- "MCP Tools (38)" → "MCP Tools (42)"로 도구 수 갱신.
- "User confirmation before sending" 도구 목록에 `delete_email` 추가.
- Scopes 줄에 `Mail.ReadWrite` 추가.
- Conventions/Architecture에서 메일 도구 설명이 있으면 4종(읽음/플래그/이동/삭제) 한 줄 요약 추가.

- [ ] **Step 2: README.md 갱신**

- 도구 목록/표에 `mark_email_read`, `flag_email`, `move_email`, `delete_email` 추가(각 한 줄 설명, 영어/한국어는 README 기존 언어 관례를 따름).
- 권한(Scopes) 목록에 `Mail.ReadWrite` 추가.

- [ ] **Step 3: TODO.md 갱신**

- 메일 관리 관련 항목이 있으면 `[x]`로 체크. 없으면 "기능 추가" 섹션에 완료 항목으로 4종을 `[x]`로 기록.

- [ ] **Step 4: Verify no Korean leaked into server.py output strings**

Run: `cd ms-teams-mcp && .venv/bin/python -c "import ms_teams_mcp.server as s, inspect; [print('CHECK', n) for n in ('mark_email_read','flag_email','move_email','delete_email','_apply_to_messages','_resolve_folder_id') if any(ord(ch)>0x3000 for ch in inspect.getsource(getattr(s,n)))]"`
Expected: 출력 없음 (신규 코드에 비-ASCII/한국어 문자열 없음)

- [ ] **Step 5: Commit**

```bash
cd ms-teams-mcp
git add CLAUDE.md README.md TODO.md
git commit -m "docs(mail): document mail management tools and Mail.ReadWrite scope"
```

---

## Task 9: 전체 신규 테스트 통합 검증

**Files:** 없음 (검증만)

- [ ] **Step 1: Run all new test classes together**

Run:
```bash
cd ms-teams-mcp && .venv/bin/python -m pytest tests/test_server.py -v -k "ApplyToMessages or ResolveFolderId or MarkEmailRead or FlagEmail or MoveEmail or DeleteEmail"
```
Expected: 모두 PASS (16 passed). 기존 stale 테스트는 이 `-k` 필터에서 제외되므로 영향 없음.

- [ ] **Step 2: Sanity import of the whole server**

Run: `cd ms-teams-mcp && .venv/bin/python -c "import ms_teams_mcp.server; print('import OK')"`
Expected: `import OK`

---

## 검증 요약 (구현 후 기대 상태)

- 신규 헬퍼 2개 + 도구 4개, 모두 영어 출력.
- 신규 pytest 클래스 6개(16 테스트) 전부 통과.
- `Mail.ReadWrite` 스코프 추가됨(재인증 필요 — 사용자 안내).
- CLAUDE.md/README/TODO 갱신.
- 기존 stale 테스트는 손대지 않음(별도 정리 항목).
