# 메일 관리 4종 도구 설계

- **날짜**: 2026-06-24
- **대상 프로젝트**: `ms-teams-mcp`
- **파일**: `ms_teams_mcp/server.py` (단일 파일 MCP 서버)

## 목표

읽기·발송·답장·전달은 있으나 **메일 상태를 변경하는 도구가 없는** 빈틈을 메운다.
다음 4개 MCP 도구를 추가한다.

1. `mark_email_read` — 읽음/안읽음 표시
2. `flag_email` — 플래그 설정/완료/해제
3. `move_email` — 다른 폴더로 이동
4. `delete_email` — 휴지통으로 삭제(soft-delete, 복구 가능)

## 설계 결정 (사용자 확정)

| 항목 | 결정 |
|------|------|
| 확인 정책 | **`delete_email`만** 명시적 사용자 확인 필수. 나머지 3종은 가역적·저위험이라 즉시 실행 |
| 일괄 처리 | 4종 모두 `message_ids`를 **쉼표 구분 문자열**로 받아 1개 이상 처리, 건별 성공/실패 요약 반환 |
| 이동 목적지 | well-known 폴더명은 그대로 통과, 그 외엔 표시이름→ID **자동 해석**(모호/없음 시 친화적 오류) |
| 삭제 의미 | Graph `DELETE`의 기본 동작인 **휴지통 이동**(복구 가능). 영구삭제 아님 |
| 테스트 | mock 기반 pytest 인프라가 `tests/test_server.py`에 이미 존재함을 확인 → 기존 패턴(graph 헬퍼 patch)을 따라 4종 도구·헬퍼의 단위 테스트를 TDD로 작성. (신규 인프라 도입 아님) |

## 아키텍처

기존 컨벤션("동작당 `@mcp.tool()` 1개", "모든 Graph 호출은 `graph_*` 헬퍼 경유")을 따른다.
일괄 처리 반복 로직은 공유 헬퍼로 모아 4중 복붙을 피한다.

### 공유 헬퍼 (내부 함수, `@mcp.tool()` 아님)

`_parse_recipients` 등 다른 공유 헬퍼 옆(서버 파일의 Shared Helpers 영역)에 배치한다.

```python
def _apply_to_messages(message_ids: str, action) -> str:
    """Apply `action(message_id)` to each comma-separated message ID and aggregate results.

    `action` performs one Graph call for a single message and may raise on failure.
    Returns a summary string: "N succeeded, M failed" with failure details when M > 0.
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

각 도구는 "한 메일에 대한 동작"만 정의하고 `_apply_to_messages`에 위임한다.

### 폴더 해석 헬퍼 (`move_email` 전용)

```python
WELL_KNOWN_FOLDERS = {
    "inbox", "archive", "drafts", "sentitems",
    "deleteditems", "junkemail", "outbox",
}

def _resolve_folder_id(destination: str) -> str:
    """Resolve a destination folder name to a Graph folder id.

    - Well-known names (case-insensitive) pass through unchanged (no API call).
    - Otherwise look up by displayName via /me/mailFolders.
      Raises on no match (lists available) or ambiguous match (suggests using ID).
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

배치 이동 시 목적지는 모든 메일에 동일하므로 **루프 진입 전 1회만 해석**한다.

## 도구 명세

모든 도구는 `/me/messages/{id}` 경로를 사용하며 기존 `graph_*` 헬퍼를 거친다.

### `mark_email_read(message_ids: str, is_read: bool = True) -> str`
- 동작: `graph_patch(f"/me/messages/{mid}", {"isRead": is_read})`
- `is_read=False`로 안읽음 처리도 가능(토글 겸용)
- 확인: 즉시 실행

### `flag_email(message_ids: str, flag_status: str = "flagged") -> str`
- 동작: `graph_patch(f"/me/messages/{mid}", {"flag": {"flagStatus": flag_status}})`
- `flag_status` 허용값: `"flagged"`(기본) / `"complete"` / `"notFlagged"`(해제)
- 호출 전 `flag_status` 값 검증 → 잘못된 값이면 친화적 오류 즉시 반환(Graph 호출 안 함)
- 확인: 즉시 실행

### `move_email(message_ids: str, destination: str) -> str`
- 목적지를 `_resolve_folder_id`로 1회 해석 → `destination_id`
- 동작: `graph_post(f"/me/messages/{mid}/move", {"destinationId": destination_id})`
- 확인: 즉시 실행

### `delete_email(message_ids: str) -> str`
- 동작: `graph_delete(f"/me/messages/{mid}")` → 휴지통(Deleted Items)으로 이동, 복구 가능
- **확인: 명시적 사용자 확인 필수.** docstring에 다음 명시:
  `"IMPORTANT: Always show the target emails to the user and get explicit confirmation before calling this tool."`

## 데이터 흐름

```
도구 호출
  → (move만) _resolve_folder_id(destination)  # 1회
  → _apply_to_messages(message_ids, action)
       → 각 id마다 action(id) = graph_patch/post/delete
       → _check_response 가 401/403/404/429 친화적 변환 (기존)
       → 건별 try/except 로 성공/실패 수집
  → "N succeeded, M failed" 요약 반환
```

## 에러 처리

- 개별 Graph 오류는 기존 `_check_response`가 이미 친화적 메시지(401/403/404/429)로 변환.
- `_apply_to_messages`가 건별로 잡아 부분 성공/실패를 함께 보고:
  - 전부 성공: `"4 succeeded, 0 failed."`
  - 부분 실패: `"3 succeeded, 1 failed.\n  <id>: Resource not found (404): ..."`
- 빈 입력: `"No message IDs provided."`
- 잘못된 `flag_status`: Graph 호출 전 검증 오류.

## 권한 (스코프)

- `SCOPES` 리스트에 **`Mail.ReadWrite` 추가**.
- 재인증 필요 → README / CLAUDE.md에 안내.

## 문서 갱신

- `CLAUDE.md`: 도구 수 38→42, "User confirmation before sending" 목록에 `delete_email` 추가, 신규 스코프 반영.
- `README.md`: 도구 표·스코프 목록에 4종 추가.
- `TODO.md`: 메일 관리 관련 항목 체크 표시.

## 테스트 — pytest (기존 mock 패턴)

`tests/test_server.py`에 mock 기반 pytest 스위트가 이미 존재한다. `graph_get`/`graph_post`/
`graph_post_action`을 `unittest.mock.patch`로 가로채는 동일 패턴을 따라, 신규 4종 도구와
헬퍼(`_apply_to_messages`, `_resolve_folder_id`)의 단위 테스트를 추가한다. (assert는 신규
코드가 반환하는 **영어** 출력 기준.)

검증 대상:
- `_apply_to_messages`: 단건 성공, 다건 성공, 일부 실패(부분 요약), 빈 입력
- `_resolve_folder_id`: well-known 통과(API 호출 없음), 표시이름 1건 매칭, 0건(없음 오류), 2건+(모호 오류)
- `mark_email_read`: 읽음(`isRead: true`) / 안읽음(`is_read=False`) 시 graph_patch 바디 검증
- `flag_email`: `flagged`/`complete`/`notFlagged` 바디 검증, 잘못된 status 시 Graph 호출 없이 오류
- `move_email`: 해석된 destinationId로 graph_post 호출 검증
- `delete_email`: graph_delete 호출 검증, 다건 처리

> **주의 (범위 밖):** 기존 테스트들은 한국어 출력 문자열을 검증하나 현 코드는 영어를 반환해
> 이미 실패(stale) 상태다. 이번 작업에서는 **신규 테스트만** 추가하고 기존 stale 테스트는
> 손대지 않는다(별도 정리 항목).

## 범위 밖 (이번에 하지 않음)

- 영구 삭제(permanentDelete) 지원
- pytest 테스트 인프라 신규 구축
- 폴더 생성/이름변경 등 폴더 관리 도구
