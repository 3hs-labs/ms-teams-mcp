# TODO

## 즉시 필요

- [x] CLAUDE.md 업데이트 — 페이지네이션 관련 Conventions 반영 (skip 파라미터, top 최대값 변경)

## 기능 추가 (우선순위 높음)

- [x] 메일 발송 기능 (`send_email`) — `Mail.Send` 권한 추가, 메일 작성/발송 도구
- [x] 메일 답장/전달 (`reply_email`, `forward_email`)
- [x] 채팅 생성 (`create_chat`) — 새 1:1/그룹 채팅 시작
- [x] 캘린더 연동 (`list_calendar_events`, `create_calendar_event`) — `Calendars.ReadWrite` 권한
- [x] 채널/채팅 메시지 답장 (`reply_to_channel_message`, `reply_to_chat_message`)
- [x] 사용자 검색 (`search_users`) — `People.Read` 권한
- [x] 읽지 않은 메일/채팅 요약 (`get_unread_summary`)
- [x] `create_reminder` — 특정 시간에 알림 설정 (캘린더 이벤트 기반)
- [x] `create_recurring_event` — 반복 캘린더 일정 생성 (매일/매주/매월)
- [x] `update_calendar_event` — 기존 일정 수정 (시간, 참석자, 장소 변경)
- [x] `delete_calendar_event` — 일정 삭제
- [x] Mail management tools (`mark_email_read`, `flag_email`, `move_email`, `delete_email`) — email read status, flags, folder moves, and deletion
- [x] daily_briefing — 일일 업무 종합 브리핑 (읽기 전용)

## 품질 개선 (우선순위 중간)

- [x] 에러 핸들링 개선 — Graph API 오류(400, 401, 403, 404, 429, 503) 시 사용자 친화적 메시지 반환
- [x] 429/503/504 자동 재시도 — `_request_with_retry()`로 `Retry-After`/지수 백오프 기반 재시도
- [x] 테스트 추가 — `pytest` + mock으로 각 도구의 포맷팅 로직 및 재시도 로직 검증 (`tests/test_server.py`)
- [x] nextLink 기반 페이지네이션 — `$skip` 미지원 엔드포인트용 `next_link` 파라미터 추가

## 사용성 (우선순위 낮음)

- [x] 첨부파일 다운로드 — 메일/Teams 메시지 첨부파일 읽기 (`list_email_attachments`, `read_email_attachment`, `list_message_attachments`, `read_message_attachment`)
- [x] 첨부파일 바이너리 저장 — 텍스트 추출 대신 디스크에 원본 파일 저장 (`download_attachment`)
- [x] 회의 초대 응답 (`respond_to_event`) — 수락/거절/임시수락
- [x] 메일 초안 작성 (`create_draft_email`) — 발송 없이 Drafts 저장
- [x] 첨부파일 메일 발송 (`send_email_with_attachment`) — 로컬 파일 첨부

## 향후 후보 (새 권한 필요)

- [ ] `find_meeting_times` / `get_schedule` — 참석자 빈 시간 탐색 (`Calendars.Read.Shared`)
- [ ] `get_presence` — 동료 현재 상태 조회 (`Presence.Read.All`)
- [ ] Microsoft To Do 연동 (`Tasks.ReadWrite`)
- [ ] `set_automatic_replies` — 부재중 자동응답 (`MailboxSettings.ReadWrite`)
