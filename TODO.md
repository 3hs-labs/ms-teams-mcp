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

## 품질 개선 (우선순위 중간)

- [ ] 에러 핸들링 개선 — Graph API 오류(401, 403, 404, 429 rate limit) 시 사용자 친화적 메시지 반환
- [ ] 테스트 추가 — `pytest` + mock으로 각 도구의 포맷팅 로직 검증
- [x] nextLink 기반 페이지네이션 — `$skip` 미지원 엔드포인트용 `next_link` 파라미터 추가

## 사용성 (우선순위 낮음)

- [ ] 첨부파일 다운로드 — 메일 첨부파일 읽기 기능
- [ ] PyPI 배포 — `pip install microsoft-teams-mcp`로 설치 간소화
