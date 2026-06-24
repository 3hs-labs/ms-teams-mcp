"""
MCP Tools formatting logic tests.
Uses unittest.mock to patch graph_get, graph_post, requests.post, and _headers
so that no real API calls are made.
"""

import pytest
from unittest.mock import patch, MagicMock  # noqa: F401 - MagicMock used in TestCheckResponse

from ms_teams_mcp.server import (
    strip_html,
    _pagination_footer,
    _check_response,
    _apply_to_messages,
    _resolve_folder_id,
    list_teams,
    list_emails,
    search_emails,
    send_email,
    reply_email,
    forward_email,
    create_chat,
    list_chats,
    mark_email_read,
    flag_email,
    move_email,
    delete_email,
    _briefing_calendar_today,
    _briefing_unread_email,
    _briefing_recent_chats,
    _briefing_flagged_email,
    _scan_channels,
    daily_briefing,
)


# ──────────────────────────────────────────
# 1. test_strip_html
# ──────────────────────────────────────────

class TestStripHtml:
    def test_basic_html(self):
        assert strip_html("<p>Hello <b>World</b></p>") == "Hello World"

    def test_empty_string(self):
        assert strip_html("") == ""

    def test_none_input(self):
        assert strip_html(None) == ""

    def test_multiple_newlines_collapsed(self):
        html = "<p>Line1</p>\n\n\n\n<p>Line2</p>"
        result = strip_html(html)
        assert "\n\n\n" not in result

    def test_plain_text_passthrough(self):
        assert strip_html("no tags here") == "no tags here"


# ──────────────────────────────────────────
# 2. test_pagination_footer
# ──────────────────────────────────────────

class TestPaginationFooter:
    def test_with_next_link(self):
        data = {"@odata.nextLink": "https://graph.microsoft.com/v1.0/next"}
        footer = _pagination_footer(data, skip=0, top=10)
        assert "next_link" in footer
        assert "https://graph.microsoft.com/v1.0/next" in footer
        assert "추가 데이터 있음" in footer

    def test_without_next_link(self):
        data = {"value": []}
        footer = _pagination_footer(data, skip=0, top=10)
        assert footer == ""


# ──────────────────────────────────────────
# 3. test_check_response
# ──────────────────────────────────────────

class TestCheckResponse:
    def _make_response(self, status_code, json_body=None):
        res = MagicMock()
        res.status_code = status_code
        res.ok = 200 <= status_code < 300
        res.text = str(json_body or "")
        if json_body:
            res.json.return_value = json_body
        else:
            res.json.side_effect = Exception("no json")
        return res

    def test_401_raises(self):
        res = self._make_response(401, {"error": {"message": "token expired"}})
        with pytest.raises(Exception, match="인증 오류"):
            _check_response(res)

    def test_403_raises(self):
        res = self._make_response(403, {"error": {"message": "forbidden"}})
        with pytest.raises(Exception, match="권한 부족"):
            _check_response(res)

    def test_404_raises(self):
        res = self._make_response(404, {"error": {"message": "not found"}})
        with pytest.raises(Exception, match="리소스를 찾을 수 없습니다"):
            _check_response(res)

    def test_429_raises(self):
        res = self._make_response(429, {"error": {"message": "throttled"}})
        with pytest.raises(Exception, match="요청 한도 초과"):
            _check_response(res)

    def test_200_no_raise(self):
        res = self._make_response(200)
        _check_response(res)  # should not raise


# ──────────────────────────────────────────
# 4. test_list_teams
# ──────────────────────────────────────────

class TestListTeams:
    @patch("ms_teams_mcp.server.graph_get")
    def test_list_teams_formatting(self, mock_graph_get):
        mock_graph_get.return_value = {
            "value": [
                {"displayName": "Team1", "description": "desc", "id": "id1"}
            ]
        }
        result = list_teams()
        assert "Team1" in result
        assert "id1" in result
        assert "desc" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_list_teams_empty(self, mock_graph_get):
        mock_graph_get.return_value = {"value": []}
        result = list_teams()
        assert "참여한 팀이 없습니다" in result


# ──────────────────────────────────────────
# 5. test_list_emails
# ──────────────────────────────────────────

class TestListEmails:
    @patch("ms_teams_mcp.server.graph_get")
    def test_list_emails_formatting(self, mock_graph_get):
        mock_graph_get.return_value = {
            "value": [
                {
                    "id": "mail1",
                    "subject": "Test Subject",
                    "from": {
                        "emailAddress": {
                            "name": "Sender Name",
                            "address": "sender@example.com",
                        }
                    },
                    "receivedDateTime": "2025-01-15T10:30:00Z",
                    "isRead": False,
                    "bodyPreview": "Hello this is a preview",
                }
            ]
        }
        result = list_emails()
        assert "Test Subject" in result
        assert "Sender Name" in result
        assert "sender@example.com" in result
        assert "2025-01-15" in result
        assert "mail1" in result


# ──────────────────────────────────────────
# 6. test_list_emails_pagination
# ──────────────────────────────────────────

class TestListEmailsPagination:
    @patch("ms_teams_mcp.server.graph_get")
    def test_pagination_footer_appears(self, mock_graph_get):
        mock_graph_get.return_value = {
            "value": [
                {
                    "id": "mail1",
                    "subject": "Subject",
                    "from": {
                        "emailAddress": {"name": "A", "address": "a@b.com"}
                    },
                    "receivedDateTime": "2025-01-15T10:30:00Z",
                    "isRead": True,
                    "bodyPreview": "preview",
                }
            ],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=10",
        }
        result = list_emails()
        assert "추가 데이터 있음" in result
        assert "next_link" in result


# ──────────────────────────────────────────
# 7. test_send_email
# ──────────────────────────────────────────

class TestSendEmail:
    @patch("ms_teams_mcp.server.graph_post_action")
    def test_send_email_success(self, mock_post_action):
        result = send_email(to="user@example.com", subject="Hi", body="Hello")
        assert "메일 발송 완료" in result
        assert "user@example.com" in result
        assert "Hi" in result
        mock_post_action.assert_called_once()


# ──────────────────────────────────────────
# 8. test_reply_email
# ──────────────────────────────────────────

class TestReplyEmail:
    @patch("ms_teams_mcp.server.graph_post_action")
    def test_reply(self, mock_post_action):
        result = reply_email(message_id="msg123", body="Thanks")
        assert "답장 완료" in result
        assert "msg123" in result

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_reply_all(self, mock_post_action):
        result = reply_email(message_id="msg123", body="Thanks", reply_all=True)
        assert "전체 답장 완료" in result
        assert "msg123" in result


# ──────────────────────────────────────────
# 9. test_forward_email
# ──────────────────────────────────────────

class TestForwardEmail:
    @patch("ms_teams_mcp.server.graph_post_action")
    def test_forward_success(self, mock_post_action):
        result = forward_email(message_id="msg456", to="other@example.com", comment="FYI")
        assert "메일 전달 완료" in result
        assert "other@example.com" in result
        assert "msg456" in result


# ──────────────────────────────────────────
# 10. test_create_chat
# ──────────────────────────────────────────

class TestCreateChat:
    @patch("ms_teams_mcp.server.graph_post")
    @patch("ms_teams_mcp.server.graph_get")
    def test_create_chat_success(self, mock_graph_get, mock_graph_post):
        mock_graph_get.return_value = {"id": "my-user-id"}
        mock_graph_post.return_value = {"id": "new-chat-id"}

        result = create_chat(members="user@example.com")
        assert "채팅 생성 완료" in result
        assert "new-chat-id" in result
        assert "user@example.com" in result


# ──────────────────────────────────────────
# 11. test_list_chats_with_next_link
# ──────────────────────────────────────────

class TestListChatsWithNextLink:
    @patch("ms_teams_mcp.server.graph_get")
    def test_next_link_passed_as_url(self, mock_graph_get):
        next_url = "https://graph.microsoft.com/v1.0/me/chats?$skip=20"
        mock_graph_get.return_value = {
            "value": [
                {
                    "id": "chat1",
                    "chatType": "oneOnOne",
                    "topic": None,
                    "members": [{"displayName": "User1"}],
                    "lastMessagePreview": {
                        "body": {"content": "Hello"},
                        "createdDateTime": "2025-01-15T10:00:00Z",
                    },
                }
            ]
        }

        result = list_chats(next_link=next_url)

        # Verify graph_get was called with url parameter
        mock_graph_get.assert_called_once_with("", url=next_url)
        assert "chat1" in result


# ──────────────────────────────────────────
# 12. test_search_emails_empty
# ──────────────────────────────────────────

class TestSearchEmailsEmpty:
    @patch("ms_teams_mcp.server.graph_get")
    def test_empty_results(self, mock_graph_get):
        mock_graph_get.return_value = {"value": []}
        result = search_emails(query="nonexistent")
        assert "검색 결과가 없습니다" in result


# ──────────────────────────────────────────
# 13. test_apply_to_messages
# ──────────────────────────────────────────

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


# ──────────────────────────────────────────
# 6. test_resolve_folder_id
# ──────────────────────────────────────────

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


# ──────────────────────────────────────────
# 14. test_mark_email_read
# ──────────────────────────────────────────

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


# ──────────────────────────────────────────
# 13. test_flag_email
# ──────────────────────────────────────────

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


# ──────────────────────────────────────────
# 15. test_move_email
# ──────────────────────────────────────────

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


# ──────────────────────────────────────────
# 16. test_delete_email
# ──────────────────────────────────────────

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


# ──────────────────────────────────────────
# 17. test_briefing_simple_sections
# ──────────────────────────────────────────

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

    @patch("ms_teams_mcp.server.graph_get")
    def test_unread_failure_degrades(self, mock_get):
        mock_get.side_effect = Exception("boom")
        result = _briefing_unread_email()
        assert "Failed to retrieve" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_recent_chats_failure_degrades(self, mock_get):
        mock_get.side_effect = Exception("boom")
        result = _briefing_recent_chats()
        assert "Failed to retrieve" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_flagged_failure_degrades(self, mock_get):
        mock_get.side_effect = Exception("boom")
        result = _briefing_flagged_email()
        assert "Failed to retrieve" in result


# ──────────────────────────────────────────
# 18. test_scan_channels
# ──────────────────────────────────────────

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


# ──────────────────────────────────────────
# 19. test_daily_briefing
# ──────────────────────────────────────────

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
