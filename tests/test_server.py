"""
MCP Tools formatting logic tests.
Uses unittest.mock to patch graph_get, graph_post, requests.post, and _headers
so that no real API calls are made.
"""

import pytest
from unittest.mock import patch, MagicMock  # noqa: F401 - MagicMock used in TestCheckResponse

import ms_teams_mcp.server as server
from ms_teams_mcp.server import (
    strip_html,
    _pagination_footer,
    _check_response,
    _request_with_retry,
    _retry_after_seconds,
    graph_get,
    _apply_to_messages,
    _resolve_folder_id,
    list_teams,
    list_emails,
    search_emails,
    search_messages,
    send_email,
    reply_email,
    forward_email,
    create_draft_email,
    send_email_with_attachment,
    respond_to_event,
    create_chat,
    list_chats,
    mark_email_read,
    flag_email,
    move_email,
    delete_email,
    download_attachment,
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
        assert "More data available" in footer

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
        with pytest.raises(Exception, match="Auth error"):
            _check_response(res)

    def test_403_raises(self):
        res = self._make_response(403, {"error": {"message": "forbidden"}})
        with pytest.raises(Exception, match="Permission denied"):
            _check_response(res)

    def test_404_raises(self):
        res = self._make_response(404, {"error": {"message": "not found"}})
        with pytest.raises(Exception, match="Resource not found"):
            _check_response(res)

    def test_429_raises(self):
        res = self._make_response(429, {"error": {"message": "throttled"}})
        with pytest.raises(Exception, match="Rate limit exceeded"):
            _check_response(res)

    def test_400_raises(self):
        res = self._make_response(400, {"error": {"message": "malformed"}})
        with pytest.raises(Exception, match="Bad request"):
            _check_response(res)

    def test_503_raises(self):
        res = self._make_response(503, {"error": {"message": "unavailable"}})
        with pytest.raises(Exception, match="Service unavailable"):
            _check_response(res)

    def test_200_no_raise(self):
        res = self._make_response(200)
        _check_response(res)  # should not raise


# ──────────────────────────────────────────
# 3b. test_request_with_retry
# ──────────────────────────────────────────

class TestRequestWithRetry:
    def _make_response(self, status_code, headers=None):
        res = MagicMock()
        res.status_code = status_code
        res.ok = 200 <= status_code < 300
        res.headers = headers or {}
        return res

    def test_retry_after_header_honored(self):
        res = self._make_response(429, {"Retry-After": "7"})
        assert _retry_after_seconds(res, attempt=0) == 7

    def test_retry_after_caps_at_max(self):
        res = self._make_response(429, {"Retry-After": "999"})
        assert _retry_after_seconds(res, attempt=0) == 30  # MAX_BACKOFF_SECONDS

    def test_exponential_backoff_without_header(self):
        res = self._make_response(429, {})
        assert _retry_after_seconds(res, attempt=0) == 1
        assert _retry_after_seconds(res, attempt=2) == 4

    def test_invalid_retry_after_falls_back_to_backoff(self):
        res = self._make_response(429, {"Retry-After": "soon"})
        assert _retry_after_seconds(res, attempt=1) == 2

    @patch("ms_teams_mcp.server.time.sleep")
    @patch("ms_teams_mcp.server.requests.request")
    def test_success_no_retry(self, mock_request, mock_sleep):
        mock_request.return_value = self._make_response(200)
        res = _request_with_retry("GET", "http://x")
        assert res.status_code == 200
        assert mock_request.call_count == 1
        mock_sleep.assert_not_called()

    @patch("ms_teams_mcp.server.time.sleep")
    @patch("ms_teams_mcp.server.requests.request")
    def test_retries_on_429_then_succeeds(self, mock_request, mock_sleep):
        mock_request.side_effect = [
            self._make_response(429, {"Retry-After": "1"}),
            self._make_response(429, {"Retry-After": "1"}),
            self._make_response(200),
        ]
        res = _request_with_retry("GET", "http://x")
        assert res.status_code == 200
        assert mock_request.call_count == 3
        assert mock_sleep.call_count == 2

    @patch("ms_teams_mcp.server.time.sleep")
    @patch("ms_teams_mcp.server.requests.request")
    def test_exhausts_retries_returns_last(self, mock_request, mock_sleep):
        mock_request.return_value = self._make_response(429, {})
        res = _request_with_retry("GET", "http://x")
        assert res.status_code == 429
        # 1 initial + MAX_RETRIES (3) re-attempts
        assert mock_request.call_count == 4
        assert mock_sleep.call_count == 3

    @patch("ms_teams_mcp.server.time.sleep")
    @patch("ms_teams_mcp.server.requests.request")
    def test_no_retry_on_non_retryable_error(self, mock_request, mock_sleep):
        mock_request.return_value = self._make_response(403)
        res = _request_with_retry("GET", "http://x")
        assert res.status_code == 403
        assert mock_request.call_count == 1
        mock_sleep.assert_not_called()

    @patch("ms_teams_mcp.server._headers", return_value={"Authorization": "Bearer t"})
    @patch("ms_teams_mcp.server.time.sleep")
    @patch("ms_teams_mcp.server.requests.request")
    def test_graph_get_retries_then_raises(self, mock_request, mock_sleep, mock_headers):
        # Persistent 429 should retry then surface a friendly error via _check_response.
        throttled = self._make_response(429)
        throttled.text = "throttled"
        throttled.json.return_value = {"error": {"message": "throttled"}}
        mock_request.return_value = throttled
        with pytest.raises(Exception, match="Rate limit exceeded"):
            graph_get("/me")
        assert mock_request.call_count == 4


# ──────────────────────────────────────────
# 3c. test_main_stdio_output
# ──────────────────────────────────────────

class TestMainStdioOutput:
    def test_stdio_startup_message_goes_to_stderr(self, monkeypatch, capsys):
        monkeypatch.setattr(server.sys, "argv", ["ms-teams-mcp"])
        monkeypatch.setattr(server, "_check_and_auto_update", lambda: None)
        mock_run = MagicMock()
        monkeypatch.setattr(server.mcp, "run", mock_run)

        server.main()

        captured = capsys.readouterr()
        assert captured.out == ""
        assert "Starting Microsoft Teams MCP server..." in captured.err
        mock_run.assert_called_once_with()


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
        assert "No teams found." in result


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
        assert "More data available" in result
        assert "next_link" in result


# ──────────────────────────────────────────
# 7. test_send_email
# ──────────────────────────────────────────

class TestSendEmail:
    @patch("ms_teams_mcp.server.graph_post_action")
    def test_send_email_success(self, mock_post_action):
        result = send_email(to="user@example.com", subject="Hi", body="Hello")
        assert "Email sent" in result
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
        assert "Reply sent" in result
        assert "msg123" in result

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_reply_all(self, mock_post_action):
        result = reply_email(message_id="msg123", body="Thanks", reply_all=True)
        assert "Reply all sent" in result
        assert "msg123" in result


# ──────────────────────────────────────────
# 9. test_forward_email
# ──────────────────────────────────────────

class TestForwardEmail:
    @patch("ms_teams_mcp.server.graph_post_action")
    def test_forward_success(self, mock_post_action):
        result = forward_email(message_id="msg456", to="other@example.com", comment="FYI")
        assert "Email forwarded" in result
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
        assert "Chat created" in result
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
        assert "No results found" in result


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

    @patch("ms_teams_mcp.server.graph_get")
    def test_scan_failure_degrades(self, mock_get):
        def _se(path, params=None, url=None):
            if path == "/me/joinedTeams":
                raise Exception("teams down")
            return _briefing_side_effect(path, params, url)
        mock_get.side_effect = _se
        result = daily_briefing()
        assert "Channel Activity" in result
        assert "Failed to retrieve" in result
        assert "(unavailable)" in result


# ──────────────────────────────────────────
# search_messages (Teams chat & channel search via /search/query)
# ──────────────────────────────────────────

def _search_response(hits, more=False):
    """Build a Graph /search/query response wrapping the given hits."""
    return {
        "value": [
            {
                "searchTerms": ["x"],
                "hitsContainers": [
                    {"hits": hits, "total": len(hits), "moreResultsAvailable": more}
                ],
            }
        ]
    }


_CHAT_HIT = {
    "hitId": "h1",
    "rank": 1,
    "summary": "...the <c0>budget</c0> review is next week...",
    "resource": {
        "@odata.type": "#microsoft.graph.chatMessage",
        "id": "1657782060227",
        "createdDateTime": "2026-06-20T07:01:01Z",
        "from": {"emailAddress": {"name": "Alice Kim", "address": "alice@x.com"}},
        "channelIdentity": {},
        "chatId": "19:abc@thread.v2",
        "webUrl": "https://teams.microsoft.com/l/message/chat",
    },
}

_CHANNEL_HIT = {
    "hitId": "h2",
    "rank": 2,
    "summary": "deployment <c0>budget</c0> approved",
    "resource": {
        "@odata.type": "#microsoft.graph.chatMessage",
        "id": "1657782099999",
        "createdDateTime": "2026-06-19T03:30:00Z",
        "from": {"emailAddress": {"name": "Bob Lee", "address": "bob@x.com"}},
        "channelIdentity": {"teamId": "t1", "channelId": "c1"},
        "webUrl": "https://teams.microsoft.com/l/message/channel",
    },
}


class TestSearchMessages:
    @patch("ms_teams_mcp.server.graph_post")
    def test_mixed_chat_and_channel(self, mock_post):
        mock_post.return_value = _search_response([_CHAT_HIT, _CHANNEL_HIT])
        result = search_messages(query="budget")
        # Both labels appear, in one combined list
        assert "[Chat]" in result
        assert "[Channel]" in result
        # Sender names rendered
        assert "Alice Kim" in result
        assert "Bob Lee" in result
        # Dates rendered (date-only)
        assert "2026-06-20" in result
        assert "2026-06-19" in result
        # Summary snippet rendered with HTML hit-highlight tags stripped
        assert "budget" in result
        assert "<c0>" not in result
        # webUrl rendered for click-through
        assert "https://teams.microsoft.com/l/message/chat" in result

    @patch("ms_teams_mcp.server.graph_post")
    def test_empty_results(self, mock_post):
        mock_post.return_value = _search_response([])
        result = search_messages(query="nonexistent")
        assert "No results found for 'nonexistent'." == result

    @patch("ms_teams_mcp.server.graph_post")
    def test_empty_hits_container(self, mock_post):
        # API may omit hitsContainers entirely when there are zero hits
        mock_post.return_value = {"value": [{"searchTerms": ["x"], "hitsContainers": []}]}
        result = search_messages(query="nope")
        assert "No results found for 'nope'." == result

    @patch("ms_teams_mcp.server.graph_post")
    def test_pagination_footer_when_more_available(self, mock_post):
        mock_post.return_value = _search_response([_CHAT_HIT], more=True)
        result = search_messages(query="budget", top=1, skip=0)
        assert "More results available" in result
        assert "skip=1" in result

    @patch("ms_teams_mcp.server.graph_post")
    def test_no_footer_when_no_more(self, mock_post):
        mock_post.return_value = _search_response([_CHAT_HIT], more=False)
        result = search_messages(query="budget")
        assert "More results available" not in result

    @patch("ms_teams_mcp.server.graph_post")
    def test_request_body_and_size_clamp(self, mock_post):
        mock_post.return_value = _search_response([_CHAT_HIT])
        search_messages(query="hello world", top=100, skip=20)
        path, body = mock_post.call_args.args
        assert path == "/search/query"
        req = body["requests"][0]
        assert req["entityTypes"] == ["chatMessage"]
        assert req["query"]["queryString"] == "hello world"
        assert req["size"] == 50  # clamped from 100 to API max
        assert req["from"] == 20


# ──────────────────────────────────────────
# test_download_attachment
# ──────────────────────────────────────────

class TestDownloadAttachment:
    @patch("ms_teams_mcp.server.graph_get")
    def test_email_attachment_with_content_bytes(self, mock_graph_get, tmp_path):
        import base64
        payload = b"hello binary data"
        mock_graph_get.return_value = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "report.pdf",
            "size": len(payload),
            "contentBytes": base64.b64encode(payload).decode("ascii"),
        }
        result = download_attachment(
            message_id="m1", attachment_id="a1", save_dir=str(tmp_path)
        )
        saved = tmp_path / "report.pdf"
        assert saved.exists()
        assert saved.read_bytes() == payload
        assert "report.pdf" in result
        assert str(len(payload)) in result

    @patch("ms_teams_mcp.server.graph_get_binary")
    @patch("ms_teams_mcp.server.graph_get")
    def test_email_attachment_via_value_endpoint(self, mock_graph_get,
                                                 mock_binary, tmp_path):
        mock_graph_get.return_value = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "data.bin",
            "size": 5,
            "contentBytes": None,
        }
        resp = MagicMock()
        resp.ok = True
        resp.status_code = 200
        resp.content = b"12345"
        mock_binary.return_value = resp
        result = download_attachment(
            message_id="m1", attachment_id="a1", save_dir=str(tmp_path)
        )
        assert (tmp_path / "data.bin").read_bytes() == b"12345"
        assert "data.bin" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_email_attachment_too_large(self, mock_graph_get, tmp_path):
        mock_graph_get.return_value = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "huge.zip",
            "size": 50 * 1024 * 1024,
        }
        result = download_attachment(
            message_id="m1", attachment_id="a1", save_dir=str(tmp_path), max_mb=25
        )
        assert "too large" in result.lower()
        assert not (tmp_path / "huge.zip").exists()

    @patch("ms_teams_mcp.server.graph_get")
    def test_item_attachment_rejected(self, mock_graph_get, tmp_path):
        mock_graph_get.return_value = {
            "@odata.type": "#microsoft.graph.itemAttachment",
            "name": "Forwarded mail",
        }
        result = download_attachment(
            message_id="m1", attachment_id="a1", save_dir=str(tmp_path)
        )
        assert "not a downloadable file" in result

    @patch("ms_teams_mcp.server._download_shared_drive_item")
    def test_teams_attachment_via_content_url(self, mock_download, tmp_path):
        mock_download.return_value = ("slides.pptx", b"PPTXBYTES", "application/octet-stream")
        result = download_attachment(
            content_url="https://sharepoint/x/slides.pptx", save_dir=str(tmp_path)
        )
        assert (tmp_path / "slides.pptx").read_bytes() == b"PPTXBYTES"
        assert "slides.pptx" in result

    def test_missing_args_returns_guidance(self, tmp_path):
        result = download_attachment(save_dir=str(tmp_path))
        assert "Provide message_id" in result

    @patch("ms_teams_mcp.server.graph_get")
    def test_path_traversal_in_name_is_stripped(self, mock_graph_get, tmp_path):
        import base64
        mock_graph_get.return_value = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "../../evil.sh",
            "size": 3,
            "contentBytes": base64.b64encode(b"rm ").decode("ascii"),
        }
        result = download_attachment(
            message_id="m1", attachment_id="a1", save_dir=str(tmp_path)
        )
        # basename only — saved inside tmp_path, not in a parent dir
        assert (tmp_path / "evil.sh").exists()
        assert "evil.sh" in result


# ──────────────────────────────────────────
# test_create_draft_email
# ──────────────────────────────────────────

class TestCreateDraftEmail:
    @patch("ms_teams_mcp.server.graph_post")
    def test_creates_draft_with_recipients(self, mock_post):
        mock_post.return_value = {"id": "draft-123"}
        result = create_draft_email(
            to="a@example.com,b@example.com", subject="Hi", body="Hello"
        )
        path, message = mock_post.call_args.args
        assert path == "/me/messages"
        assert message["subject"] == "Hi"
        assert message["body"]["content"] == "Hello"
        assert len(message["toRecipients"]) == 2
        assert "draft-123" in result
        assert "Drafts" in result

    @patch("ms_teams_mcp.server.graph_post")
    def test_draft_without_recipients(self, mock_post):
        mock_post.return_value = {"id": "d1"}
        result = create_draft_email(subject="Memo")
        _path, message = mock_post.call_args.args
        assert "toRecipients" not in message
        assert "d1" in result


# ──────────────────────────────────────────
# test_send_email_with_attachment
# ──────────────────────────────────────────

class TestSendEmailWithAttachment:
    @patch("ms_teams_mcp.server.graph_post_action")
    def test_sends_with_attachment(self, mock_action, tmp_path):
        f = tmp_path / "doc.txt"
        f.write_bytes(b"file content")
        result = send_email_with_attachment(
            to="x@example.com", subject="S", body="B", attachments=str(f)
        )
        path, payload = mock_action.call_args.args
        assert path == "/me/sendMail"
        atts = payload["message"]["attachments"]
        assert len(atts) == 1
        assert atts[0]["name"] == "doc.txt"
        assert atts[0]["@odata.type"] == "#microsoft.graph.fileAttachment"
        assert "doc.txt" in result

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_missing_file_returns_error(self, mock_action, tmp_path):
        result = send_email_with_attachment(
            to="x@example.com", subject="S", body="B",
            attachments=str(tmp_path / "nope.txt"),
        )
        assert "File not found" in result
        mock_action.assert_not_called()

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_no_attachments_returns_error(self, mock_action):
        result = send_email_with_attachment(
            to="x@example.com", subject="S", body="B", attachments="  "
        )
        assert "No attachment paths" in result
        mock_action.assert_not_called()

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_oversize_attachment_rejected(self, mock_action, tmp_path):
        big = tmp_path / "big.bin"
        big.write_bytes(b"0" * (4 * 1024 * 1024))  # 4MB > 3MB cap
        result = send_email_with_attachment(
            to="x@example.com", subject="S", body="B", attachments=str(big)
        )
        assert "exceeds" in result.lower()
        mock_action.assert_not_called()


# ──────────────────────────────────────────
# test_respond_to_event
# ──────────────────────────────────────────

class TestRespondToEvent:
    @patch("ms_teams_mcp.server.graph_post_action")
    def test_accept(self, mock_action):
        result = respond_to_event(event_id="e1", response="accept")
        path, body = mock_action.call_args.args
        assert path == "/me/events/e1/accept"
        assert body["sendResponse"] is True
        assert "accepted" in result

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_tentative_maps_to_endpoint(self, mock_action):
        result = respond_to_event(event_id="e1", response="tentative", comment="maybe")
        path, body = mock_action.call_args.args
        assert path == "/me/events/e1/tentativelyAccept"
        assert body["comment"] == "maybe"
        assert "tentatively accepted" in result

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_decline_with_no_response(self, mock_action):
        result = respond_to_event(event_id="e1", response="decline", send_response=False)
        path, body = mock_action.call_args.args
        assert path == "/me/events/e1/decline"
        assert body["sendResponse"] is False
        assert "declined" in result

    @patch("ms_teams_mcp.server.graph_post_action")
    def test_invalid_response_rejected(self, mock_action):
        result = respond_to_event(event_id="e1", response="maybe")
        assert "Invalid response" in result
        mock_action.assert_not_called()
