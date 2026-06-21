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
    list_teams,
    list_emails,
    search_emails,
    send_email,
    reply_email,
    forward_email,
    create_chat,
    list_chats,
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
