import json
import unittest
from unittest import mock

import tradeflow_server


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload
        self.read_calls = 0

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self):
        self.read_calls += 1
        return json.dumps(self.payload).encode("utf-8")


class TestNotionContext(unittest.TestCase):
    def test_notion_relevant_context_returns_filtered_page_context(self):
        search_payload = {
            "results": [
                {
                    "object": "page",
                    "id": "page-1",
                    "properties": {
                        "title": {"title": [{"plain_text": "Q4 promotion plan"}]}
                    }
                }
            ]
        }
        blocks_payload = {
            "results": [
                {
                    "type": "paragraph",
                    "paragraph": {
                        "rich_text": [
                            {"type": "text", "text": {"content": "Focus on promotion uplift and base vs incremental volume."}}
                        ]
                    },
                },
                {
                    "type": "heading_2",
                    "heading_2": {
                        "rich_text": [{"type": "text", "text": {"content": "Key actions"}}]
                    },
                },
                {
                    "type": "bulleted_list_item",
                    "bulleted_list_item": {
                        "rich_text": [{"type": "text", "text": {"content": "Validate retailer funding and margin"}}]
                    },
                },
            ]
        }

        with mock.patch.dict(tradeflow_server.os.environ, {"NOTION_API": "secret_test"}, clear=False):
            with mock.patch.object(tradeflow_server, "trusted_urlopen", side_effect=[FakeResponse(search_payload), FakeResponse(blocks_payload)]):
                entries = tradeflow_server.notion_relevant_context("promotion ROI", limit=1)

        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["title"], "Q4 promotion plan")
        self.assertIn("promotion uplift", entries[0]["excerpt"].lower())
        self.assertEqual(entries[0]["source"], "Notion · Q4 promotion plan")


if __name__ == "__main__":
    unittest.main()
