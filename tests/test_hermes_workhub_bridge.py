import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BRIDGE_PATH = ROOT / "deploy" / "scripts" / "hermes_workhub_bridge.py"


def load_bridge_module():
    spec = importlib.util.spec_from_file_location("hermes_workhub_bridge", BRIDGE_PATH)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class HermesWorkhubBridgeTests(unittest.TestCase):
    def test_bridge_builds_chat_prompt_from_workhub_message(self) -> None:
        bridge = load_bridge_module()

        prompt = bridge.build_prompt({"message": "summarize open CS cases", "source": "workhub"}, "chat")

        self.assertIn("Workhub", prompt)
        self.assertIn("summarize open CS cases", prompt)
        self.assertIn("Korean", prompt)

    def test_bridge_prompt_includes_workhub_context_snapshot(self) -> None:
        bridge = load_bridge_module()

        prompt = bridge.build_prompt({
            "message": "오늘 매출 요약해줘",
            "workhub_context": {
                "sales_report": {"period": "2026-06", "month": {"sales_amount": 12345}},
                "cs_status_counts": [{"status": "open", "count": 2}],
            },
            "capabilities": {"workhub_context": True, "requested_intent": "chat"},
        }, "chat")

        self.assertIn("Workhub context snapshot", prompt)
        self.assertIn("2026-06", prompt)
        self.assertIn("Available Workhub AI capabilities", prompt)

    def test_bridge_authorization_accepts_bearer_or_x_hermes_key(self) -> None:
        bridge = load_bridge_module()

        self.assertTrue(bridge.is_authorized({"Authorization": "Bearer token-123"}, "token-123"))
        self.assertTrue(bridge.is_authorized({"X-Hermes-Api-Key": "Bearer token-123"}, "token-123"))
        self.assertFalse(bridge.is_authorized({"Authorization": "Bearer wrong"}, "token-123"))

    def test_bridge_automation_prompt_includes_title_and_body(self) -> None:
        bridge = load_bridge_module()

        prompt = bridge.build_prompt({"title": "settlement check", "body": "find missing vendor purchases"}, "automation")

        self.assertIn("settlement check", prompt)
        self.assertIn("find missing vendor purchases", prompt)
        self.assertIn("actionable steps", prompt)

    def test_bridge_routes_search_and_image_intents(self) -> None:
        bridge = load_bridge_module()

        self.assertEqual(
            bridge.requested_intent({"message": "최신 택배비 정책 검색해줘"}, "chat"),
            "web_search",
        )
        self.assertEqual(
            bridge.requested_intent({"message": "상품 배너 이미지 만들어줘"}, "chat"),
            "image_generation",
        )
        self.assertEqual(
            bridge.requested_intent({"message": "최신 자료 조사해줘"}, "automation"),
            "",
        )

    def test_bridge_uses_shared_hermes_backend_for_tool_intents_by_default(self) -> None:
        bridge = load_bridge_module()

        bridge.AI_TOOL_PROVIDER = "hermes"
        bridge.OPENAI_API_KEY = ""
        self.assertFalse(bridge.should_use_openai_for_intent("web_search"))
        self.assertFalse(bridge.should_use_openai_for_intent("image_generation"))

        bridge.AI_TOOL_PROVIDER = "openai"
        self.assertTrue(bridge.should_use_openai_for_intent("web_search"))

        bridge.AI_TOOL_PROVIDER = "auto"
        bridge.OPENAI_API_KEY = "set"
        self.assertFalse(bridge.should_use_openai_for_intent("image_generation"))

        bridge.AI_TOOL_PROVIDER = ""
        self.assertFalse(bridge.should_use_openai_for_intent("web_search"))

    def test_bridge_prompt_marks_shared_tool_intent_for_hermes(self) -> None:
        bridge = load_bridge_module()

        prompt = bridge.build_prompt({"message": "최신 자료 조사해줘", "intent": "web_search"}, "chat")

        self.assertIn("Requested tool intent: web_search", prompt)
        self.assertIn("shared Hermes research/search backend", prompt)

    def test_bridge_blocks_unconfigured_shared_image_generation(self) -> None:
        bridge = load_bridge_module()

        bridge.AI_TOOL_PROVIDER = "hermes"
        bridge.FAL_KEY = ""
        bridge.OPENAI_API_KEY = ""

        self.assertTrue(bridge.should_block_unconfigured_image_generation("image_generation"))
        response = bridge.image_generation_not_configured_response()
        self.assertFalse(response["ok"])
        self.assertIn("FAL_KEY", response["answer"])

        bridge.FAL_KEY = "set"
        self.assertFalse(bridge.should_block_unconfigured_image_generation("image_generation"))
