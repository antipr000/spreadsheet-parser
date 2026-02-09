"""
AIService implementation backed by the Anthropic Claude API.

Supports:
  - Text-only prompts  (get_decision)
  - Image + text       (get_decision_for_media) — images (png, jpeg,
    gif, webp) are sent as base64 image content blocks.  For any
    non-image media type the file is skipped and only the text prompt
    is sent.

Reads ANTHROPIC_API_KEY from the environment.
Default model: claude-opus-4-6
"""

from __future__ import annotations

import base64
import logging

from anthropic import Anthropic

from ai.service import AIService

logger = logging.getLogger(__name__)

_DEFAULT_MODEL = "claude-opus-4-6"

# MIME types that Claude accepts as image content blocks.
_IMAGE_MIMES = frozenset(
    {
        "image/png",
        "image/jpeg",
        "image/gif",
        "image/webp",
    }
)


class ClaudeService(AIService):
    """AIService backed by the Anthropic Claude API."""

    def __init__(self, model: str = _DEFAULT_MODEL):
        self._model = model
        self._client = Anthropic()  # reads ANTHROPIC_API_KEY from env

    def get_decision(self, prompt: str) -> str:
        message = self._client.messages.create(
            model=self._model,
            max_tokens=16384,
            messages=[{"role": "user", "content": prompt}],
        )
        return message.content[0].text if message.content else ""

    def get_decision_for_media(
        self,
        prompt: str,
        image_bytes: bytes,
        mime_type: str = "image/png",
    ) -> str:
        # Images: send as image content block
        if mime_type in _IMAGE_MIMES:
            b64_data = base64.standard_b64encode(image_bytes).decode("ascii")
            message = self._client.messages.create(
                model=self._model,
                max_tokens=16384,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": mime_type,
                                    "data": b64_data,
                                },
                            },
                            {"type": "text", "text": prompt},
                        ],
                    }
                ],
            )
            return message.content[0].text if message.content else ""

        # Non-image media: Claude doesn't support xlsx etc., send text only
        logger.debug(
            "  [Claude] Unsupported media type %s — sending text-only prompt",
            mime_type,
        )
        return self.get_decision(prompt)
