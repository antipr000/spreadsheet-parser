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

Retries transient errors (rate-limit, overloaded, connection, timeout)
with exponential backoff via tenacity.
"""

from __future__ import annotations

import base64
import logging

from anthropic import (
    Anthropic,
    APIConnectionError,
    APITimeoutError,
    RateLimitError,
    InternalServerError,
)
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    before_sleep_log,
)

from ai.service import AIService

logger = logging.getLogger(__name__)

_DEFAULT_MODEL = "claude-opus-4-6"

# Retry configuration
_MAX_RETRIES = 4
_MIN_WAIT_SECONDS = 2
_MAX_WAIT_SECONDS = 60

# Transient exception types that should trigger a retry.
_RETRYABLE_EXCEPTIONS = (
    APIConnectionError,
    APITimeoutError,
    RateLimitError,
    InternalServerError,
)

_retry_decorator = retry(
    retry=retry_if_exception_type(_RETRYABLE_EXCEPTIONS),
    stop=stop_after_attempt(_MAX_RETRIES),
    wait=wait_exponential(multiplier=1, min=_MIN_WAIT_SECONDS, max=_MAX_WAIT_SECONDS),
    before_sleep=before_sleep_log(logger, logging.WARNING),
    reraise=True,
)

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

    @_retry_decorator
    def get_decision(self, prompt: str) -> str:
        message = self._client.messages.create(
            model=self._model,
            max_tokens=16384,
            messages=[{"role": "user", "content": prompt}],
        )
        return message.content[0].text if message.content else ""

    @_retry_decorator
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
