import base64
import logging

from openai import OpenAI, APIConnectionError, APITimeoutError, RateLimitError, InternalServerError
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    before_sleep_log,
)

from ai.service import AIService

logger = logging.getLogger(__name__)

_DEFAULT_MODEL = "gpt-5.2"

_MAX_RETRIES = 4
_MIN_WAIT_SECONDS = 2
_MAX_WAIT_SECONDS = 60

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


class OpenAIService(AIService):
    """AIService backed by the OpenAI API (text-only decisions)."""

    def __init__(self, model: str = _DEFAULT_MODEL):
        self._model = model
        self._client = OpenAI()  # reads OPENAI_API_KEY from env

    @_retry_decorator
    def get_decision(self, prompt: str) -> str:
        response = self._client.chat.completions.create(
            model=self._model,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.choices[0].message.content or ""

    @_retry_decorator
    def get_decision_for_media(
        self, prompt: str, image_bytes: bytes, mime_type: str = "image/png"
    ) -> str:
        b64 = base64.b64encode(image_bytes).decode()
        data_url = f"data:{mime_type};base64,{b64}"
        response = self._client.chat.completions.create(
            model=self._model,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {"url": data_url},
                        },
                    ],
                }
            ],
        )
        return response.choices[0].message.content or ""
