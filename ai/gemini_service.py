import logging

from google import genai
from google.genai import types

from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    before_sleep_log,
)

from ai.service import AIService

logger = logging.getLogger(__name__)

_DEFAULT_MODEL = "gemini-2.5-flash"

_MAX_RETRIES = 4
_MIN_WAIT_SECONDS = 2
_MAX_WAIT_SECONDS = 60

_RETRYABLE_EXCEPTIONS = (ConnectionError,)

_retry_decorator = retry(
    retry=retry_if_exception_type(_RETRYABLE_EXCEPTIONS),
    stop=stop_after_attempt(_MAX_RETRIES),
    wait=wait_exponential(multiplier=1, min=_MIN_WAIT_SECONDS, max=_MAX_WAIT_SECONDS),
    before_sleep=before_sleep_log(logger, logging.WARNING),
    reraise=True,
)


class GeminiService(AIService):
    """AIService backed by the Google Gemini API."""

    def __init__(self, model: str = _DEFAULT_MODEL):
        self._model = model
        self._client = genai.Client()  # reads GEMINI_API_KEY / GOOGLE_API_KEY from env

    @_retry_decorator
    def get_decision(self, prompt: str) -> str:
        response = self._client.models.generate_content(
            model=self._model,
            contents=prompt,
        )
        return response.text or ""

    @_retry_decorator
    def get_decision_for_media(
        self, prompt: str, image_bytes: bytes, mime_type: str = "image/png"
    ) -> str:
        image_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)
        response = self._client.models.generate_content(
            model=self._model,
            contents=[prompt, image_part],
        )
        return response.text or ""
