from google import genai
from google.genai import types

from ai.service import AIService


_DEFAULT_MODEL = "gemini-2.5-flash"


class GeminiService(AIService):
    """AIService backed by the Google Gemini API."""

    def __init__(self, model: str = _DEFAULT_MODEL):
        self._model = model
        self._client = genai.Client()  # reads GEMINI_API_KEY / GOOGLE_API_KEY from env

    def get_decision(self, prompt: str) -> str:
        response = self._client.models.generate_content(
            model=self._model,
            contents=prompt,
        )
        return response.text or ""

    def get_decision_for_media(
        self, prompt: str, image_bytes: bytes, mime_type: str = "image/png"
    ) -> str:
        image_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)
        response = self._client.models.generate_content(
            model=self._model,
            contents=[prompt, image_part],
        )
        return response.text or ""
