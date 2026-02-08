import base64
from openai import OpenAI

from ai.service import AIService


_DEFAULT_MODEL = "gpt-5.2"


class OpenAIService(AIService):
    """AIService backed by the OpenAI API (text-only decisions)."""

    def __init__(self, model: str = _DEFAULT_MODEL):
        self._model = model
        self._client = OpenAI()  # reads OPENAI_API_KEY from env

    def get_decision(self, prompt: str) -> str:
        response = self._client.chat.completions.create(
            model=self._model,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.choices[0].message.content or ""

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
