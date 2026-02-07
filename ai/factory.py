import os

from ai.service import AIService
from ai.openai_service import OpenAIService
from ai.gemini_service import GeminiService


def get_decision_service() -> AIService:
    """
    Return an AIService for text-only decisions.

    The provider is chosen via the AI_DECISION_PROVIDER env var:
      - "openai"  → OpenAIService   (default)
      - "gemini"  → GeminiService
    """
    provider = os.getenv("AI_DECISION_PROVIDER", "openai").lower()
    if provider == "gemini":
        return GeminiService()
    return OpenAIService()


def get_decision_for_media_service() -> AIService:
    """
    Return an AIService for image + text decisions.

    The provider is chosen via the AI_MEDIA_PROVIDER env var:
      - "openai"  → OpenAIService
      - "gemini"  → GeminiService  (default)
    """
    provider = os.getenv("AI_MEDIA_PROVIDER", "gemini").lower()
    if provider == "openai":
        return OpenAIService()
    return GeminiService()
