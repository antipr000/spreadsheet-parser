import os

from ai.service import AIService
from ai.openai_service import OpenAIService
from ai.gemini_service import GeminiService
from ai.claude_service import ClaudeService


def _make_service(provider: str) -> AIService:
    """Instantiate the appropriate AIService for a provider name."""
    provider = provider.lower().strip()
    if provider == "gemini":
        return GeminiService()
    if provider in ("claude", "anthropic"):
        return ClaudeService()
    if provider == "openai":
        return OpenAIService()
    raise ValueError(f"Unknown AI provider: {provider!r}")


def get_decision_service() -> AIService:
    """
    Return an AIService for text-only decisions.

    The provider is chosen via the AI_DECISION_PROVIDER env var:
      - "openai"     → OpenAIService
      - "gemini"     → GeminiService
      - "claude"     → ClaudeService  (default)
    """
    provider = os.getenv("AI_DECISION_PROVIDER", "claude")
    return _make_service(provider)


def get_decision_for_media_service() -> AIService:
    """
    Return an AIService for media + text decisions (images, xlsx files, etc.).

    The provider is chosen via the AI_MEDIA_PROVIDER env var:
      - "openai"     → OpenAIService
      - "gemini"     → GeminiService
      - "claude"     → ClaudeService  (default)
    """
    provider = os.getenv("AI_MEDIA_PROVIDER", "claude")
    return _make_service(provider)
