from abc import ABC, abstractmethod


class AIService(ABC):
    """
    Base class for AI-powered decision services.

    Subclasses must implement either or both of:
      - get_decision        (text-only prompt → text response)
      - get_decision_for_media  (image + prompt → text response)
    """

    @abstractmethod
    def get_decision(self, prompt: str) -> str:
        """Send a text prompt to the LLM and return its response."""
        ...

    @abstractmethod
    def get_decision_for_media(self, prompt: str, image_bytes: bytes, mime_type: str = "image/png") -> str:
        """Send an image together with a text prompt and return the LLM response."""
        ...
