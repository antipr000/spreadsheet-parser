"""
Prompt for the Image Extractor — vision-based image description.
"""

from __future__ import annotations


def get_image_description_prompt() -> str:
    return """You are an analyst reviewing an embedded image from an Excel worksheet.

Describe this image in 2-4 sentences.  Include:
- What the image shows (logo, diagram, photograph, etc.)
- Any text visible in the image
- How it might relate to the surrounding spreadsheet context

Output ONLY the description text, no JSON or formatting.
"""
