import os

from typing import Literal

DETECTION_TYPE: Literal["heuristic", "ai", "heuristic_then_ai"] = os.getenv(
    "DETECTION_TYPE", "heuristic"
).lower()
