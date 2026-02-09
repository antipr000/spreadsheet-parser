"""
Agentic workflow for structure-aware Excel parsing.

Two-phase pipeline:
  1. PlannerAgent  — multimodal LLM identifies blocks (type + bbox + hints)
  2. Orchestrator  — dispatches each block to a specialised extractor
"""
