import base64
import json
import os
from typing import Any

from openai import OpenAI

MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")

SYSTEM_PROMPT = """
You are VinQu AI Assistant, an inspection assistant for technical inspection workflows.
Help inspectors understand work instructions, ITP/QCP items, inspection activities, evidence, electrical/telecom components, and report wording.
Never invent a standard clause. If a standard or acceptance criterion is not present in the provided context, say that the exact acceptance criterion must be checked in the referenced ITP/QCP/project specification.
For image findings, classify only as GREEN, YELLOW, RED, or MANUAL. If visual evidence is insufficient, say so and use MANUAL.
Return practical, inspector-focused answers.
"""


def available() -> bool:
    return bool(os.getenv("OPENAI_API_KEY"))


def _client() -> OpenAI:
    return OpenAI()


def assistant_chat(question: str, document_context: str = "{}", findings: str = "[]") -> dict[str, Any]:
    if not available():
        return {
            "connected": False,
            "answer": "VinQu AI Assistant is installed in the code, but OPENAI_API_KEY is not set on Render yet. Add the API key in Render Environment Variables to activate it.",
        }

    prompt = f"""
User question:
{question}

Current inspection document context JSON:
{document_context[:12000]}

Saved findings JSON:
{findings[:8000]}
"""
    response = _client().responses.create(
        model=MODEL,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
    )
    return {"connected": True, "answer": response.output_text}


def analyse_photo_with_ai(photo_bytes: bytes, mime_type: str, inspection_area: str, inspection_item: str, inspector_notes: str, document_context: str) -> dict[str, Any]:
    if not available():
        return {
            "status_label": "AI assistant not connected",
            "severity": "manual",
            "suggested_finding": "OPENAI_API_KEY is not set on Render",
            "inspection_area": inspection_area,
            "inspection_item": inspection_item,
            "inspector_note_suggestion": "Photo received. Add OPENAI_API_KEY in Render Environment Variables to activate automated photo analysis.",
            "reasoning": "The code is ready for OpenAI vision analysis, but the server has no API key configured yet.",
            "standard_reasoning": "No standard-based defect decision can be made until the AI assistant is activated and the applicable standards/ITP/QCP context is available.",
            "recommended_action": "Add OPENAI_API_KEY to Render, redeploy the backend, then analyse the photo again.",
        }

    data_url = f"data:{mime_type or 'image/jpeg'};base64," + base64.b64encode(photo_bytes).decode("utf-8")
    instruction = f"""
Analyse this inspection photo for the selected inspection item.

Inspection item/test: {inspection_item}
Inspection area/component: {inspection_area}
Inspector notes: {inspector_notes}

Inspection document context JSON:
{document_context[:12000]}

Return ONLY valid JSON with this schema:
{{
  "status_label": "GREEN - PASS" | "YELLOW - MINOR ISSUE" | "RED - MAJOR ISSUE" | "MANUAL REVIEW REQUIRED",
  "severity": "green" | "yellow" | "red" | "manual",
  "suggested_finding": "short finding title",
  "inspection_area": "component/area visible in the photo",
  "inspection_item": "selected inspection item",
  "inspector_note_suggestion": "editable inspector note",
  "reasoning": "what is visible and why it matters",
  "standard_reasoning": "specific standard/ITP/QCP basis if present; otherwise state exact acceptance criterion is not in context",
  "recommended_action": "next action for inspector"
}}
"""
    response = _client().responses.create(
        model=MODEL,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "input_text", "text": instruction},
                {"type": "input_image", "image_url": data_url},
            ]},
        ],
    )
    text = response.output_text.strip()
    try:
        start = text.find("{")
        end = text.rfind("}") + 1
        return json.loads(text[start:end])
    except Exception:
        return {
            "status_label": "MANUAL REVIEW REQUIRED",
            "severity": "manual",
            "suggested_finding": "AI response could not be parsed as JSON",
            "inspection_area": inspection_area,
            "inspection_item": inspection_item,
            "inspector_note_suggestion": inspector_notes or "Review photo manually and repeat analysis if needed.",
            "reasoning": text,
            "standard_reasoning": "No valid structured standard reasoning returned.",
            "recommended_action": "Review manually or retry analysis.",
        }
