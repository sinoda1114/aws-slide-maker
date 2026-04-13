"""Convert natural-language prompts or YAML strings into PresentationSpec via Claude API."""
from __future__ import annotations

import os
import re

import anthropic
import yaml

from .schema import PresentationSpec

_SYSTEM_PROMPT = """\
You are a professional presentation designer.
Given a topic or description, produce a slide specification in **YAML only** — no explanation, no markdown fences.

Rules:
- 6–12 slides total
- Always start with a `title` slide and end with a `closing` slide
- Include an `agenda` slide as slide 2
- Use `content` slides for main body (bullets preferred, max 5 bullets each)
- Use `two_column` slides to compare/contrast two topics
- Write in the same language as the user's input (Japanese → Japanese, English → English)
- Be concise and professional

YAML schema:
```
title: "Presentation title"
subtitle: "Optional subtitle"
author: ""
theme: aws          # aws | minimal | dark
slides:
  - type: title
    title: "..."
    subtitle: "..."      # optional

  - type: agenda
    title: "アジェンダ"  # or "Agenda"
    items:
      - "Item 1"

  - type: content
    title: "Slide title"
    body: "Optional lead sentence"   # optional
    bullets:
      - "Bullet point"

  - type: two_column
    title: "Comparison"
    left:
      title: "Left heading"          # optional
      bullets:
        - "..."
    right:
      title: "Right heading"         # optional
      bullets:
        - "..."

  - type: closing
    title: "ご清聴ありがとうございました"
    message: "Optional message"      # optional
```

Return ONLY the YAML content, nothing else.
"""


def _is_yaml_spec(text: str) -> bool:
    """Heuristic: does the string look like a YAML slide spec?"""
    stripped = text.strip()
    return stripped.startswith("title:") or "type: title" in stripped


def _clean_yaml(raw: str) -> str:
    """Strip markdown code fences if the model accidentally added them."""
    raw = raw.strip()
    raw = re.sub(r"^```(?:yaml)?\s*\n", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"\n```\s*$", "", raw)
    return raw.strip()


def spec_from_yaml(yaml_text: str) -> PresentationSpec:
    """Parse raw YAML into a PresentationSpec."""
    data = yaml.safe_load(_clean_yaml(yaml_text))
    return PresentationSpec.model_validate(data)


def spec_from_description(description: str, model: str = "claude-sonnet-4-6") -> PresentationSpec:
    """Call Claude API to turn a free-text description into a PresentationSpec."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "ANTHROPIC_API_KEY environment variable is not set. "
            "Export it before running the slide-maker MCP server."
        )

    client = anthropic.Anthropic(api_key=api_key)

    message = client.messages.create(
        model=model,
        max_tokens=4096,
        system=_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": description}],
    )

    raw_yaml = message.content[0].text
    return spec_from_yaml(raw_yaml)


def parse_input(text: str, model: str = "claude-sonnet-4-6") -> PresentationSpec:
    """
    Accept either:
    - A free-text description  → call Claude to generate spec
    - A YAML slide spec string → parse directly
    """
    if _is_yaml_spec(text):
        return spec_from_yaml(text)
    return spec_from_description(text, model=model)
