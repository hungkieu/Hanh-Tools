from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Protocol

import tiktoken
from openai import APITimeoutError, OpenAI


# Hệ số phình token đầu ra so với đầu vào, đã cộng overhead JSON/markers.
# Dịch Anh→Việt thường phình 1.5-2x; để 2.2 + buffer cố định cho an toàn.
OUTPUT_TOKEN_MULTIPLIER = 2.2
OUTPUT_TOKEN_OVERHEAD = 200


class BatchTooLargeError(Exception):
    """Output bị cắt hoặc JSON hỏng — caller nên chia nhỏ batch và retry."""


class Translator(Protocol):
    def translate_batch(self, items: list[dict[str, str]]) -> dict[str, str]:
        ...


@dataclass
class DryRunTranslator:
    prefix: str = ""

    def translate_batch(self, items: list[dict[str, str]]) -> dict[str, str]:
        return {item["id"]: f"{self.prefix}{item['text']}" for item in items}


class OpenAITranslator:
    def __init__(
        self,
        *,
        model: str | None = None,
        source_language: str = "auto",
        target_language: str = "Vietnamese",
        timeout: float = 60,
        max_completion_tokens: int = 8192,
    ) -> None:
        self.client = OpenAI(timeout=timeout)
        self.model = model or os.getenv("OPENAI_TRANSLATION_MODEL", "gpt-4.1-nano")
        self.source_language = source_language
        self.target_language = target_language
        self.timeout = timeout
        self.max_completion_tokens = max_completion_tokens
        try:
            self.encoding = tiktoken.encoding_for_model(self.model)
        except KeyError:
            self.encoding = tiktoken.get_encoding("o200k_base")

    def estimate_output_tokens(self, items: list[dict[str, str]]) -> int:
        """Ước lượng số token đầu ra cho 1 batch (đã tính overhead)."""
        if not items:
            return 0
        text = "".join(item.get("text", "") for item in items)
        input_tokens = len(self.encoding.encode(text))
        return int(input_tokens * OUTPUT_TOKEN_MULTIPLIER) + OUTPUT_TOKEN_OVERHEAD

    def is_batch_too_large(self, items: list[dict[str, str]]) -> bool:
        return self.estimate_output_tokens(items) > self.max_completion_tokens

    def translate_batch(self, items: list[dict[str, str]]) -> dict[str, str]:
        if not items:
            return {}

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                response_format={"type": "json_object"},
                max_completion_tokens=self.max_completion_tokens,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You translate document text while preserving formatting markers. "
                            "Return only valid JSON. Preserve every marker exactly, including "
                            "markers like [[[0]]] and [[[/0]]]. Do not add commentary. "
                            "Whitespace inside markers is significant — never trim, add, or "
                            "collapse it. Never insert text between [[[/n]]] and the next "
                            "[[[m]]] marker."
                        ),
                    },
                    {
                        "role": "user",
                        "content": json.dumps(
                            {
                                "task": "translate",
                                "source_language": self.source_language,
                                "target_language": self.target_language,
                                "rules": [
                                    "Translate natural-language text only.",
                                    "Keep all markers exactly unchanged (same count, same ids, same order).",
                                    "Return the same ids in the JSON output.",
                                    "Do not translate URLs, emails, field codes, or placeholders.",
                                    "Preserve whitespace at the boundaries of each marker exactly. "
                                    "If source slot starts/ends with a space, the translated slot must too.",
                                    "Do not output anything between a closing marker [[[/n]]] and the "
                                    "next opening marker [[[m]]] — they must be adjacent.",
                                    "If a slot contains only whitespace or punctuation, copy it verbatim.",
                                ],
                                "example": {
                                    "input": [{"id": "x", "text": "[[[0]]]Hello [[[/0]]][[[1]]]world[[[/1]]]"}],
                                    "output": {"translations": [
                                        {"id": "x", "translation": "[[[0]]]Xin chào [[[/0]]][[[1]]]thế giới[[[/1]]]"}
                                    ]},
                                },
                                "input": items,
                                "output_schema": {
                                    "translations": [
                                        {
                                            "id": "same id as input",
                                            "translation": "translated text",
                                        }
                                    ]
                                },
                            },
                            ensure_ascii=False,
                        ),
                    },
                ],
            )
        except APITimeoutError as exc:
            raise TimeoutError(
                f"OpenAI request timed out after {self.timeout}s for {len(items)} paragraphs"
            ) from exc
        choice = response.choices[0]
        content = choice.message.content
        finish_reason = getattr(choice, "finish_reason", None)
        if finish_reason == "length":
            raise BatchTooLargeError(
                f"Output bị cắt do max_completion_tokens (batch={len(items)} items)"
            )
        if not content:
            raise ValueError("OpenAI trả về response rỗng")

        try:
            parsed = json.loads(content)
        except json.JSONDecodeError as exc:
            raise BatchTooLargeError(
                f"OpenAI trả về JSON hỏng (batch={len(items)} items): {exc}"
            ) from exc
        if isinstance(parsed, list):
            rows = parsed
        elif isinstance(parsed, dict):
            rows = parsed.get("translations") or parsed.get("output") or parsed.get("items")
        else:
            rows = None
        if not isinstance(rows, list):
            raise ValueError("OpenAI response không đúng schema JSON mong đợi")

        translations: dict[str, str] = {}
        for row in rows:
            if not isinstance(row, dict):
                continue
            item_id = row.get("id")
            text = row.get("translation")
            if isinstance(item_id, str) and isinstance(text, str):
                translations[item_id] = text

        return translations
