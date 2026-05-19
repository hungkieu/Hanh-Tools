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


SYSTEM_PROMPT = (
    "You translate document text into the target language while preserving formatting markers.\n"
    "Output strictly valid JSON: {\"translations\":[{\"id\":...,\"translation\":...}]}\n"
    "Rules:\n"
    "1) Keep every marker 〈n〉...〈/n〉 byte-for-byte (do NOT use ASCII <n>; do NOT add whitespace inside).\n"
    "2) Whitespace at marker boundaries is significant — do not trim/add/collapse it.\n"
    "3) Do not output anything between 〈/n〉 and the next 〈m〉.\n"
    "4) Do not translate URLs, emails, field codes, placeholders, or symbols.\n"
    "5) Return the same ids as input. No commentary.\n"
    "Example input:  [{\"id\":\"0\",\"text\":\"〈0〉Hello 〈/0〉〈1〉world〈/1〉\"}]\n"
    "Example output: {\"translations\":[{\"id\":\"0\",\"translation\":\"〈0〉Xin chào 〈/0〉〈1〉thế giới〈/1〉\"}]}"
)

ENGLISH_ONLY_RULE = (
    "\n6) Translate ONLY English text. Keep any non-English text "
    "(Chinese, Japanese, Korean, etc.) byte-for-byte as-is in the output — "
    "do not translate or transliterate it.\n"
    "Example input:  [{\"id\":\"0\",\"text\":\"〈0〉Chapter 1 第一章〈/0〉\"}]\n"
    "Example output: {\"translations\":[{\"id\":\"0\",\"translation\":\"〈0〉Chương 1 第一章〈/0〉\"}]}"
)


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
        only_english: bool = True,
    ) -> None:
        self.client = OpenAI(timeout=timeout)
        self.model = model or os.getenv("OPENAI_TRANSLATION_MODEL", "gpt-4.1-nano")
        self.source_language = source_language
        self.target_language = target_language
        self.timeout = timeout
        self.max_completion_tokens = max_completion_tokens
        self.only_english = only_english
        self.system_prompt = SYSTEM_PROMPT + (ENGLISH_ONLY_RULE if only_english else "")
        try:
            self.encoding = tiktoken.encoding_for_model(self.model)
        except KeyError:
            self.encoding = tiktoken.get_encoding("o200k_base")

    def estimate_input_tokens(self, items: list[dict[str, str]]) -> int:
        if not items:
            return 0
        text = "".join(item.get("text", "") for item in items)
        return len(self.encoding.encode(text))

    def estimate_output_tokens(self, items: list[dict[str, str]]) -> int:
        """Ước lượng số token đầu ra cho 1 batch (đã tính overhead)."""
        input_tokens = self.estimate_input_tokens(items)
        return int(input_tokens * OUTPUT_TOKEN_MULTIPLIER) + OUTPUT_TOKEN_OVERHEAD

    def is_batch_too_large(self, items: list[dict[str, str]]) -> bool:
        return self.estimate_output_tokens(items) > self.max_completion_tokens

    def translate_batch(self, items: list[dict[str, str]]) -> dict[str, str]:
        if not items:
            return {}

        # Compact ID: thay id dài bằng integer "0","1",... để tiết kiệm token.
        compact_items = [{"id": str(i), "text": it["text"]} for i, it in enumerate(items)]
        int_to_id = {str(i): it["id"] for i, it in enumerate(items)}

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                response_format={"type": "json_object"},
                max_completion_tokens=self.max_completion_tokens,
                messages=[
                    {"role": "system", "content": self.system_prompt},
                    {
                        "role": "user",
                        "content": json.dumps(
                            {
                                "target_language": self.target_language,
                                "input": compact_items,
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
            if not isinstance(text, str):
                continue
            # Map compact id (int hoặc string số) ngược lại id gốc.
            if isinstance(item_id, int):
                item_id = str(item_id)
            if isinstance(item_id, str):
                original = int_to_id.get(item_id)
                if original is not None:
                    translations[original] = text

        return translations
