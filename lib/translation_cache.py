from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from pathlib import Path


# Tăng số này khi prompt/marker format đổi → cache cũ sẽ bị bỏ qua.
CACHE_KEY_VERSION = 1


def cache_dir() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    folder = base / "HanhTools"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def cache_file() -> Path:
    return cache_dir() / f"cache_v{CACHE_KEY_VERSION}.json"


def open_cache_folder() -> None:
    folder = cache_dir()
    if sys.platform == "win32":
        os.startfile(str(folder))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(folder)])
    else:
        subprocess.Popen(["xdg-open", str(folder)])


class TranslationCache:
    """Cache persistent (text → translation) theo (model, target_lang).
    Bỏ qua text quá ngắn để tránh dính nhầm vùng mơ hồ ngữ nghĩa."""

    MIN_PLAIN_LEN = 15  # plain text (đã strip marker) phải dài tối thiểu chừng này

    def __init__(self, path: Path | None = None, *, enabled: bool = True) -> None:
        self.path = path or cache_file()
        self.enabled = enabled
        self._lock = threading.Lock()
        self._data: dict[str, dict[str, str]] = {}
        self._dirty = False
        self.hits = 0
        self.misses = 0
        if enabled:
            self._load()

    def _load(self) -> None:
        if not self.path.exists():
            return
        try:
            raw = self.path.read_text(encoding="utf-8")
            data = json.loads(raw)
            if isinstance(data, dict):
                self._data = data
        except Exception:
            self._data = {}

    def save(self) -> None:
        if not self.enabled or not self._dirty:
            return
        with self._lock:
            tmp = self.path.with_suffix(self.path.suffix + ".tmp")
            tmp.write_text(json.dumps(self._data, ensure_ascii=False), encoding="utf-8")
            tmp.replace(self.path)
            self._dirty = False

    def clear(self) -> None:
        with self._lock:
            self._data = {}
            self._dirty = False
            if self.path.exists():
                self.path.unlink()

    def _bucket_key(self, model: str, target_lang: str) -> str:
        return f"{model}|{target_lang}"

    def get(self, model: str, target_lang: str, text: str) -> str | None:
        if not self.enabled:
            return None
        bucket = self._data.get(self._bucket_key(model, target_lang))
        if bucket is None:
            self.misses += 1
            return None
        result = bucket.get(text)
        if result is None:
            self.misses += 1
        else:
            self.hits += 1
        return result

    def put(self, model: str, target_lang: str, text: str, translation: str, plain_len: int) -> None:
        if not self.enabled:
            return
        if plain_len < self.MIN_PLAIN_LEN:
            return
        with self._lock:
            self._data.setdefault(self._bucket_key(model, target_lang), {})[text] = translation
            self._dirty = True

    def entry_count(self) -> int:
        return sum(len(b) for b in self._data.values())

    def file_size_kb(self) -> float:
        try:
            return self.path.stat().st_size / 1024.0
        except FileNotFoundError:
            return 0.0
