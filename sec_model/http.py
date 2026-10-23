"""SEC EDGAR HTTP client: polite rate limiting, retries, and a 7-day disk cache."""
import hashlib
import json
import time
from pathlib import Path
from urllib.parse import urlparse

import requests

from sec_model.types import UserFacingError

CACHE_TTL_SECONDS = 7 * 24 * 3600
TIMEOUT_SECONDS = 30
BACKOFFS = (1, 2, 4)  # seconds before retry 1, 2, 3


class SecClient:
    def __init__(self, user_agent: str, cache_dir: Path = Path(".cache"), min_interval: float = 0.11):
        if "@" not in user_agent:
            raise UserFacingError("SEC User-Agent must include an email address")
        self.user_agent = user_agent
        self.cache_dir = Path(cache_dir)
        self.min_interval = min_interval
        self._last_request: float | None = None

    # -- public API ---------------------------------------------------------
    def get_bytes(self, url: str) -> bytes:
        cached = self._read_cache(url)
        if cached is not None:
            return cached
        content = self._fetch(url)
        self._write_cache(url, content)
        return content

    def get_text(self, url: str) -> str:
        return self.get_bytes(url).decode("utf-8", errors="replace")

    def get_json(self, url: str) -> dict:
        return json.loads(self.get_bytes(url))

    # -- cache --------------------------------------------------------------
    def _cache_paths(self, url: str) -> tuple[Path, Path]:
        base = self.cache_dir / hashlib.sha1(url.encode("utf-8")).hexdigest()
        return base, base.with_suffix(".meta")

    def _read_cache(self, url: str) -> bytes | None:
        body, meta = self._cache_paths(url)
        try:
            fetched_at = json.loads(meta.read_text())["fetched_at"]
            if time.time() - fetched_at > CACHE_TTL_SECONDS:
                return None
            return body.read_bytes()
        except (OSError, ValueError, KeyError):
            return None

    def _write_cache(self, url: str, content: bytes) -> None:
        body, meta = self._cache_paths(url)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        body.write_bytes(content)
        meta.write_text(json.dumps({"fetched_at": time.time()}))

    # -- network ------------------------------------------------------------
    def _throttle(self) -> None:
        if self._last_request is not None:
            wait = self.min_interval - (time.monotonic() - self._last_request)
            if wait > 0:
                time.sleep(wait)
        self._last_request = time.monotonic()

    def _fetch(self, url: str) -> bytes:
        headers = {
            "User-Agent": self.user_agent,
            "Accept-Encoding": "gzip, deflate",
            "Host": urlparse(url).netloc,
        }
        for attempt in range(len(BACKOFFS) + 1):
            self._throttle()
            try:
                resp = requests.get(url, headers=headers, timeout=TIMEOUT_SECONDS)
            except requests.RequestException:
                if attempt == len(BACKOFFS):
                    raise
            else:
                retryable = resp.status_code == 429 or resp.status_code >= 500
                if not retryable or attempt == len(BACKOFFS):
                    resp.raise_for_status()  # raises for any non-200; nothing cached
                    return resp.content
            time.sleep(BACKOFFS[attempt])
        raise AssertionError("unreachable")
