import json
import time

import pytest
import requests

from sec_model.http import SecClient
from sec_model.types import UserFacingError


class FakeResp:
    def __init__(self, status_code, content=b""):
        self.status_code = status_code
        self.content = content
        self.text = content.decode("utf-8", errors="replace")

    def json(self):
        return json.loads(self.text)

    def raise_for_status(self):
        if self.status_code != 200:
            raise requests.HTTPError(f"{self.status_code}")


def test_rejects_user_agent_without_email(tmp_path):
    with pytest.raises(UserFacingError):
        SecClient("no email here", cache_dir=tmp_path)


def test_second_call_hits_cache(tmp_path, monkeypatch):
    calls = []

    def fake_get(url, headers, timeout):
        calls.append(url)
        return FakeResp(200, b'{"a":1}')

    monkeypatch.setattr("sec_model.http.requests.get", fake_get)
    c = SecClient("Test test@example.com", cache_dir=tmp_path)
    assert c.get_json("https://data.sec.gov/x.json") == {"a": 1}
    assert c.get_json("https://data.sec.gov/x.json") == {"a": 1}
    assert len(calls) == 1


def test_host_header_matches_url(tmp_path, monkeypatch):
    seen = {}
    monkeypatch.setattr(
        "sec_model.http.requests.get",
        lambda url, headers, timeout: seen.update(headers) or FakeResp(200, b"ok"),
    )
    SecClient("Test test@example.com", cache_dir=tmp_path).get_text("https://www.sec.gov/y.htm")
    assert seen["Host"] == "www.sec.gov" and "@" in seen["User-Agent"]
    assert seen["Accept-Encoding"] == "gzip, deflate"


def test_retries_then_raises(tmp_path, monkeypatch):
    sleeps = []
    attempts = []
    monkeypatch.setattr("sec_model.http.time.sleep", lambda s: sleeps.append(s))

    def boom(*a, **k):
        attempts.append(1)
        raise requests.ConnectionError()

    monkeypatch.setattr("sec_model.http.requests.get", boom)
    with pytest.raises(requests.ConnectionError):
        SecClient("Test test@example.com", cache_dir=tmp_path).get_text("https://www.sec.gov/z")
    assert len(attempts) == 4  # initial try + 3 retries
    assert [s for s in sleeps if s >= 1] == [1, 2, 4]


def test_non_200_not_cached_and_raises(tmp_path, monkeypatch):
    calls = []

    def fake_get(url, headers, timeout):
        calls.append(url)
        return FakeResp(404, b"nope")

    monkeypatch.setattr("sec_model.http.requests.get", fake_get)
    c = SecClient("Test test@example.com", cache_dir=tmp_path)
    for _ in range(2):
        with pytest.raises(requests.HTTPError):
            c.get_text("https://www.sec.gov/missing")
    assert len(calls) == 2  # 404 is not retried and not cached
    assert list(tmp_path.iterdir()) == []


def test_expired_cache_is_refetched(tmp_path, monkeypatch):
    calls = []

    def fake_get(url, headers, timeout):
        calls.append(url)
        return FakeResp(200, b"fresh")

    monkeypatch.setattr("sec_model.http.requests.get", fake_get)
    c = SecClient("Test test@example.com", cache_dir=tmp_path)
    url = "https://www.sec.gov/old"
    assert c.get_bytes(url) == b"fresh"
    meta = next(p for p in tmp_path.iterdir() if p.suffix == ".meta")
    meta.write_text(json.dumps({"fetched_at": time.time() - 8 * 86400}))
    assert c.get_bytes(url) == b"fresh"
    assert len(calls) == 2


def test_min_interval_between_requests(tmp_path, monkeypatch):
    sleeps = []
    monkeypatch.setattr("sec_model.http.time.sleep", lambda s: sleeps.append(s))
    monkeypatch.setattr(
        "sec_model.http.requests.get", lambda url, headers, timeout: FakeResp(200, b"ok")
    )
    c = SecClient("Test test@example.com", cache_dir=tmp_path, min_interval=5.0)
    c.get_text("https://www.sec.gov/a")
    c.get_text("https://www.sec.gov/b")
    assert len(sleeps) == 1 and 0 < sleeps[0] <= 5.0


def test_503_then_200_is_retried(tmp_path, monkeypatch):
    monkeypatch.setattr("sec_model.http.time.sleep", lambda s: None)
    responses = iter([FakeResp(503, b"busy"), FakeResp(200, b"ok")])
    calls = []

    def fake_get(url, headers, timeout):
        calls.append(url)
        return next(responses)

    monkeypatch.setattr("sec_model.http.requests.get", fake_get)
    c = SecClient("Test test@example.com", cache_dir=tmp_path)
    assert c.get_bytes("https://www.sec.gov/flaky") == b"ok"
    assert len(calls) == 2


def test_persistent_503_raises_after_retries(tmp_path, monkeypatch):
    sleeps = []
    calls = []
    monkeypatch.setattr("sec_model.http.time.sleep", lambda s: sleeps.append(s))

    def fake_get(url, headers, timeout):
        calls.append(url)
        return FakeResp(503, b"busy")

    monkeypatch.setattr("sec_model.http.requests.get", fake_get)
    c = SecClient("Test test@example.com", cache_dir=tmp_path)
    with pytest.raises(requests.HTTPError):
        c.get_bytes("https://www.sec.gov/down")
    assert len(calls) == 4  # initial try + 3 retries
    assert [s for s in sleeps if s >= 1] == [1, 2, 4]
    assert list(tmp_path.iterdir()) == []
