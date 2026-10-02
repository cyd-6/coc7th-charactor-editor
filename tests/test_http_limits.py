"""Body and concurrency limits apply before expensive request parsing."""

import asyncio
import json
import tempfile
from threading import Event

import pytest
import starlette.formparsers

import app as web
import coc7_card.importers


def scope(path, headers=(), method="POST"):
    return {
        "type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1",
        "scheme": "http", "method": method, "path": path, "raw_path": path.encode(),
        "root_path": "", "query_string": b"", "headers": list(headers),
        "server": ("testserver", 80), "client": ("127.0.0.1", 1234),
    }


async def request(path, chunks=(), headers=(), method="POST", receive=None):
    messages = []
    if receive is None:
        pending = iter(chunks)

        async def receive():
            try:
                return {"type": "http.request", "body": next(pending), "more_body": True}
            except StopIteration:
                return {"type": "http.request", "body": b"", "more_body": False}

    async def send(message):
        messages.append(message)

    await web.app(scope(path, headers, method), receive, send)
    start = next(message for message in messages if message["type"] == "http.response.start")
    data = b"".join(message.get("body", b"") for message in messages if message["type"] == "http.response.body")
    return start["status"], dict(start["headers"]), json.loads(data)


def assert_security_headers(headers):
    assert headers[b"cache-control"] == b"no-store"
    assert headers[b"x-content-type-options"] == b"nosniff"
    assert headers[b"x-frame-options"] == b"SAMEORIGIN"
    assert headers[b"referrer-policy"] == b"no-referrer"
    assert b"default-src 'self'" in headers[b"content-security-policy"]


async def unread_body():
    pytest.fail("Rejected requests must not read or parse the body")


@pytest.mark.parametrize("declared", [None, b"1", b"1025"])
def test_actual_json_body_size_is_limited(monkeypatch, declared):
    monkeypatch.setattr(web, "MAX_REQUEST_BYTES", 1024)
    headers = [(b"content-type", b"application/json")]
    if declared is not None:
        headers.append((b"content-length", declared))
    payload = b'{"padding":"' + b"x" * 1011 + b'"}'
    assert len(payload) == 1025
    status, response_headers, body = asyncio.run(request(
        "/api/calculate", (payload[:900], payload[900:]), headers,
        receive=unread_body if declared == b"1025" else None,
    ))
    assert status == 413
    assert body == {"detail": "请求体过大。"}
    assert_security_headers(response_headers)


def test_exact_body_limit_is_allowed(monkeypatch):
    monkeypatch.setattr(web, "MAX_REQUEST_BYTES", 1024)
    payload = b'{"padding":"' + b"x" * 1010 + b'"}'
    assert len(payload) == 1024
    status, _, body = asyncio.run(request(
        "/api/calculate", (payload[:512], payload[512:]), [(b"content-type", b"application/json")],
    ))
    assert status == 200
    assert "derived" in body


MULTIPART_HEADER = [(b"content-type", b"multipart/form-data; boundary=boundary")]
FILE_HEADER = (
    b'--boundary\r\nContent-Disposition: form-data; name="workbook"; filename="attributes.tsv"\r\n'
    b"Content-Type: text/tab-separated-values\r\n\r\n"
)
FILE_END = b"\r\n--boundary--\r\n"


@pytest.mark.parametrize("failure", ["size", "disconnect"])
def test_partial_multipart_files_close_and_admission_recovers(monkeypatch, failure):
    monkeypatch.setattr(web, "MAX_REQUEST_BYTES", 1024)
    files = []

    def track_file(*args, **kwargs):
        file = tempfile.SpooledTemporaryFile(*args, **kwargs)
        files.append(file)
        return file

    monkeypatch.setattr(starlette.formparsers, "SpooledTemporaryFile", track_file)

    async def scenario():
        chunks = iter([
            {"type": "http.request", "body": FILE_HEADER + b"STR\t60\n", "more_body": True},
            {"type": "http.disconnect"} if failure == "disconnect" else
            {"type": "http.request", "body": b"x" * 1024, "more_body": True},
        ])

        async def receive():
            return next(chunks)

        status, headers, _ = await request("/api/import/attributes", headers=MULTIPART_HEADER, receive=receive)
        assert status == (413 if failure == "size" else 400)
        assert_security_headers(headers)
        assert files and all(file.closed for file in files)
        status, _, _ = await request("/api/import/excel")
        assert status == 422

    asyncio.run(scenario())


def test_busy_request_rejected_before_body_and_parse_cancellation_releases_slot():
    async def scenario():
        reading = asyncio.Event()

        async def slow_receive():
            reading.set()
            await asyncio.Event().wait()

        first = asyncio.create_task(request("/api/import/attributes", headers=MULTIPART_HEADER, receive=slow_receive))
        try:
            await asyncio.wait_for(reading.wait(), timeout=5)
            for path in web.HEAVY_REQUEST_PATHS:
                status, headers, body = await request(path, receive=unread_body)
                assert status == 503
                assert headers[b"retry-after"] == b"1"
                assert "稍后重试" in body["detail"]
                assert_security_headers(headers)
            assert (await request("/api/health", method="GET"))[0] == 200
        finally:
            first.cancel()
            with pytest.raises(asyncio.CancelledError):
                await first
        assert (await request("/api/import/excel"))[0] == 422

    asyncio.run(scenario())


@pytest.mark.parametrize("worker_fails", [False, True])
def test_cancelled_request_holds_slot_until_worker_actually_finishes(monkeypatch, worker_fails):
    stop = Event()

    async def scenario():
        loop = asyncio.get_running_loop()
        started = asyncio.Event()
        finished = asyncio.Event()

        def import_attributes(*_args):
            loop.call_soon_threadsafe(started.set)
            try:
                assert stop.wait(timeout=5), "Test did not release worker"
                if worker_fails:
                    raise RuntimeError("synthetic worker failure")
                return {"attributes": {"STR": 60}}
            finally:
                loop.call_soon_threadsafe(finished.set)

        monkeypatch.setattr(coc7_card.importers, "import_attributes", import_attributes)
        first = asyncio.create_task(request(
            "/api/import/attributes", (FILE_HEADER + b"STR\t60\n" + FILE_END,), MULTIPART_HEADER,
        ))
        try:
            await asyncio.wait_for(started.wait(), timeout=5)
            first.cancel()
            with pytest.raises(asyncio.CancelledError):
                await first
            assert (await request("/api/export/pdf", receive=unread_body))[0] == 503
            assert (await request("/api/health", method="GET"))[0] == 200
        finally:
            stop.set()
            await asyncio.wait_for(finished.wait(), timeout=5)
        # The worker event precedes the threadpool future callback; wait for admission recovery.
        async with asyncio.timeout(5):
            while (await request("/api/import/excel"))[0] == 503:
                await asyncio.sleep(0.001)
        assert (await request("/api/import/excel"))[0] == 422

    asyncio.run(scenario())
