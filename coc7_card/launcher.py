from __future__ import annotations

import socket
import threading
import time
import urllib.error
import urllib.request
import webbrowser

import uvicorn


HOST = "127.0.0.1"
PREFERRED_PORT = 8765
PORT_SEARCH_LIMIT = 100


def _reserve_listener(
    host: str = HOST,
    preferred_port: int = PREFERRED_PORT,
    search_limit: int = PORT_SEARCH_LIMIT,
) -> tuple[int, socket.socket]:
    """Reserve the preferred local port, or the next free one without a race."""
    if not 1 <= preferred_port <= 65535:
        raise ValueError("preferred_port must be between 1 and 65535")
    if search_limit < 1:
        raise ValueError("search_limit must be at least 1")

    last_port = min(preferred_port + search_limit, 65536)
    for port in range(preferred_port, last_port):
        listener = _new_listener_socket()
        try:
            listener.bind((host, port))
            listener.listen(2048)
            listener.set_inheritable(True)
            return port, listener
        except OSError:
            listener.close()

    # If the configured range is full, let Windows assign a free ephemeral
    # port while retaining the socket for Uvicorn, which avoids a bind race.
    listener = _new_listener_socket()
    try:
        listener.bind((host, 0))
        listener.listen(2048)
        listener.set_inheritable(True)
        return int(listener.getsockname()[1]), listener
    except Exception:
        listener.close()
        raise


def _new_listener_socket() -> socket.socket:
    listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
        listener.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
    return listener


def _open_browser_when_ready(url: str) -> None:
    health_url = f"{url}/api/health"
    for _ in range(40):
        try:
            with urllib.request.urlopen(health_url, timeout=0.4) as response:  # noqa: S310
                if response.status == 200:
                    webbrowser.open(url)
                    return
        except (OSError, urllib.error.URLError):
            time.sleep(0.25)


def main() -> None:
    port, listener = _reserve_listener()
    url = f"http://{HOST}:{port}"
    if port == PREFERRED_PORT:
        print(f"COC7 Investigator Builder: {url}", flush=True)
    else:
        print(f"Port {PREFERRED_PORT} is occupied; using {url}", flush=True)

    threading.Thread(target=_open_browser_when_ready, args=(url,), daemon=True).start()
    config = uvicorn.Config("app:app", host=HOST, port=port, reload=False, access_log=False)
    server = uvicorn.Server(config)
    try:
        server.run(sockets=[listener])
    finally:
        listener.close()


if __name__ == "__main__":
    main()
