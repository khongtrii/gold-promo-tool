from pathlib import Path
from collections.abc import Callable

import requests


def download_update(
    url: str,
    destination: Path,
    on_progress: Callable[[int, int], None] | None = None,
) -> Path:
    destination.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with requests.get(
        url,
        stream=True,
        timeout=120,
    ) as response:
        response.raise_for_status()

        total = int(response.headers.get("content-length", 0))
        downloaded = 0
        with destination.open("wb") as file:
            for chunk in response.iter_content(
                chunk_size=256 * 1024
            ):
                if chunk:
                    file.write(chunk)
                    downloaded += len(chunk)
                    if on_progress is not None:
                        on_progress(downloaded, total)

    return destination
