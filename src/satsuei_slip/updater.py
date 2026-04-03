from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from dataclasses import dataclass

from satsuei_slip import __version__
from satsuei_slip.release_config import GITHUB_OWNER, GITHUB_REPO


class UpdateCheckError(RuntimeError):
    pass


@dataclass(slots=True)
class UpdateInfo:
    current_version: str
    latest_version: str
    release_url: str
    has_update: bool


def check_github_update() -> UpdateInfo:
    if not GITHUB_OWNER or GITHUB_OWNER == "YOUR_GITHUB_USER":
        raise UpdateCheckError(
            "GitHubリポジトリ設定が未設定です。src/satsuei_slip/release_config.py の GITHUB_OWNER を変更してください。"
        )

    api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    request = urllib.request.Request(
        api_url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"SatsueiSlip/{__version__}",
        },
    )

    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise UpdateCheckError(f"GitHubの更新確認に失敗しました。HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise UpdateCheckError(f"GitHubへ接続できませんでした。{exc.reason}") from exc
    except (TimeoutError, json.JSONDecodeError) as exc:
        raise UpdateCheckError("GitHubの更新情報を正しく取得できませんでした。") from exc

    latest_version = str(payload.get("tag_name") or "").strip()
    release_url = str(payload.get("html_url") or "").strip()
    if not latest_version or not release_url:
        raise UpdateCheckError("GitHub Releases の最新リリース情報が見つかりませんでした。")

    return UpdateInfo(
        current_version=__version__,
        latest_version=latest_version,
        release_url=release_url,
        has_update=_parse_version(latest_version) > _parse_version(__version__),
    )


def _parse_version(version_text: str) -> tuple[int, ...]:
    normalized = version_text.strip().lstrip("vV")
    parts = re.findall(r"\d+", normalized)
    if not parts:
        return (0,)
    return tuple(int(part) for part in parts)
