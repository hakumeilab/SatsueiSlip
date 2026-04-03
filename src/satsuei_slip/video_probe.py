from __future__ import annotations

import json
import math
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterator

from satsuei_slip.models import VideoItem

SUPPORTED_EXTENSIONS = {".mp4", ".mov", ".avi", ".mxf", ".mkv"}


class VideoProbeError(RuntimeError):
    pass


def find_ffprobe_executable() -> str | None:
    candidates: list[Path] = []

    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
        candidates.extend(
            [
                base_dir / "tools" / "ffprobe" / "ffprobe.exe",
                base_dir / "tools" / "ffprobe" / "bin" / "ffprobe.exe",
                base_dir / "ffprobe.exe",
            ]
        )

    project_root = Path(__file__).resolve().parents[2]
    candidates.extend(
        [
            project_root / "tools" / "ffprobe" / "ffprobe.exe",
            project_root / "tools" / "ffprobe" / "bin" / "ffprobe.exe",
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)

    return shutil.which("ffprobe")


def collect_video_files(paths: list[Path]) -> list[Path]:
    return sorted(iter_video_files(paths), key=lambda item: item.name.lower())


def iter_video_files(paths: list[Path]) -> Iterator[Path]:
    video_files: set[Path] = set()
    for path in paths:
        resolved = path.expanduser().resolve()
        if resolved.is_file() and resolved.suffix.lower() in SUPPORTED_EXTENSIONS:
            if resolved not in video_files:
                video_files.add(resolved)
                yield resolved
        elif resolved.is_dir():
            stack = [resolved]
            while stack:
                current_dir = stack.pop()
                try:
                    with os.scandir(current_dir) as entries:
                        entry_list = sorted(entries, key=lambda entry: entry.name.lower(), reverse=True)
                except OSError:
                    continue

                for entry in entry_list:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(Path(entry.path))
                        elif entry.is_file(follow_symlinks=False):
                            child_path = Path(entry.path)
                            if child_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                                continue
                            child_resolved = child_path.resolve()
                            if child_resolved not in video_files:
                                video_files.add(child_resolved)
                                yield child_resolved
                    except OSError:
                        continue


def _parse_fps(rate_text: str | None) -> float:
    if not rate_text or rate_text == "0/0":
        return 0.0
    if "/" in rate_text:
        numerator_text, denominator_text = rate_text.split("/", 1)
        numerator = float(numerator_text)
        denominator = float(denominator_text)
        if denominator == 0:
            return 0.0
        return numerator / denominator
    return float(rate_text)


def _parse_number(value: object) -> float:
    if value is None or value == "N/A":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _parse_int(value: object) -> int:
    if value is None or value == "N/A":
        return 0
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


class FFprobeVideoAnalyzer:
    def __init__(self, ffprobe_path: str | None = None) -> None:
        self.ffprobe_path = ffprobe_path or find_ffprobe_executable()
        if not self.ffprobe_path:
            raise VideoProbeError(
                "ffprobe が見つかりません。PATH に追加するか tools\\ffprobe 配下へ配置してください。"
            )

    def analyze(self, file_path: Path) -> VideoItem:
        return self._analyze(file_path, count_frames=True)

    def analyze_fast(self, file_path: Path) -> VideoItem:
        return self._analyze(file_path, count_frames=False)

    def _analyze(self, file_path: Path, count_frames: bool) -> VideoItem:
        ffprobe_input_path = _to_ffprobe_input_path(file_path)
        command = [
            self.ffprobe_path,
            "-v",
            "error",
            "-select_streams",
            "v:0",
            "-show_entries",
            (
                "stream=width,height,avg_frame_rate,r_frame_rate,nb_frames,"
                "nb_read_frames,duration:format=duration,size"
            ),
            "-of",
            "json",
            ffprobe_input_path,
        ]
        if count_frames:
            command.insert(3, "-count_frames")

        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )
        except OSError as exc:
            raise VideoProbeError(f"ffprobe の実行に失敗しました: {file_path}") from exc

        if completed.returncode != 0:
            message = completed.stderr.strip() or "動画情報の取得に失敗しました。"
            raise VideoProbeError(f"{file_path.name}: {message}")

        try:
            payload = json.loads(completed.stdout)
        except json.JSONDecodeError as exc:
            raise VideoProbeError(f"{file_path.name}: ffprobe の出力解析に失敗しました。") from exc

        stream = (payload.get("streams") or [{}])[0]
        format_info = payload.get("format") or {}

        width = _parse_int(stream.get("width"))
        height = _parse_int(stream.get("height"))
        fps = _parse_fps(stream.get("avg_frame_rate")) or _parse_fps(stream.get("r_frame_rate"))
        duration_seconds = _parse_number(stream.get("duration")) or _parse_number(format_info.get("duration"))
        file_size_bytes = _parse_int(format_info.get("size")) or file_path.stat().st_size

        frame_count = _parse_int(stream.get("nb_read_frames")) or _parse_int(stream.get("nb_frames"))
        frame_count_estimated = False
        if frame_count <= 0 and duration_seconds > 0 and fps > 0:
            frame_count = int(round(duration_seconds * fps))
            frame_count_estimated = True

        if frame_count < 0:
            frame_count = 0
        if math.isnan(fps) or math.isinf(fps):
            fps = 0.0

        return VideoItem(
            file_path=file_path.resolve(),
            width=width,
            height=height,
            fps=fps,
            frame_count=frame_count,
            duration_seconds=duration_seconds,
            file_size_bytes=file_size_bytes,
            frame_count_estimated=frame_count_estimated,
        )


def _to_ffprobe_input_path(file_path: Path) -> str:
    resolved = str(file_path.resolve())
    if os.name != "nt" or resolved.isascii():
        return resolved

    try:
        import ctypes
        from ctypes import wintypes

        buffer = ctypes.create_unicode_buffer(32768)
        result = ctypes.windll.kernel32.GetShortPathNameW(
            wintypes.LPCWSTR(resolved),
            buffer,
            wintypes.DWORD(len(buffer)),
        )
        if result:
            return buffer.value
    except Exception:
        pass

    return resolved
