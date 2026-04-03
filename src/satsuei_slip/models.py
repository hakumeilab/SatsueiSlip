from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path


@dataclass(slots=True)
class DeliveryInfo:
    company_name: str
    project_name: str
    delivery_date: date
    recipient: str
    staff_name: str
    note: str
    episode_name: str = ""
    folder_name: str = ""
    head_trim_frames: int = 8
    sender_footer: str = ""


@dataclass(slots=True)
class VideoItem:
    file_path: Path
    width: int
    height: int
    fps: float
    frame_count: int
    duration_seconds: float
    file_size_bytes: int
    note: str = ""
    frame_count_estimated: bool = False

    @property
    def file_name(self) -> str:
        return self.file_path.name

    @property
    def resolution_text(self) -> str:
        if self.width <= 0 or self.height <= 0:
            return "-"
        return f"{self.width} x {self.height}"

    @property
    def fps_text(self) -> str:
        if self.fps <= 0:
            return "-"
        return f"{self.fps:.3f}".rstrip("0").rstrip(".")

    @property
    def frame_count_text(self) -> str:
        suffix = "*" if self.frame_count_estimated else ""
        return f"{self.frame_count}{suffix}"

    @property
    def duration_text(self) -> str:
        total_seconds = max(0, int(round(self.duration_seconds)))
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    @property
    def file_size_text(self) -> str:
        size = float(self.file_size_bytes)
        for unit in ["B", "KB", "MB", "GB", "TB"]:
            if size < 1024 or unit == "TB":
                if unit == "B":
                    return f"{int(size)} {unit}"
                return f"{size:.2f} {unit}"
            size /= 1024
        return f"{self.file_size_bytes} B"

    def trimmed_frame_count(self, head_trim_frames: int) -> int:
        return max(0, self.frame_count - max(0, head_trim_frames))

    def cut_duration_text(self, head_trim_frames: int, base_fps: int = 24) -> str:
        frames = self.trimmed_frame_count(head_trim_frames)
        seconds, remain_frames = divmod(frames, base_fps)
        return f"{seconds} + {remain_frames:02d}"
