"""Presentation session management and path sandboxing for Office PowerPoint MCP Server."""

import json
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pptx
from pptx import Presentation
from pptx.util import Inches


@dataclass
class PresentationRecord:
    presentation_id: str
    presentation: pptx.presentation.Presentation
    file_path: Path | None = None
    auto_save: bool = False
    created_at: float = field(default_factory=time.time)
    modified_at: float = field(default_factory=time.time)

    def mark_modified(self) -> None:
        self.modified_at = time.time()


class SessionManager:
    """Manages PowerPoint presentation instances, hybrid in-memory/disk state,

    and strict directory sandboxing.
    """

    def __init__(self, allowed_dirs: list[str | Path] | None = None):
        if allowed_dirs:
            self.allowed_dirs = [Path(d).expanduser().resolve() for d in allowed_dirs]
        else:
            self.allowed_dirs = [Path.cwd().resolve()]

        self.presentations: dict[str, PresentationRecord] = {}
        self.active_presentation_id: str | None = None
        self._templates_cache: dict[str, Any] | None = None

    def add_allowed_dir(self, directory: str | Path) -> None:
        """Add an allowed directory to the sandbox."""
        resolved = Path(directory).expanduser().resolve()
        if resolved not in self.allowed_dirs:
            self.allowed_dirs.append(resolved)

    def set_allowed_dirs(self, directories: list[str | Path]) -> None:
        """Replace the allowed sandbox directories."""
        self.allowed_dirs = [Path(d).expanduser().resolve() for d in directories]

    def validate_path(
        self,
        file_path: str | Path,
        allow_nonexistent: bool = True,
        must_be_file: bool = True,
    ) -> Path:
        """Validate that a path is securely within the allowed sandbox directories.

        Guards against path traversal ('..') and unauthorized directory access.
        """
        raw_path = Path(file_path).expanduser()
        resolved_path = raw_path.resolve()

        # Check if the resolved path is within any allowed directory
        is_allowed = any(
            resolved_path == allowed_dir or allowed_dir in resolved_path.parents for allowed_dir in self.allowed_dirs
        )

        if not is_allowed:
            allowed_str = ", ".join(str(d) for d in self.allowed_dirs)
            raise ValueError(
                f"Access denied: path '{file_path}' resolves to '{resolved_path}', "
                f"which is outside allowed sandbox directories: [{allowed_str}]"
            )

        if not allow_nonexistent and not resolved_path.exists():
            raise FileNotFoundError(f"File not found: {resolved_path}")

        if must_be_file and resolved_path.exists() and not resolved_path.is_file():
            raise ValueError(f"Path is not a regular file: {resolved_path}")

        return resolved_path

    def load_templates_data(self) -> dict[str, Any]:
        """Load templates configuration from templates.json."""
        if self._templates_cache is not None:
            return self._templates_cache

        templates_path = Path(__file__).parent / "templates.json"
        if not templates_path.exists():
            raise FileNotFoundError(f"templates.json not found at {templates_path}")

        with open(templates_path, "r", encoding="utf-8") as f:
            self._templates_cache = json.load(f)

        return self._templates_cache

    def create_presentation(
        self,
        presentation_id: str | None = None,
        file_path: str | Path | None = None,
        width_inches: float = 13.333,
        height_inches: float = 7.5,
        auto_save: bool = False,
    ) -> PresentationRecord:
        """Create a new widescreen 16:9 presentation."""
        prs = Presentation()
        prs.slide_width = Inches(width_inches)
        prs.slide_height = Inches(height_inches)

        target_path: Path | None = None
        if file_path:
            target_path = self.validate_path(file_path, allow_nonexistent=True)
            if auto_save:
                target_path.parent.mkdir(parents=True, exist_ok=True)
                prs.save(str(target_path))

        pid = presentation_id or (f"pres_{target_path.stem}" if target_path else f"pres_{uuid.uuid4().hex[:8]}")

        record = PresentationRecord(
            presentation_id=pid,
            presentation=prs,
            file_path=target_path,
            auto_save=auto_save,
        )
        self.presentations[pid] = record
        self.active_presentation_id = pid
        return record

    def open_presentation(
        self,
        file_path: str | Path,
        presentation_id: str | None = None,
        auto_save: bool = True,
    ) -> PresentationRecord:
        """Open an existing PowerPoint presentation from disk within allowed sandbox."""
        resolved = self.validate_path(file_path, allow_nonexistent=False, must_be_file=True)

        # Check if already open
        for pid, rec in self.presentations.items():
            if rec.file_path and rec.file_path.resolve() == resolved:
                rec.auto_save = auto_save
                self.active_presentation_id = pid
                return rec

        prs = Presentation(str(resolved))
        pid = presentation_id or f"pres_{resolved.stem}_{uuid.uuid4().hex[:6]}"
        record = PresentationRecord(
            presentation_id=pid,
            presentation=prs,
            file_path=resolved,
            auto_save=auto_save,
        )
        self.presentations[pid] = record
        self.active_presentation_id = pid
        return record

    def get_or_load_presentation(
        self,
        presentation_id: str | None = None,
        file_path: str | Path | None = None,
        auto_save: bool = True,
    ) -> PresentationRecord:
        """Resolve a presentation record by ID, file path, or active session."""
        if file_path:
            resolved = self.validate_path(file_path, allow_nonexistent=True)
            # Find in open presentations
            for pid, rec in self.presentations.items():
                if rec.file_path and rec.file_path.resolve() == resolved:
                    self.active_presentation_id = pid
                    return rec

            if resolved.exists():
                return self.open_presentation(resolved, presentation_id, auto_save=auto_save)
            else:
                # Create new presentation at this path
                return self.create_presentation(
                    presentation_id=presentation_id,
                    file_path=resolved,
                    auto_save=auto_save,
                )

        if presentation_id:
            if presentation_id in self.presentations:
                self.active_presentation_id = presentation_id
                return self.presentations[presentation_id]
            raise ValueError(
                f"Presentation '{presentation_id}' not found. Open presentations: {list(self.presentations.keys())}"
            )

        if self.active_presentation_id and self.active_presentation_id in self.presentations:
            return self.presentations[self.active_presentation_id]

        # If nothing exists yet, create a default widescreen presentation
        return self.create_presentation(presentation_id="default_presentation")

    def save_record(
        self,
        record: PresentationRecord,
        target_path: str | Path | None = None,
    ) -> Path:
        """Save a presentation to disk."""
        path_to_use = target_path or record.file_path
        if not path_to_use:
            raise ValueError(f"No file path specified to save presentation '{record.presentation_id}'")

        resolved = self.validate_path(path_to_use, allow_nonexistent=True)
        resolved.parent.mkdir(parents=True, exist_ok=True)
        record.presentation.save(str(resolved))
        record.file_path = resolved
        record.mark_modified()
        return resolved

    def auto_save_if_needed(self, record: PresentationRecord) -> None:
        """Automatically persist to disk if record has auto_save enabled and a file_path."""
        record.mark_modified()
        if record.auto_save and record.file_path:
            self.save_record(record)

    def close_presentation(self, presentation_id: str) -> bool:
        """Close presentation and remove from in-memory registry."""
        if presentation_id in self.presentations:
            del self.presentations[presentation_id]
            if self.active_presentation_id == presentation_id:
                self.active_presentation_id = next(iter(self.presentations.keys())) if self.presentations else None
            return True
        return False


# Global session manager instance
session_manager = SessionManager()
