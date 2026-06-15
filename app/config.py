from __future__ import annotations

import os
import platform
import sys
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class Settings:
    base_dir: Path
    static_dir: Path
    runs_dir: Path
    platform_name: str
    downloads_dir: Path = field(default_factory=lambda: Path.home() / "Downloads")
    cleanup_after_hours: int = 24
    render_dpi: int = 240
    enable_powerpoint_fallback: bool = False
    powerpoint_app_path: Optional[Path] = None
    libreoffice_bin: Optional[Path] = None
    powershell_bin: str = "powershell.exe"
    google_document_ai_project_id: Optional[str] = None
    google_document_ai_location: Optional[str] = None
    google_document_ai_processor_id: Optional[str] = None
    google_document_ai_processor_version: Optional[str] = None
    enable_ai_qc: bool = True
    openai_api_key: Optional[str] = None
    openai_qc_model: str = "gpt-5.3-chat-latest"
    openai_qc_parallelism: int = 4
    openai_qc_timeout_seconds: float = 90.0
    openai_qc_max_image_dimension: int = 0
    prefer_powerpoint_for_ai_qc: bool = True
    prefer_poppler_for_ai_qc: bool = True
    poppler_bin_dir: Optional[Path] = None
    powerpoint_slide_export_macro_name: Optional[str] = None
    powerpoint_reference_slide_insert_macro_name: Optional[str] = None
    powerpoint_slide_export_staging_dir: Optional[Path] = None
    powerpoint_slide_export_long_edge: int = 2800


def _parse_env_value(raw_value: str) -> str:
    value = raw_value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        return value[1:-1]
    if " #" in value:
        return value.split(" #", 1)[0].rstrip()
    return value


def _apply_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped.startswith("export "):
            stripped = stripped[len("export ") :].strip()
        if "=" not in stripped:
            continue
        name, raw_value = stripped.split("=", 1)
        name = name.strip()
        if not name:
            continue
        os.environ[name] = _parse_env_value(raw_value)


def _detect_platform_name() -> str:
    override = os.getenv("PDF_PPTX_PLATFORM", "").strip().lower()
    raw_platform = override or platform.system().lower()
    aliases = {
        "mac": "darwin",
        "macos": "darwin",
        "osx": "darwin",
        "win": "windows",
        "win32": "windows",
    }
    return aliases.get(raw_platform, raw_platform)


def _default_powerpoint_app_path(platform_name: str) -> Optional[Path]:
    if platform_name == "darwin":
        return Path("/Applications/Microsoft PowerPoint.app")
    return None


def _default_libreoffice_bin(platform_name: str) -> Optional[Path]:
    if platform_name == "darwin":
        return Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if platform_name == "windows":
        windows_roots = [os.getenv("PROGRAMFILES"), os.getenv("PROGRAMFILES(X86)")]
        for root in windows_roots:
            if root:
                return Path(root) / "LibreOffice" / "program" / "soffice.exe"
        return Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
    return None


def _path_from_env(name: str, fallback: Optional[Path]) -> Optional[Path]:
    value = os.getenv(name, "").strip()
    if value:
        return Path(value)
    return fallback


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _resource_root(default_app_dir: Path) -> Path:
    if _is_frozen():
        return Path(getattr(sys, "_MEIPASS"))
    return default_app_dir.parent


def _project_env_roots(
    default_app_dir: Path,
    *,
    frozen: Optional[bool] = None,
    executable_path: Optional[Path] = None,
    resource_root: Optional[Path] = None,
) -> tuple[Path, ...]:
    resolved_frozen = _is_frozen() if frozen is None else frozen
    runtime_root = resource_root or _resource_root(default_app_dir)
    roots: list[Path] = [runtime_root]
    if not resolved_frozen:
        return tuple(roots)

    executable_dir = Path(executable_path or sys.executable).resolve().parent
    roots.append(executable_dir)
    if executable_dir.name == "MacOS":
        executable_parents = executable_dir.parents
        if len(executable_parents) >= 4 and executable_parents[1].suffix.lower() == ".app":
            roots.append(executable_parents[2])

    deduped: list[Path] = []
    seen: set[str] = set()
    for root in roots:
        key = str(root.resolve(strict=False))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(root)
    return tuple(deduped)


def _load_project_env_files() -> None:
    app_dir = Path(__file__).resolve().parent
    for base_dir in _project_env_roots(app_dir):
        for filename in (".env", ".env.local"):
            _apply_env_file(base_dir / filename)


def _default_runtime_root(platform_name: str) -> Path:
    if platform_name == "windows":
        local_app_data = os.getenv("LOCALAPPDATA")
        base = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
        return base / "PDF to PPTX Reference Placement"
    if platform_name == "darwin":
        return Path.home() / "Library" / "Application Support" / "PDF to PPTX Reference Placement"
    return Path.home() / ".pdf-to-pptx-reference-placement"


def _default_downloads_dir(platform_name: str) -> Path:
    override = os.getenv("PDF_PPTX_DOWNLOADS_DIR", "").strip()
    if override:
        return Path(override)
    if platform_name == "windows":
        user_profile = os.getenv("USERPROFILE", "").strip()
        if user_profile:
            return Path(user_profile) / "Downloads"
    return Path.home() / "Downloads"


def _default_powerpoint_slide_export_staging_dir(platform_name: str) -> Optional[Path]:
    if platform_name == "darwin":
        return (
            Path.home()
            / "Library"
            / "Containers"
            / "com.microsoft.Powerpoint"
            / "Data"
            / "Documents"
            / "codex-png-out"
        )
    return None


def _runs_dir(platform_name: str) -> Path:
    override = os.getenv("PDF_PPTX_RUNS_DIR", "").strip()
    if override:
        return Path(override)
    return _default_runtime_root(platform_name) / "runs"


def _default_enable_powerpoint_fallback(platform_name: str) -> bool:
    value = os.getenv("PDF_PPTX_ENABLE_POWERPOINT_FALLBACK", "").strip().lower()
    if value:
        return value in {"1", "true", "yes", "on"}
    return platform_name == "windows"


def _env_flag(name: str, default: bool) -> bool:
    value = os.getenv(name, "").strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on"}


_load_project_env_files()


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    app_dir = Path(__file__).resolve().parent
    platform_name = _detect_platform_name()
    base_dir = _resource_root(app_dir)
    return Settings(
        base_dir=base_dir,
        static_dir=base_dir / "app" / "static" if _is_frozen() else app_dir / "static",
        runs_dir=_runs_dir(platform_name),
        downloads_dir=_default_downloads_dir(platform_name),
        platform_name=platform_name,
        enable_powerpoint_fallback=_default_enable_powerpoint_fallback(platform_name),
        powerpoint_app_path=_path_from_env(
            "PDF_PPTX_POWERPOINT_APP_PATH",
            _default_powerpoint_app_path(platform_name),
        ),
        libreoffice_bin=_path_from_env(
            "PDF_PPTX_LIBREOFFICE_BIN",
            _default_libreoffice_bin(platform_name),
        ),
        powershell_bin=os.getenv("PDF_PPTX_POWERSHELL_BIN", "powershell.exe"),
        google_document_ai_project_id=os.getenv("PDF_PPTX_GOOGLE_DOC_AI_PROJECT_ID", "").strip() or None,
        google_document_ai_location=os.getenv("PDF_PPTX_GOOGLE_DOC_AI_LOCATION", "").strip() or None,
        google_document_ai_processor_id=os.getenv("PDF_PPTX_GOOGLE_DOC_AI_PROCESSOR_ID", "").strip() or None,
        google_document_ai_processor_version=os.getenv("PDF_PPTX_GOOGLE_DOC_AI_PROCESSOR_VERSION", "").strip() or None,
        enable_ai_qc=_env_flag("PDF_PPTX_ENABLE_AI_QC", True),
        openai_api_key=(
            os.getenv("PDF_PPTX_OPENAI_API_KEY", "").strip()
            or os.getenv("OPENAI_API_KEY", "").strip()
            or None
        ),
        openai_qc_model=os.getenv("PDF_PPTX_OPENAI_QC_MODEL", "gpt-5.3-chat-latest").strip() or "gpt-5.3-chat-latest",
        openai_qc_parallelism=max(1, int(os.getenv("PDF_PPTX_OPENAI_QC_PARALLELISM", "4") or "4")),
        openai_qc_timeout_seconds=max(10.0, float(os.getenv("PDF_PPTX_OPENAI_QC_TIMEOUT_SECONDS", "90") or "90")),
        openai_qc_max_image_dimension=max(
            0,
            int(os.getenv("PDF_PPTX_OPENAI_QC_MAX_IMAGE_DIMENSION", "0") or "0"),
        ),
        prefer_powerpoint_for_ai_qc=_env_flag("PDF_PPTX_PREFER_POWERPOINT_FOR_AI_QC", True),
        prefer_poppler_for_ai_qc=_env_flag("PDF_PPTX_PREFER_POPPLER_FOR_AI_QC", True),
        poppler_bin_dir=_path_from_env("PDF_PPTX_POPPLER_BIN_DIR", None),
        powerpoint_slide_export_macro_name=(
            os.getenv("PDF_PPTX_POWERPOINT_SLIDE_EXPORT_MACRO_NAME", "").strip()
            or ("ExportSlidesToFolder" if platform_name == "darwin" else None)
        ),
        powerpoint_reference_slide_insert_macro_name=(
            os.getenv("PDF_PPTX_POWERPOINT_REFERENCE_SLIDE_INSERT_MACRO_NAME", "").strip()
            or ("InsertReferenceSlidesFromManifest" if platform_name == "darwin" else None)
        ),
        powerpoint_slide_export_staging_dir=_path_from_env(
            "PDF_PPTX_POWERPOINT_SLIDE_EXPORT_STAGING_DIR",
            _default_powerpoint_slide_export_staging_dir(platform_name),
        ),
        powerpoint_slide_export_long_edge=max(
            1400,
            int(os.getenv("PDF_PPTX_POWERPOINT_SLIDE_EXPORT_LONG_EDGE", "2800") or "2800"),
        ),
    )
