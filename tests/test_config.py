from __future__ import annotations

from pathlib import Path

from app import config


def test_windows_defaults_enable_powerpoint_fallback(monkeypatch) -> None:
    monkeypatch.setenv("PDF_PPTX_PLATFORM", "windows")
    monkeypatch.delenv("PDF_PPTX_ENABLE_POWERPOINT_FALLBACK", raising=False)
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.platform_name == "windows"
        assert settings.enable_powerpoint_fallback is True
        assert settings.libreoffice_bin is not None
        assert str(settings.libreoffice_bin).lower().endswith("soffice.exe")
    finally:
        config.get_settings.cache_clear()


def test_platform_aliases_normalize_to_expected_values(monkeypatch) -> None:
    monkeypatch.setenv("PDF_PPTX_PLATFORM", "macos")
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.platform_name == "darwin"
    finally:
        config.get_settings.cache_clear()


def test_runs_dir_defaults_outside_repo_for_macos(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("PDF_PPTX_PLATFORM", "macos")
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.delenv("PDF_PPTX_RUNS_DIR", raising=False)
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.runs_dir == tmp_path / "Library" / "Application Support" / "PDF to PPTX Reference Placement" / "runs"
    finally:
        config.get_settings.cache_clear()


def test_runs_dir_can_be_overridden(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("PDF_PPTX_RUNS_DIR", str(tmp_path / "custom-runs"))
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.runs_dir == tmp_path / "custom-runs"
    finally:
        config.get_settings.cache_clear()
