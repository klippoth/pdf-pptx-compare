from __future__ import annotations

from pathlib import Path

from app import config


def test_project_specific_openai_key_overrides_global_openai_key(monkeypatch) -> None:
    monkeypatch.setenv("OPENAI_API_KEY", "global-key")
    monkeypatch.setenv("PDF_PPTX_OPENAI_API_KEY", "project-key")
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.openai_api_key == "project-key"
    finally:
        config.get_settings.cache_clear()


def test_ai_qc_can_be_disabled_via_env(monkeypatch) -> None:
    monkeypatch.setenv("PDF_PPTX_ENABLE_AI_QC", "false")
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.enable_ai_qc is False
    finally:
        config.get_settings.cache_clear()


def test_apply_env_file_loads_project_local_values(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.delenv("PDF_PPTX_OPENAI_API_KEY", raising=False)
    env_file = tmp_path / ".env.local"
    env_file.write_text(
        'PDF_PPTX_OPENAI_API_KEY="from-local-file"\nPDF_PPTX_OPENAI_QC_MODEL=gpt-5.3-chat-latest\n',
        encoding="utf-8",
    )

    config._apply_env_file(env_file)
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.openai_api_key == "from-local-file"
        assert settings.openai_qc_model == "gpt-5.3-chat-latest"
    finally:
        config.get_settings.cache_clear()


def test_project_env_roots_include_packaged_sidecar_locations_for_macos_bundle(tmp_path: Path) -> None:
    app_dir = tmp_path / "bundle-root" / "app"
    executable_path = tmp_path / "dist" / "PDFtoPPTXReference.app" / "Contents" / "MacOS" / "PDFtoPPTXReference"
    roots = config._project_env_roots(
        app_dir,
        frozen=True,
        executable_path=executable_path,
        resource_root=tmp_path / "bundle-root",
    )

    assert roots == (
        tmp_path / "bundle-root",
        tmp_path / "dist" / "PDFtoPPTXReference.app" / "Contents" / "MacOS",
        tmp_path / "dist",
    )


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


def test_powerpoint_slide_export_staging_dir_defaults_inside_powerpoint_container_on_macos(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("PDF_PPTX_PLATFORM", "macos")
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.delenv("PDF_PPTX_POWERPOINT_SLIDE_EXPORT_STAGING_DIR", raising=False)
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.powerpoint_slide_export_staging_dir == (
            tmp_path
            / "Library"
            / "Containers"
            / "com.microsoft.Powerpoint"
            / "Data"
            / "Documents"
            / "codex-png-out"
        )
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


def test_ai_qc_renderer_preferences_can_be_overridden(monkeypatch, tmp_path: Path) -> None:
    monkeypatch.setenv("PDF_PPTX_PREFER_POWERPOINT_FOR_AI_QC", "false")
    monkeypatch.setenv("PDF_PPTX_PREFER_POPPLER_FOR_AI_QC", "false")
    monkeypatch.setenv("PDF_PPTX_POPPLER_BIN_DIR", str(tmp_path / "poppler-bin"))
    monkeypatch.setenv("PDF_PPTX_POWERPOINT_SLIDE_EXPORT_MACRO_NAME", "ExportSlidesToFolder")
    monkeypatch.setenv("PDF_PPTX_POWERPOINT_REFERENCE_SLIDE_INSERT_MACRO_NAME", "InsertReferenceSlidesFromManifest")
    monkeypatch.setenv("PDF_PPTX_POWERPOINT_SLIDE_EXPORT_STAGING_DIR", str(tmp_path / "ppt-stage"))
    monkeypatch.setenv("PDF_PPTX_POWERPOINT_SLIDE_EXPORT_LONG_EDGE", "3200")
    monkeypatch.setenv("PDF_PPTX_OPENAI_QC_MAX_IMAGE_DIMENSION", "2800")
    config.get_settings.cache_clear()

    try:
        settings = config.get_settings()

        assert settings.prefer_powerpoint_for_ai_qc is False
        assert settings.prefer_poppler_for_ai_qc is False
        assert settings.poppler_bin_dir == tmp_path / "poppler-bin"
        assert settings.powerpoint_slide_export_macro_name == "ExportSlidesToFolder"
        assert settings.powerpoint_reference_slide_insert_macro_name == "InsertReferenceSlidesFromManifest"
        assert settings.powerpoint_slide_export_staging_dir == tmp_path / "ppt-stage"
        assert settings.powerpoint_slide_export_long_edge == 3200
        assert settings.openai_qc_max_image_dimension == 2800
    finally:
        config.get_settings.cache_clear()
