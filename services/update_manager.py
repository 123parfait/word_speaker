# -*- coding: utf-8 -*-
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path


VERSION_FILE_NAME = "version.json"
DEFAULT_ENTRY_EXE = "WordSpeaker.exe"
ONLINE_MANIFEST_TIMEOUT_SECONDS = 20
DOWNLOAD_TIMEOUT_SECONDS = 120
UPDATE_STAGE_PREFIX = "wordspeaker_update_"

PROTECTED_UPDATE_PREFIXES = [
    "data/audio_cache/",
]

PROTECTED_UPDATE_FILES = [
    "data/app.instance.lock",
    "data/app_config.json",
    "data/corpus_index.db",
    "data/dictation_stats.json",
    "data/history.json",
    "data/pos_cache.json",
    "data/synonyms_cache.json",
    "data/translation_cache.json",
    "data/user_dictionary.json",
    "data/word_stats.json",
]


def _project_root():
    return Path(__file__).resolve().parent.parent


def app_install_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return _project_root()


def is_packaged_runtime():
    return bool(getattr(sys, "frozen", False))


def _version_file_path(base_dir=None):
    root = Path(base_dir or app_install_dir())
    return root / VERSION_FILE_NAME


def _version_file_candidates(base_dir=None):
    root = Path(base_dir or app_install_dir())
    return [
        root / VERSION_FILE_NAME,
        root / "_internal" / VERSION_FILE_NAME,
    ]


def _read_version_payload(base_dir=None):
    for version_path in _version_file_candidates(base_dir):
        if not version_path.exists():
            continue
        try:
            data = json.loads(version_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(data, dict):
            return data, version_path
    return None, None


def load_version_info(base_dir=None):
    data, _version_path = _read_version_payload(base_dir)
    if isinstance(data, dict):
        return {
            "version": str(data.get("version") or "0.0.0").strip() or "0.0.0",
            "entry_exe": str(data.get("entry_exe") or DEFAULT_ENTRY_EXE).strip() or DEFAULT_ENTRY_EXE,
            "channel_url": str(data.get("channel_url") or "").strip(),
        }
    return {
        "version": "0.0.0",
        "entry_exe": DEFAULT_ENTRY_EXE,
        "channel_url": "",
    }


def load_local_version_info():
    return load_version_info()


def _version_key(version_text):
    parts = []
    token = ""
    is_digit = None
    for ch in str(version_text or "").strip():
        current_is_digit = ch.isdigit()
        if is_digit is None:
            token = ch
            is_digit = current_is_digit
            continue
        if current_is_digit == is_digit:
            token += ch
            continue
        parts.append(int(token) if is_digit else token.casefold())
        token = ch
        is_digit = current_is_digit
    if token:
        parts.append(int(token) if is_digit else token.casefold())
    return tuple(parts)


def is_newer_version(candidate_version, current_version):
    return _version_key(candidate_version) > _version_key(current_version)


def _normalize_zip_path(path):
    return str(path or "").replace("\\", "/").strip("/ ")


def _detect_package_root(zip_file):
    files = [name for name in zip_file.namelist() if not name.endswith("/")]
    candidates = [name for name in files if name.endswith("/" + VERSION_FILE_NAME) or name == VERSION_FILE_NAME]
    if not candidates:
        raise RuntimeError("The update package does not contain version.json.")
    preferred = sorted(candidates, key=lambda item: (item.count("/"), len(item)))[0]
    if preferred == VERSION_FILE_NAME:
        return ""
    return preferred[: -len(VERSION_FILE_NAME)].rstrip("/")


def inspect_update_package(package_path):
    target = os.path.abspath(str(package_path or "").strip())
    if not target or not os.path.exists(target):
        raise FileNotFoundError("Update package not found.")
    with zipfile.ZipFile(target, "r") as zf:
        root_prefix = _detect_package_root(zf)
        version_member = f"{root_prefix}/{VERSION_FILE_NAME}" if root_prefix else VERSION_FILE_NAME
        try:
            data = json.loads(zf.read(version_member).decode("utf-8", errors="ignore"))
        except Exception as exc:
            raise RuntimeError("Failed to read version.json from the update package.") from exc
        if not isinstance(data, dict):
            raise RuntimeError("The update package version.json is invalid.")
        entry_exe = str(data.get("entry_exe") or DEFAULT_ENTRY_EXE).strip() or DEFAULT_ENTRY_EXE
        return {
            "package_path": target,
            "root_prefix": root_prefix,
            "version": str(data.get("version") or "0.0.0").strip() or "0.0.0",
            "entry_exe": entry_exe,
        }


def should_skip_update_path(relative_path):
    relative_unix = _normalize_zip_path(relative_path)
    if not relative_unix:
        return True
    for prefix in PROTECTED_UPDATE_PREFIXES:
        if relative_unix.startswith(_normalize_zip_path(prefix)):
            return True
    for item in PROTECTED_UPDATE_FILES:
        if relative_unix == _normalize_zip_path(item):
            return True
    return False


def build_update_package(source_dir, output_zip_path):
    source_root = Path(str(source_dir or "").strip())
    if not source_root.exists() or not source_root.is_dir():
        raise FileNotFoundError("Update source directory not found.")
    version_info = load_version_info(source_root)
    version_payload, version_file = _read_version_payload(source_root)
    if version_file is None or not isinstance(version_payload, dict):
        raise RuntimeError("The selected folder does not contain version.json.")

    target_zip = Path(os.path.abspath(str(output_zip_path or "").strip()))
    if not str(target_zip):
        raise ValueError("Output zip path is empty.")
    if target_zip.parent:
        target_zip.parent.mkdir(parents=True, exist_ok=True)

    root_name = source_root.name.strip() or "WordSpeaker"
    file_count = 0
    total_bytes = 0
    with zipfile.ZipFile(target_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        wrote_root_version = False
        for path in sorted(source_root.rglob("*")):
            if not path.is_file():
                continue
            relative = _normalize_zip_path(path.relative_to(source_root))
            if should_skip_update_path(relative):
                continue
            archive_name = _normalize_zip_path(Path(root_name) / relative)
            zf.write(path, archive_name)
            if archive_name == _normalize_zip_path(Path(root_name) / VERSION_FILE_NAME):
                wrote_root_version = True
            file_count += 1
            try:
                total_bytes += int(path.stat().st_size)
            except Exception:
                pass
        if not wrote_root_version:
            archive_name = _normalize_zip_path(Path(root_name) / VERSION_FILE_NAME)
            serialized = json.dumps(version_payload, ensure_ascii=False, indent=2)
            zf.writestr(archive_name, serialized)
            file_count += 1
            total_bytes += len(serialized.encode("utf-8"))
    return {
        "source_dir": str(source_root),
        "output_path": str(target_zip),
        "version": str(version_info.get("version") or "0.0.0").strip() or "0.0.0",
        "entry_exe": str(version_info.get("entry_exe") or DEFAULT_ENTRY_EXE).strip() or DEFAULT_ENTRY_EXE,
        "files": file_count,
        "bytes": total_bytes,
    }


def build_online_manifest(version, package_url, output_path, *, notes=""):
    target_path = Path(os.path.abspath(str(output_path or "").strip()))
    if not str(target_path):
        raise ValueError("Manifest output path is empty.")
    if target_path.parent:
        target_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": str(version or "").strip(),
        "url": str(package_url or "").strip(),
    }
    notes_text = str(notes or "").strip()
    if notes_text:
        payload["notes"] = notes_text
    if not payload["version"] or not payload["url"]:
        raise ValueError("Manifest requires both version and url.")
    target_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "output_path": str(target_path),
        "version": payload["version"],
        "url": payload["url"],
    }


def stage_update_package(package_path):
    info = inspect_update_package(package_path)
    stage_root = Path(tempfile.mkdtemp(prefix=UPDATE_STAGE_PREFIX))
    with zipfile.ZipFile(info["package_path"], "r") as zf:
        zf.extractall(stage_root)
    source_dir = stage_root / info["root_prefix"] if info["root_prefix"] else stage_root
    if not source_dir.exists():
        raise RuntimeError("The staged update package is missing its extracted files.")
    staged = dict(info)
    staged["stage_root"] = str(stage_root)
    staged["source_dir"] = str(source_dir)
    return staged


def fetch_online_manifest(manifest_url):
    url = str(manifest_url or "").strip()
    if not url:
        raise ValueError("Online update URL is empty.")
    req = urllib.request.Request(url, headers={"Cache-Control": "no-cache"})
    try:
        with urllib.request.urlopen(req, timeout=ONLINE_MANIFEST_TIMEOUT_SECONDS) as response:
            payload = json.loads(response.read().decode("utf-8", errors="ignore"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"Update check failed with HTTP {exc.code}.") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Update check failed: {exc.reason}") from exc
    except Exception as exc:
        raise RuntimeError(f"Update check failed: {exc}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("The online update manifest is invalid.")
    version = str(payload.get("version") or "").strip()
    package_url = urllib.parse.urljoin(url, str(payload.get("url") or "").strip())
    if not version or not package_url:
        raise RuntimeError("The online update manifest is missing version or url.")
    return {
        "manifest_url": url,
        "version": version,
        "package_url": package_url,
        "notes": str(payload.get("notes") or "").strip(),
    }


def download_update_package(package_url):
    url = str(package_url or "").strip()
    if not url:
        raise ValueError("Update package URL is empty.")
    stage_root = Path(tempfile.mkdtemp(prefix=UPDATE_STAGE_PREFIX))
    target_path = stage_root / "update_package.zip"
    req = urllib.request.Request(url, headers={"Cache-Control": "no-cache"})
    try:
        with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT_SECONDS) as response, open(target_path, "wb") as fp:
            shutil.copyfileobj(response, fp)
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"Update download failed with HTTP {exc.code}.") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Update download failed: {exc.reason}") from exc
    except Exception as exc:
        raise RuntimeError(f"Update download failed: {exc}") from exc
    return str(target_path)


def _ps_quote(text):
    return "'" + str(text or "").replace("'", "''") + "'"


def launch_staged_update(source_dir, *, entry_exe=None):
    if not is_packaged_runtime():
        raise RuntimeError("Self-update is only available in the packaged app.")

    install_dir = app_install_dir()
    source_root = Path(str(source_dir or "").strip())
    if not source_root.exists():
        raise RuntimeError("The staged update files are missing.")

    exe_name = str(entry_exe or load_local_version_info().get("entry_exe") or DEFAULT_ENTRY_EXE).strip() or DEFAULT_ENTRY_EXE
    script_path = Path(tempfile.mkdtemp(prefix=UPDATE_STAGE_PREFIX)) / "apply_update.ps1"
    current_pid = os.getpid()

    protected_prefixes = ", ".join(_ps_quote(item) for item in PROTECTED_UPDATE_PREFIXES)
    protected_files = ", ".join(_ps_quote(item) for item in PROTECTED_UPDATE_FILES)
    script = f"""$ErrorActionPreference = 'Stop'
$appDir = {_ps_quote(str(install_dir))}
$sourceDir = {_ps_quote(str(source_root))}
$entryExe = {_ps_quote(exe_name)}
$currentPid = {int(current_pid)}
$protectedPrefixes = @({protected_prefixes})
$protectedFiles = @({protected_files})
$deadline = (Get-Date).AddMinutes(5)

while ($true) {{
    $proc = Get-Process -Id $currentPid -ErrorAction SilentlyContinue
    if (-not $proc) {{
        break
    }}
    if ((Get-Date) -gt $deadline) {{
        exit 1
    }}
    Start-Sleep -Milliseconds 500
}}

Get-ChildItem -LiteralPath $sourceDir -Recurse -Force | Where-Object {{ -not $_.PSIsContainer }} | ForEach-Object {{
    $relative = $_.FullName.Substring($sourceDir.Length).TrimStart('\\', '/')
    $relativeUnix = $relative -replace '\\\\', '/'
    $skip = $false
    foreach ($prefix in $protectedPrefixes) {{
        if ($relativeUnix.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {{
            $skip = $true
            break
        }}
    }}
    if (-not $skip) {{
        foreach ($protected in $protectedFiles) {{
            if ($relativeUnix.Equals($protected, [System.StringComparison]::OrdinalIgnoreCase)) {{
                $skip = $true
                break
            }}
        }}
    }}
    if ($skip) {{
        return
    }}
    $destination = Join-Path $appDir $relative
    $destinationDir = Split-Path -Parent $destination
    if ($destinationDir) {{
        New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
    }}
    Copy-Item -LiteralPath $_.FullName -Destination $destination -Force
}}

Start-Sleep -Milliseconds 300
Start-Process -FilePath (Join-Path $appDir $entryExe)

try {{
    Remove-Item -LiteralPath $sourceDir -Recurse -Force
}} catch {{}}
try {{
    Remove-Item -LiteralPath {_ps_quote(str(script_path))} -Force
}} catch {{}}
"""
    script_path.write_text(script, encoding="utf-8")

    creation_flags = 0
    if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
        creation_flags |= subprocess.CREATE_NEW_PROCESS_GROUP
    if hasattr(subprocess, "DETACHED_PROCESS"):
        creation_flags |= subprocess.DETACHED_PROCESS

    subprocess.Popen(
        [
            "powershell.exe",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
        ],
        close_fds=True,
        creationflags=creation_flags,
    )
    return {
        "script_path": str(script_path),
        "source_dir": str(source_root),
        "entry_exe": exe_name,
    }
