# -*- coding: utf-8 -*-
import json
import os
import shutil
import tempfile
import time
import zipfile


def export_shared_audio_cache_package(
    package_path,
    *,
    get_export_entries,
    zip_entry_path,
    shared_cache_metadata_file,
    shared_cache_package_manifest,
    shared_cache_package_version,
    export_shared_metadata_payload,
    sha1_file,
):
    target_path = os.path.abspath(str(package_path or "").strip())
    if not target_path:
        raise ValueError("Package path is empty.")

    entries = get_export_entries()
    if not entries:
        return {
            "ok": False,
            "package_path": target_path,
            "entries": 0,
            "bytes": 0,
            "message": "No shared audio cache is available yet.",
        }

    target_dir = os.path.dirname(target_path)
    if target_dir:
        os.makedirs(target_dir, exist_ok=True)
    manifest_entries = []
    total_bytes = 0

    with zipfile.ZipFile(target_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for item in entries:
            audio_path = item["audio_path"]
            meta_payload = dict(item["metadata"] or {})
            relative_path = item["relative_path"]
            meta_relative_path = item["meta_relative_path"]
            arc_audio_path = zip_entry_path("global", relative_path)
            arc_meta_path = zip_entry_path("global", meta_relative_path)
            zf.write(audio_path, arc_audio_path)
            zf.writestr(arc_meta_path, json.dumps(meta_payload, ensure_ascii=False, indent=2))
            audio_size = int(os.path.getsize(audio_path)) if os.path.exists(audio_path) else 0
            total_bytes += audio_size
            manifest_entries.append(
                {
                    "relative_path": relative_path,
                    "meta_relative_path": meta_relative_path,
                    "text": str(meta_payload.get("text") or ""),
                    "backend": str(meta_payload.get("backend") or ""),
                    "desired_backend": str(meta_payload.get("desired_backend") or ""),
                    "updated_at": int(meta_payload.get("updated_at") or 0),
                    "audio_size": audio_size,
                    "audio_sha1": sha1_file(audio_path),
                }
            )

        manifest = {
            "kind": "wordspeaker.shared_audio_cache",
            "version": shared_cache_package_version,
            "exported_at": int(time.time()),
            "entry_count": len(manifest_entries),
            "entries": manifest_entries,
        }
        zf.writestr(
            shared_cache_metadata_file,
            json.dumps(export_shared_metadata_payload(), ensure_ascii=False, indent=2),
        )
        zf.writestr(
            shared_cache_package_manifest,
            json.dumps(manifest, ensure_ascii=False, indent=2),
        )

    return {
        "ok": True,
        "package_path": target_path,
        "entries": len(manifest_entries),
        "bytes": total_bytes,
        "message": "Shared cache package exported.",
    }


def import_shared_audio_cache_package(
    package_path,
    *,
    shared_cache_package_manifest,
    shared_cache_metadata_file,
    safe_rel_path,
    zip_entry_path,
    shared_cache_target_path,
    sha1_file,
    load_cache_metadata,
    save_cache_metadata,
    import_shared_metadata_payload,
    cleanup_duplicate_source_cache_entries,
    collapse_existing_lightweight_source_caches,
    collapse_all_source_cache_entities_to_aliases,
):
    target_path = os.path.abspath(str(package_path or "").strip())
    if not target_path or not os.path.exists(target_path):
        raise FileNotFoundError("Cache package does not exist.")

    summary = {
        "ok": True,
        "package_path": target_path,
        "imported": 0,
        "replaced": 0,
        "skipped_same": 0,
        "skipped_older": 0,
        "metadata_translations": 0,
        "metadata_pos": 0,
        "metadata_phonetics": 0,
        "errors": [],
    }

    with zipfile.ZipFile(target_path, "r") as zf:
        if shared_cache_package_manifest not in zf.namelist():
            raise RuntimeError("This file is not a valid Word Speaker shared-cache package.")
        manifest = json.loads(zf.read(shared_cache_package_manifest).decode("utf-8", errors="ignore"))
        if not isinstance(manifest, dict) or manifest.get("kind") != "wordspeaker.shared_audio_cache":
            raise RuntimeError("Unsupported shared-cache package format.")
        entries = manifest.get("entries") or []
        if not isinstance(entries, list):
            raise RuntimeError("Shared-cache package manifest is invalid.")
        if shared_cache_metadata_file in zf.namelist():
            try:
                metadata_payload = json.loads(zf.read(shared_cache_metadata_file).decode("utf-8", errors="ignore"))
                metadata_result = import_shared_metadata_payload(metadata_payload)
                summary["metadata_translations"] = int(metadata_result.get("translations") or 0)
                summary["metadata_pos"] = int(metadata_result.get("pos") or 0)
                summary["metadata_phonetics"] = int(metadata_result.get("phonetics") or 0)
            except Exception as exc:
                summary["errors"].append(f"metadata.json: {exc}")

        for entry in entries:
            if not isinstance(entry, dict):
                continue
            relative_path = safe_rel_path(entry.get("relative_path"))
            meta_relative_path = safe_rel_path(entry.get("meta_relative_path"))
            if not relative_path or not meta_relative_path:
                summary["errors"].append("Skipped one malformed cache entry.")
                continue

            arc_audio_path = zip_entry_path("global", relative_path)
            arc_meta_path = zip_entry_path("global", meta_relative_path)
            if arc_audio_path not in zf.namelist() or arc_meta_path not in zf.namelist():
                summary["errors"].append(f"Missing files for cache entry: {relative_path}")
                continue

            try:
                metadata = json.loads(zf.read(arc_meta_path).decode("utf-8", errors="ignore"))
            except Exception:
                summary["errors"].append(f"Invalid metadata for cache entry: {relative_path}")
                continue
            if not isinstance(metadata, dict):
                summary["errors"].append(f"Invalid metadata object for cache entry: {relative_path}")
                continue

            cache_path = shared_cache_target_path(relative_path=relative_path, metadata=metadata)
            if not cache_path:
                summary["errors"].append(f"Unable to resolve cache path for: {relative_path}")
                continue

            incoming_sha1 = str(entry.get("audio_sha1") or "").strip().lower()
            incoming_updated_at = int(entry.get("updated_at") or metadata.get("updated_at") or 0)
            existing_sha1 = ""
            existing_updated_at = 0
            if os.path.exists(cache_path):
                try:
                    existing_sha1 = sha1_file(cache_path)
                except Exception:
                    existing_sha1 = ""
            existing_meta = load_cache_metadata(cache_path)
            if isinstance(existing_meta, dict):
                try:
                    existing_updated_at = int(existing_meta.get("updated_at") or 0)
                except Exception:
                    existing_updated_at = 0
            if not existing_updated_at and os.path.exists(cache_path):
                try:
                    existing_updated_at = int(os.path.getmtime(cache_path))
                except Exception:
                    existing_updated_at = 0

            if incoming_sha1 and existing_sha1 and incoming_sha1 == existing_sha1:
                if not isinstance(existing_meta, dict) or not existing_meta:
                    payload = dict(metadata)
                    payload["source_path"] = "shared"
                    try:
                        payload["updated_at"] = int(incoming_updated_at or os.path.getmtime(cache_path))
                    except Exception:
                        payload["updated_at"] = incoming_updated_at
                    save_cache_metadata(cache_path, payload)
                summary["skipped_same"] += 1
                continue
            if os.path.exists(cache_path) and incoming_updated_at and existing_updated_at and incoming_updated_at < existing_updated_at:
                summary["skipped_older"] += 1
                continue

            os.makedirs(os.path.dirname(cache_path), exist_ok=True)
            fd, temp_audio_path = tempfile.mkstemp(prefix="wordspeaker_import_", suffix=".wav")
            os.close(fd)
            try:
                with zf.open(arc_audio_path, "r") as src, open(temp_audio_path, "wb") as dst:
                    shutil.copyfileobj(src, dst)
                shutil.copyfile(temp_audio_path, cache_path)
                if incoming_updated_at > 0:
                    try:
                        os.utime(cache_path, (incoming_updated_at, incoming_updated_at))
                    except Exception:
                        pass
                payload = dict(metadata)
                payload["source_path"] = "shared"
                payload["updated_at"] = int(incoming_updated_at or os.path.getmtime(cache_path))
                save_cache_metadata(cache_path, payload)
                if existing_sha1:
                    summary["replaced"] += 1
                else:
                    summary["imported"] += 1
            except Exception as exc:
                summary["errors"].append(f"{relative_path}: {exc}")
            finally:
                try:
                    os.remove(temp_audio_path)
                except Exception:
                    pass

    cleanup_duplicate_source_cache_entries()
    collapse_existing_lightweight_source_caches()
    collapse_all_source_cache_entities_to_aliases()
    return summary
