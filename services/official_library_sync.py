# -*- coding: utf-8 -*-
import os
import shutil

from services.bundled_corpus import import_bundled_corpus_package
from services.resource_pack import import_word_resource_pack
from services.tts import import_shared_audio_cache_package
from services.update_manager import download_update_package


def resolve_official_library_urls(*, current_info, release_base_url=""):
    info = dict(current_info or {})
    release_base = str(release_base_url or "").strip()
    resource_pack_url = str(info.get("word_resource_pack_url") or "").strip()
    bundled_corpus_url = str(info.get("bundled_corpus_package_url") or "").strip()
    if not resource_pack_url and release_base:
        resource_pack_url = f"{release_base}/wordspeaker_word_resource_pack.wspack"
    if not bundled_corpus_url and release_base:
        bundled_corpus_url = f"{release_base}/wordspeaker_bundled_corpus.zip"
    return {
        "resource_pack_url": resource_pack_url,
        "bundled_corpus_url": bundled_corpus_url,
    }


def sync_official_library(*, shared_cache_package_url, resource_pack_url, bundled_corpus_url, load_word_resource_entries):
    if not shared_cache_package_url:
        raise RuntimeError("Official shared-cache package URL is missing.")
    if not resource_pack_url:
        raise RuntimeError("Official word resource pack URL is missing.")
    if not bundled_corpus_url:
        raise RuntimeError("Official bundled corpus package URL is missing.")
    if not callable(load_word_resource_entries):
        raise RuntimeError("Word resource pack loader is not available.")

    package_path = ""
    resource_pack_path = ""
    bundled_corpus_path = ""
    try:
        package_path = download_update_package(shared_cache_package_url)
        shared_cache_result = import_shared_audio_cache_package(package_path)

        resource_pack_path = download_update_package(resource_pack_url)
        resource_pack_result = import_word_resource_pack(resource_pack_path)
        load_result = load_word_resource_entries(resource_pack_result.get("entries") or [])
        if not load_result:
            raise RuntimeError("Official word resource pack contained no valid entries.")

        bundled_corpus_path = download_update_package(bundled_corpus_url)
        corpus_result = import_bundled_corpus_package(bundled_corpus_path)
        return {
            "shared_cache_result": shared_cache_result,
            "word_pack_result": load_result,
            "corpus_result": corpus_result,
        }
    finally:
        for path in (package_path, resource_pack_path, bundled_corpus_path):
            if not path:
                continue
            try:
                shutil.rmtree(os.path.dirname(path), ignore_errors=True)
            except Exception:
                pass
