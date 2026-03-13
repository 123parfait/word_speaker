# -*- coding: utf-8 -*-
import threading

_lock = threading.Lock()
_translation = None
_init_error = None
_prepare_started = False


def _find_lang(langs, prefix):
    prefix = str(prefix).lower()
    for lang in langs:
        code = str(getattr(lang, "code", "")).lower()
        if code == prefix or code.startswith(prefix + "_") or code.startswith(prefix + "-"):
            return lang
    return None


def _ensure_translation():
    global _translation, _init_error
    with _lock:
        if _translation is not None:
            return _translation
        if _init_error:
            raise RuntimeError(_init_error)

    try:
        import argostranslate.package
        import argostranslate.translate
    except Exception as e:
        err = (
            "Argos Translate is not available. Install it with: "
            "pip install argostranslate"
        )
        with _lock:
            _init_error = err
        raise RuntimeError(err) from e

    try:
        installed_languages = argostranslate.translate.get_installed_languages()
        from_lang = _find_lang(installed_languages, "en")
        to_lang = _find_lang(installed_languages, "zh")
        if not (from_lang and to_lang):
            argostranslate.package.update_package_index()
            available_packages = argostranslate.package.get_available_packages()
            package = None
            for pkg in available_packages:
                if pkg.from_code == "en" and pkg.to_code.startswith("zh"):
                    package = pkg
                    break
            if package is None:
                raise RuntimeError("No Argos package found for English -> Chinese.")
            download_path = package.download()
            argostranslate.package.install_from_path(download_path)
            installed_languages = argostranslate.translate.get_installed_languages()
            from_lang = _find_lang(installed_languages, "en")
            to_lang = _find_lang(installed_languages, "zh")
            if not (from_lang and to_lang):
                raise RuntimeError("Failed to initialize English -> Chinese translation package.")

        translation = from_lang.get_translation(to_lang)
        with _lock:
            _translation = translation
            _init_error = None
        return translation
    except Exception as e:
        with _lock:
            _init_error = str(e)
        raise


def prepare_async():
    global _prepare_started
    with _lock:
        if _prepare_started:
            return
        _prepare_started = True

    def _run():
        global _prepare_started
        try:
            _ensure_translation()
        finally:
            with _lock:
                _prepare_started = False

    threading.Thread(target=_run, daemon=True).start()


def translate_text(text):
    translation = _ensure_translation()
    return translation.translate(str(text))


def translate_words(words):
    result = {}
    for w in words:
        try:
            result[w] = translate_text(w)
        except Exception:
            result[w] = ""
    return result
