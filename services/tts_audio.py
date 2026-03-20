# -*- coding: utf-8 -*-
import os
import tempfile
import wave


def prepend_silence_to_wav(
    path,
    *,
    silence_ms=0,
    default_channels=1,
    default_sample_width=2,
    default_sample_rate=22050,
):
    duration_ms = max(0, int(silence_ms or 0))
    if duration_ms <= 0:
        return path
    try:
        with wave.open(path, "rb") as wav_fp:
            channels = int(wav_fp.getnchannels() or default_channels)
            sample_width = int(wav_fp.getsampwidth() or default_sample_width)
            sample_rate = int(wav_fp.getframerate() or default_sample_rate)
            frames = wav_fp.readframes(wav_fp.getnframes())
        silence_frames = max(1, int(sample_rate * duration_ms / 1000.0))
        silence_bytes = b"\x00" * silence_frames * channels * sample_width
        fd, temp_path = tempfile.mkstemp(prefix="wordspeaker_", suffix=".wav")
        os.close(fd)
        with wave.open(temp_path, "wb") as wav_fp:
            wav_fp.setnchannels(channels)
            wav_fp.setsampwidth(sample_width)
            wav_fp.setframerate(sample_rate)
            wav_fp.writeframes(silence_bytes + frames)
        try:
            os.remove(path)
        except Exception:
            pass
        return temp_path
    except Exception:
        return path


def wav_duration_seconds(path):
    try:
        with wave.open(path, "rb") as wav_fp:
            frame_rate = int(wav_fp.getframerate() or 0)
            frame_count = int(wav_fp.getnframes() or 0)
        if frame_rate <= 0:
            return 0.0
        return max(0.0, float(frame_count) / float(frame_rate))
    except Exception:
        return 0.0


def cleanup_temp_wavs(paths):
    for wav_path in paths or []:
        try:
            if wav_path and os.path.exists(wav_path):
                os.remove(wav_path)
        except Exception:
            pass
