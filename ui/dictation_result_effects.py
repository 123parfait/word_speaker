# -*- coding: utf-8 -*-
import math
import random


def stop_result_effect(host):
    for after_id in list(getattr(host, "dictation_result_fx_after_ids", []) or []):
        try:
            host.after_cancel(after_id)
        except Exception:
            pass
    host.dictation_result_fx_after_ids = []
    canvas = getattr(host, "dictation_result_fx_canvas", None)
    if canvas and canvas.winfo_exists():
        try:
            canvas.destroy()
        except Exception:
            pass
    host.dictation_result_fx_canvas = None


def start_result_effect(host, accuracy):
    stop_result_effect(host)
    frame = getattr(host, "dictation_result_frame", None)
    if not frame or not frame.winfo_exists():
        return
    host.dictation_result_frame.update_idletasks()
    width = max(420, int(frame.winfo_width() or 760))
    canvas_height = 220
    canvas = __import__("tkinter").Canvas(
        frame,
        width=width,
        height=canvas_height,
        highlightthickness=0,
        bd=0,
        bg="#f6f7fb",
    )
    canvas.place(x=0, y=0, relwidth=1.0, height=canvas_height)
    host.dictation_result_fx_canvas = canvas
    host.dictation_result_fx_after_ids = []

    if accuracy < 50.0:
        _animate_rainbow(host, canvas)
    elif accuracy < 80.0:
        _animate_flower(host, canvas)
    else:
        _animate_fireworks(host, canvas)


def _schedule(host, delay_ms, callback):
    after_id = host.after(delay_ms, callback)
    host.dictation_result_fx_after_ids.append(after_id)
    return after_id


def _finish_after(host, duration_ms):
    _schedule(host, duration_ms, lambda: stop_result_effect(host))


def _animate_rainbow(host, canvas):
    duration_ms = 3000
    colors = ["#ef4444", "#f97316", "#facc15", "#22c55e", "#38bdf8", "#6366f1", "#d946ef"]
    width = int(canvas.winfo_reqwidth() or 760)
    center_x = width // 2
    center_y = 200
    max_radius = min(width // 2 - 24, 180)

    def _step(frame_idx=0):
        if not canvas.winfo_exists():
            return
        progress = min(1.0, frame_idx / 45.0)
        extent = 180.0 * progress
        canvas.delete("all")
        for idx, color in enumerate(colors):
            radius = max_radius - idx * 16
            canvas.create_arc(
                center_x - radius,
                center_y - radius,
                center_x + radius,
                center_y + radius,
                start=180,
                extent=extent,
                style="arc",
                width=12,
                outline=color,
            )
        alpha = min(1.0, progress * 1.6)
        cloud_fill = "#ffffff"
        for offset in (-120, 100):
            x = center_x + offset
            y = center_y - 28
            canvas.create_oval(x - 34, y - 16, x + 10, y + 16, fill=cloud_fill, outline="")
            canvas.create_oval(x - 6, y - 22, x + 42, y + 18, fill=cloud_fill, outline="")
            canvas.create_oval(x + 20, y - 16, x + 58, y + 16, fill=cloud_fill, outline="")
            canvas.create_rectangle(x - 14, y, x + 36, y + 18, fill=cloud_fill, outline="")
        if progress < 1.0:
            _schedule(host, 60, lambda: _step(frame_idx + 1))

    _step()
    _finish_after(host, duration_ms)


def _animate_flower(host, canvas):
    duration_ms = 3000
    width = int(canvas.winfo_reqwidth() or 760)
    center_x = width // 2
    center_y = 116
    petal_colors = ["#fb7185", "#f472b6", "#c084fc", "#60a5fa", "#34d399", "#fbbf24"]

    def _step(frame_idx=0):
        if not canvas.winfo_exists():
            return
        progress = min(1.0, frame_idx / 42.0)
        bloom = 24 + (74 * progress)
        petal_length = 38 + 54 * progress
        sway = math.sin(frame_idx / 6.0) * 3.0
        canvas.delete("all")
        canvas.create_line(center_x, center_y + 22, center_x + sway, 208, fill="#16a34a", width=6, smooth=True)
        canvas.create_oval(center_x - 10, 166, center_x + 26, 188, fill="#22c55e", outline="")
        canvas.create_oval(center_x - 30, 146, center_x + 8, 168, fill="#22c55e", outline="")
        for idx in range(8):
            angle = (math.pi * 2 / 8.0) * idx + progress * 0.25
            dx = math.cos(angle) * petal_length
            dy = math.sin(angle) * petal_length * 0.72
            color = petal_colors[idx % len(petal_colors)]
            canvas.create_oval(
                center_x + dx - bloom * 0.42,
                center_y + dy - bloom * 0.22,
                center_x + dx + bloom * 0.42,
                center_y + dy + bloom * 0.22,
                fill=color,
                outline="",
            )
        canvas.create_oval(center_x - 22, center_y - 22, center_x + 22, center_y + 22, fill="#facc15", outline="")
        if progress < 1.0:
            _schedule(host, 60, lambda: _step(frame_idx + 1))

    _step()
    _finish_after(host, duration_ms)


def _animate_fireworks(host, canvas):
    duration_ms = 3000
    width = int(canvas.winfo_reqwidth() or 760)
    height = 220
    bursts = []
    confetti = []
    rng = random.Random()
    burst_colors = ["#f43f5e", "#f59e0b", "#22c55e", "#38bdf8", "#a855f7", "#f97316"]
    for _ in range(6):
        bursts.append(
            {
                "x": rng.randint(80, max(120, width - 80)),
                "y": rng.randint(40, 120),
                "start": rng.randint(0, 18),
                "color": rng.choice(burst_colors),
            }
        )
    for _ in range(48):
        confetti.append(
            {
                "x": rng.randint(10, max(20, width - 10)),
                "y": rng.randint(-120, 20),
                "speed": rng.uniform(1.8, 4.6),
                "drift": rng.uniform(-1.6, 1.6),
                "size": rng.randint(6, 12),
                "color": rng.choice(burst_colors + ["#14b8a6", "#eab308"]),
            }
        )

    def _step(frame_idx=0):
        if not canvas.winfo_exists():
            return
        canvas.delete("all")
        for burst in bursts:
            age = frame_idx - burst["start"]
            if age < 0:
                continue
            radius = 8 + age * 4.8
            if radius > 90:
                continue
            x = burst["x"]
            y = burst["y"]
            for ray in range(12):
                angle = (math.pi * 2 / 12.0) * ray
                inner = max(4, radius * 0.18)
                outer = radius
                canvas.create_line(
                    x + math.cos(angle) * inner,
                    y + math.sin(angle) * inner,
                    x + math.cos(angle) * outer,
                    y + math.sin(angle) * outer,
                    fill=burst["color"],
                    width=3,
                )
            canvas.create_oval(x - 6, y - 6, x + 6, y + 6, fill="#ffffff", outline="")
        for piece in confetti:
            y = piece["y"] + frame_idx * piece["speed"]
            x = piece["x"] + frame_idx * piece["drift"]
            if y > height + 20:
                continue
            size = piece["size"]
            canvas.create_rectangle(x, y, x + size, y + size * 0.55, fill=piece["color"], outline="")
        if frame_idx < 50:
            _schedule(host, 60, lambda: _step(frame_idx + 1))

    _step()
    _finish_after(host, duration_ms)
