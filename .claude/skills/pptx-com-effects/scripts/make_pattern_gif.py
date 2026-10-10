"""Original looping pattern GIF: vivid capsules, dots and rings drifting over a hue-shifting base."""
import colorsys
import math
import random
import sys

from PIL import Image, ImageDraw

W, H, SS = 640, 480, 2          # output size, supersampling factor
FRAMES, FPS = 60, 20            # 3 second seamless loop
PALETTE = ["#FF3D2E", "#FF8A00", "#FFD400", "#00E5A8", "#00B7FF", "#3A3DFF", "#B026FF", "#FF2E93", "#FFFFFF", "#F7B8C4"]


def make_shapes(rng):
    shapes = []
    for _ in range(150):
        shapes.append({
            "kind": rng.choices(["capsule", "dot", "ring"], [6, 3, 1])[0],
            "cx": rng.uniform(0, W), "cy": rng.uniform(0, H),
            "ax": rng.uniform(30, 120), "ay": rng.uniform(20, 90),      # drift amplitude
            "fx": rng.choice([1, 1, 2]), "fy": rng.choice([1, 2, 2]),   # whole cycles per loop -> seamless
            "px": rng.uniform(0, math.tau), "py": rng.uniform(0, math.tau),
            "len": rng.uniform(30, 110), "thick": rng.uniform(12, 34),
            "rot0": rng.uniform(0, math.tau), "spin": rng.choice([-2, -1, 1, 2]),
            "color": rng.choice(PALETTE),
            "pulse": rng.uniform(0, math.tau),
        })
    return shapes


def draw_frame(shapes, t):
    """t in [0, 1): position inside the loop."""
    a = math.tau * t
    r, g, b = colorsys.hsv_to_rgb(t, 0.85, 0.32)                        # base colour cycles through the hues once
    img = Image.new("RGB", (W * SS, H * SS), (int(r * 255), int(g * 255), int(b * 255)))
    d = ImageDraw.Draw(img)
    for s in shapes:
        x = (s["cx"] + s["ax"] * math.sin(s["fx"] * a + s["px"])) * SS
        y = (s["cy"] + s["ay"] * math.cos(s["fy"] * a + s["py"])) * SS
        scale = 1 + 0.25 * math.sin(2 * a + s["pulse"])
        th = s["thick"] * scale * SS
        if s["kind"] == "capsule":
            ang = s["rot0"] + s["spin"] * a
            dx, dy = math.cos(ang) * s["len"] * SS / 2, math.sin(ang) * s["len"] * SS / 2
            d.line([(x - dx, y - dy), (x + dx, y + dy)], fill=s["color"], width=int(th))
            for ex, ey in ((x - dx, y - dy), (x + dx, y + dy)):
                d.ellipse([ex - th / 2, ey - th / 2, ex + th / 2, ey + th / 2], fill=s["color"])
        elif s["kind"] == "dot":
            d.ellipse([x - th, y - th, x + th, y + th], fill=s["color"])
        else:
            rad = th * 1.6
            d.ellipse([x - rad, y - rad, x + rad, y + rad], outline=s["color"], width=int(6 * SS))
    return img.resize((W, H), Image.LANCZOS)


def main(out_path):
    shapes = make_shapes(random.Random(7))
    frames = [draw_frame(shapes, i / FRAMES) for i in range(FRAMES)]
    frames = [f.quantize(colors=128, method=Image.MEDIANCUT, dither=Image.NONE) for f in frames]
    frames[0].save(out_path, save_all=True, append_images=frames[1:], duration=1000 // FPS, loop=0, optimize=False)
    assert Image.open(out_path).n_frames == FRAMES


if __name__ == "__main__":
    main(sys.argv[1])
