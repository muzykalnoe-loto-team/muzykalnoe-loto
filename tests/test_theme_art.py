import os

import pytest
from PIL import Image

from scripts.theme_art import render_scene
from scripts.theme_scenes import THEMES


def test_reference_themes_have_two_line_titles():
    assert set(THEMES) == {"newyear", "gothic", "gatsby", "disco90", "ussr"}
    for theme in THEMES.values():
        assert theme["title"].count("\v") == 1 and theme["glow"].startswith("#") and callable(theme["scene"])


@pytest.mark.integration
def test_render_scene_gives_transparent_full_hd_png(tmp_path):
    out = render_scene(THEMES["gatsby"]["scene"], str(tmp_path / "g.png"), seed="gatsby")
    img = Image.open(out)
    assert img.size == (1920, 1080) and img.mode == "RGBA"
    assert img.getpixel((10, 540))[3] == 0  # слева, под заголовком — прозрачно
    assert img.getpixel((1690, 900))[3] > 0  # справа — рисунок


@pytest.mark.integration
def test_all_reference_themes_render(tmp_path):
    for key, theme in THEMES.items():
        assert os.path.getsize(render_scene(theme["scene"], str(tmp_path / f"{key}.png"), seed=key)) > 50_000
