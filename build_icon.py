from pathlib import Path

from PIL import Image


source = Path(__file__).parent / "assets" / "app_icon.png"
target = Path(__file__).parent / "assets" / "app_icon.ico"
image = Image.open(source).convert("RGBA")
image.save(target, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
