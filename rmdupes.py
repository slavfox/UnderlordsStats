from pathlib import Path
from filecmp import cmp

processed = Path("processed")
screenshots = Path("screenshots")
for screenshot in screenshots.glob("*.png"):
    if not (processed / f"{screenshot.stem.split('_')[0]}.json").exists():
        print(f"Deleting {screenshot}")
        screenshot.unlink()
