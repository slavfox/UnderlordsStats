import pyautogui
from pathlib import Path

pyautogui.FAILSAFE = False

screenshots = Path("screenshots")
# find last screenshot by index
if any(screenshots.glob("*.png")):
    index = sorted(int(s.stem.split("_")[0]) for s in screenshots.glob("*.png"))[-1]
else:
    index = 0

games = pyautogui.locateAllOnScreen("KnockoutButton.png", grayscale=True)
for game in reversed(list(games)):
    center = pyautogui.center(game)
    pyautogui.click(center)

    button = None
    while not button:
        try:
            button = pyautogui.locateOnScreen("CloseButton.png", grayscale=True)
        except pyautogui.ImageNotFoundException:
            pass

    scrollable = pyautogui.center(button)
    scrollable = (scrollable[0], scrollable[1] - 200)
    pyautogui.moveTo(scrollable)
    print("Found game")
    pyautogui.scroll(-1000)
    pyautogui.moveTo(0, 0)
    contraptions = None
    while not contraptions:
        try:
            contraptions = pyautogui.locateOnScreen("Contraptions.png", grayscale=True)
        except pyautogui.ImageNotFoundException:
            pass

    try:
        fox = pyautogui.locateOnScreen("fox_dark.png", grayscale=True, confidence=0.9)
    except pyautogui.ImageNotFoundException:
        fox = pyautogui.locateOnScreen("fox_bright.png", grayscale=True, confidence=0.9)

    fox_area = (fox[0] - 120, fox[1], 40, 40)
    index += 1
    pyautogui.screenshot(str(screenshots / f"{index}_fox_rank.png"), region=fox_area)

    screenshot_area = (
        button[0],
        contraptions[1] + contraptions[3],
        contraptions[0] - button[0],
        button[1] - (contraptions[1] + contraptions[3]),
    )
    pyautogui.screenshot(
        str(screenshots / f"{index}_teams.png"), region=screenshot_area
    )
    print(f"Dumped game #{index}")
    print(button)
    print(pyautogui.center(button))
    pyautogui.moveTo(pyautogui.center(button))
    pyautogui.click()
    while not pyautogui.locateOnScreen("recentMatches.png", grayscale=True):
        pass
    print("Starting next dump")

print("Saved visible games")
