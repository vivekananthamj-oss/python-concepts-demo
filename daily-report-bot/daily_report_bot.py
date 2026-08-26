"""Create a dated daily weather report by controlling Chrome and Excel.

This script intentionally performs the browser and spreadsheet interactions through
PyAutoGUI so the automation is visible on screen, as required by the assignment.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pyautogui
import pyperclip


DEFAULT_CITY = "Chennai"
APP_OPEN_TIMEOUT = 20
PAGE_LOAD_TIMEOUT = 20
CLIPBOARD_TIMEOUT = 8

# Moving the pointer to the top-left corner immediately stops PyAutoGUI.
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.35


def log(message: str) -> None:
    """Print a timestamped progress message."""
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def capture_screen(path: Path) -> None:
    """Capture the visible macOS screen without Pillow/PyScreeze."""
    result = subprocess.run(
        ["screencapture", "-x", str(path)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not path.exists() or path.stat().st_size == 0:
        detail = result.stderr.strip() or "no screenshot file was created"
        raise RuntimeError(f"The screen capture failed: {detail}")


def open_mac_app(app_name: str) -> None:
    """Launch or activate a macOS application."""
    result = subprocess.run(
        ["open", "-a", app_name],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or "unknown macOS launch error"
        raise RuntimeError(f"Could not open {app_name}: {detail}")

    # `open -a` launches an app, but it does not reliably give it keyboard focus
    # when this script is started from an IDE. Explicit activation ensures the
    # following PyAutoGUI keystrokes go to Chrome/Excel instead of the IDE.
    activation = subprocess.run(
        ["osascript", "-e", f'tell application "{app_name}" to activate'],
        capture_output=True,
        text=True,
        check=False,
    )
    if activation.returncode != 0:
        detail = activation.stderr.strip() or "unknown macOS activation error"
        raise RuntimeError(f"Could not activate {app_name}: {detail}")


def choose_excel_menu_item(menu_name: str, item_name: str) -> None:
    """Choose an exact Excel menu item through a visible PyAutoGUI click."""
    # Ask macOS where the menu is because its x-coordinate changes
    # with the application-name width and display scaling. PyAutoGUI still
    # performs the visible click and menu selection.
    position = subprocess.run(
        [
            "osascript",
            "-e",
            'tell application "System Events" to tell process "Microsoft Excel" '
            f'to get position of menu bar item "{menu_name}" of menu bar 1',
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    coordinates = [int(value) for value in re.findall(r"\d+", position.stdout)]
    if position.returncode != 0 or len(coordinates) < 2:
        detail = position.stderr.strip() or f"{menu_name} menu position was unavailable"
        raise RuntimeError(f"Could not locate Excel's {menu_name} menu: {detail}")

    menu_x, menu_y = coordinates[:2]
    pyautogui.click(menu_x + 15, menu_y + 10)
    time.sleep(1)

    item_position = subprocess.run(
        [
            "osascript",
            "-e",
            'tell application "System Events" to tell process "Microsoft Excel" ',
            "-e",
            f'set targetItem to first menu item of menu 1 of menu bar item "{menu_name}" '
            f'of menu bar 1 whose name starts with "{item_name}"',
            "-e",
            "get {position of targetItem, size of targetItem}",
            "-e",
            "end tell",
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    item_geometry = [int(value) for value in re.findall(r"\d+", item_position.stdout)]
    if item_position.returncode != 0 or len(item_geometry) < 4:
        pyautogui.press("escape")
        detail = item_position.stderr.strip() or f"{item_name} was unavailable"
        raise RuntimeError(f"Could not locate Excel's {item_name} menu item: {detail}")

    item_x, item_y, item_width, item_height = item_geometry[:4]
    pyautogui.click(item_x + item_width // 2, item_y + item_height // 2)


def mac_keystroke(key: str, modifiers: str = "") -> None:
    """Send a macOS keystroke while preserving Command/Shift modifiers."""
    using_clause = f" using {{{modifiers}}}" if modifiers else ""
    result = subprocess.run(
        [
            "osascript",
            "-e",
            f'tell application "System Events" to keystroke "{key}"{using_clause}',
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or "unknown keyboard automation error"
        raise RuntimeError(f"Could not send macOS keystroke {key}: {detail}")


def finish_excel_save_dialog(filename: str) -> None:
    """Close Go to Folder, set the filename, and activate the visible Save button."""
    script = f'''
tell application "System Events"
    tell process "Microsoft Excel"
        repeat with attemptNumber from 1 to 5
            if (count of sheets of sheet 1 of front window) is 0 then exit repeat
            perform action "AXPress" of button 1 of sheet 1 of sheet 1 of front window
            delay 1
        end repeat
        if (count of sheets of sheet 1 of front window) is not 0 then
            error "Go to Folder popup did not close"
        end if
        set value of text field "Save As:" of splitter group 1 of sheet 1 of front window to "{filename}"
        delay 1
        perform action "AXPress" of button "Save" of splitter group 1 of sheet 1 of front window
    end tell
end tell
'''
    result = subprocess.run(
        [
            "osascript",
            "-e",
            script,
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or "Save dialog controls were unavailable"
        raise RuntimeError(f"Could not complete Excel's Save As dialog: {detail}")


def size_excel_report_columns() -> None:
    """Set report column widths so Excel displays every value instead of hashes."""
    script = """
tell application "Microsoft Excel"
    tell active sheet
        set column width of column 1 to 22
        set column width of column 2 to 30
        set column width of column 3 to 52
    end tell
end tell
"""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or "unknown Excel formatting error"
        raise RuntimeError(f"Could not resize the Excel report columns: {detail}")


def wait_for_clipboard(previous_value: str, timeout: int) -> str:
    """Wait until the clipboard contains a new, non-empty text value."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        value = pyperclip.paste().strip()
        if value and value != previous_value:
            return value
        time.sleep(0.25)
    raise TimeoutError("Timed out while waiting for Chrome to copy the weather data.")


def extract_weather(city: str, copied: str) -> str:
    """Build a compact weather summary from copied Google results text."""
    temperature_match = re.search(r"[+-]?\d+(?:\.\d+)?\s*°\s*[CF]?", copied)
    if not temperature_match:
        return ""

    lowered = copied.casefold()
    conditions = (
        "thunderstorm",
        "drizzle",
        "showers",
        "rain",
        "snow",
        "overcast",
        "partly cloudy",
        "cloudy",
        "sunny",
        "clear",
        "haze",
        "mist",
    )
    condition = next((item for item in conditions if item in lowered), "")
    temperature = re.sub(r"\s+", "", temperature_match.group(0))
    details = " ".join(part for part in (condition.title(), temperature) if part)
    return f"{city}: {details}"


def fetch_weather_from_chrome(city: str) -> str:
    """Search Google in Chrome and copy weather data from the results page."""
    query = f"{city} Weather"
    marker = f"waiting-for-weather-{time.time_ns()}"
    pyperclip.copy(marker)

    log(f"Opening Chrome and loading weather for {city}...")
    open_mac_app("Google Chrome")
    time.sleep(3)
    # Clicking the omnibox avoids an intermittent macOS/PyAutoGUI issue where
    # Command-L is received as a literal leading "l".
    screen_width, _ = pyautogui.size()
    pyautogui.click(screen_width // 2, 50)
    pyautogui.hotkey("command", "a")
    pyautogui.press("backspace")
    pyperclip.copy(query)
    pyautogui.hotkey("command", "v")
    pyautogui.press("enter")

    # The endpoint returns a small plain-text page. Retry copying while it loads.
    deadline = time.monotonic() + PAGE_LOAD_TIMEOUT
    while time.monotonic() < deadline:
        time.sleep(2)
        # Chrome can leave keyboard focus in the address bar after navigation.
        # Tab moves focus into the loaded document; the click is a fallback for
        # Chrome versions that keep the omnibox active after Tab.
        pyautogui.press("tab")
        screen_width, screen_height = pyautogui.size()
        pyautogui.click(screen_width // 2, screen_height // 2)
        pyautogui.hotkey("command", "a")
        pyautogui.hotkey("command", "c")
        try:
            copied = wait_for_clipboard(marker, CLIPBOARD_TIMEOUT)
        except TimeoutError:
            continue
        weather = extract_weather(city, copied)
        if weather:
            log(f"Copied from the page: {weather}")
            return weather
        pyperclip.copy(marker)

    raise RuntimeError(
        "Chrome did not expose the expected weather text. Check the internet "
        "connection and make sure the page is not showing a consent or error screen."
    )


def comment_for(weather: str) -> str:
    """Generate a short operations comment from the visible weather text."""
    lowered = weather.casefold()
    if any(word in lowered for word in ("rain", "drizzle", "shower", "thunder")):
        return "Carry an umbrella and allow extra travel time."
    if any(word in lowered for word in ("sunny", "clear", "☀")):
        return "Good conditions for outdoor activities."
    if any(word in lowered for word in ("cloud", "overcast", "☁")):
        return "Cloudy conditions; monitor for weather changes."
    return "Weather update captured for daily planning."


def write_report_in_excel(
    observed_at: datetime,
    weather: str,
    comment: str,
    workbook_path: Path,
) -> None:
    """Create the report in Excel and save it through the visible user interface."""
    row_text = (
        "Date & Time\tFetched Weather Data\tComment\n"
        f"{observed_at:%Y-%m-%d %H:%M:%S}\t{weather}\t{comment}"
    )
    pyperclip.copy(row_text)

    log("Opening Microsoft Excel...")
    open_mac_app("Microsoft Excel")
    time.sleep(5)
    choose_excel_menu_item("File", "New")
    time.sleep(3)

    # Paste the tab-separated values at A1; Excel splits them into three columns.
    choose_excel_menu_item("Edit", "Paste")
    time.sleep(1)
    size_excel_report_columns()
    time.sleep(1)

    log(f"Saving workbook as {workbook_path.name}...")
    # Use the visible File menu because Command-key shortcuts are intermittently
    # redirected to Excel's search field on some macOS/Excel versions.
    choose_excel_menu_item("File", "Save As")
    time.sleep(3)
    mac_keystroke("g", "command down, shift down")
    time.sleep(1)
    pyperclip.copy(str(workbook_path.parent))
    mac_keystroke("v", "command down")
    # Return navigates to the folder but leaves the Go to Folder popup open.
    pyautogui.press("enter")
    time.sleep(1)
    finish_excel_save_dialog(workbook_path.name)
    time.sleep(8)

    # Some Excel versions show a format confirmation dialog.
    if not workbook_path.exists():
        pyautogui.press("enter")
        time.sleep(4)
    if not workbook_path.exists():
        failure_path = workbook_path.parent / "excel_save_failure.png"
        capture_screen(failure_path)
        raise RuntimeError(
            "Excel did not create the expected workbook. The Save As dialog may "
            f"have required an unexpected choice. Screenshot: {failure_path}"
        )


def save_final_screenshot(screenshot_path: Path) -> None:
    """Capture the final visible Excel sheet."""
    time.sleep(2)
    capture_screen(screenshot_path)
    log(f"Saved screenshot as {screenshot_path.name}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--city", default=DEFAULT_CITY, help="City shown in the weather report")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "output",
        help="Folder for the dated workbook and screenshot",
    )
    parser.add_argument(
        "--start-delay",
        type=int,
        default=5,
        help="Seconds to wait before controlling the mouse and keyboard",
    )
    return parser.parse_args()


def available_output_path(output_dir: Path, stem: str, suffix: str) -> Path:
    """Return a non-existing path while keeping the required date in its name."""
    candidate = output_dir / f"{stem}{suffix}"
    sequence = 2
    while candidate.exists():
        candidate = output_dir / f"{stem}_{sequence}{suffix}"
        sequence += 1
    return candidate


def main() -> int:
    if sys.platform != "darwin":
        print("This version targets macOS with Chrome and Microsoft Excel.", file=sys.stderr)
        return 2

    args = parse_args()
    output_dir = args.output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    started_at = datetime.now()
    date_stamp = started_at.strftime("%Y-%m-%d")
    stem = f"daily_report_{date_stamp}"
    workbook_path = available_output_path(output_dir, stem, ".xlsx")
    screenshot_path = available_output_path(output_dir, stem, ".png")

    log(
        f"Starting in {args.start_delay} seconds. Do not move the mouse or type. "
        "Move the pointer to the top-left corner to stop the bot."
    )
    time.sleep(max(args.start_delay, 0))

    try:
        weather = fetch_weather_from_chrome(args.city)
        write_report_in_excel(
            observed_at=datetime.now(),
            weather=weather,
            comment=comment_for(weather),
            workbook_path=workbook_path,
        )
        save_final_screenshot(screenshot_path)
    except pyautogui.FailSafeException:
        print("Automation stopped by the PyAutoGUI fail-safe.", file=sys.stderr)
        return 130
    except (OSError, RuntimeError, TimeoutError) as exc:
        print(f"Automation failed: {exc}", file=sys.stderr)
        return 1

    log("Daily report automation completed successfully.")
    print(f"Workbook: {workbook_path}")
    print(f"Screenshot: {screenshot_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
