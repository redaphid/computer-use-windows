"""
Computer Use MCP Server

Exposes computer control capabilities (screenshot, mouse, keyboard) as MCP tools.
Can be used with Claude Code or any MCP-compatible client.

Usage:
    # Stdio transport (for Claude Code)
    python computer_use_mcp.py

    # HTTP transport (for curl testing)
    python computer_use_mcp.py --transport streamable-http --port 8080
"""

import argparse
import base64
import io
import os
import sys
import time
import uuid
from datetime import datetime
from pathlib import Path

from mcp.server.fastmcp import FastMCP

# Directory for saving screenshots - organized by session
SCREENSHOTS_BASE = Path(__file__).parent / "screenshots"
SCREENSHOTS_BASE.mkdir(exist_ok=True)

# Generate unique session ID for this run
SESSION_ID = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
SCREENSHOTS_DIR = SCREENSHOTS_BASE / SESSION_ID
SCREENSHOTS_DIR.mkdir(exist_ok=True)

# Only import GUI libraries if not just showing help
if "--help" not in sys.argv and "-h" not in sys.argv:
    import mss
    import pyautogui
    from PIL import Image, ImageOps, ImageFilter, ImageEnhance, ImageDraw

    # Safety settings
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.05

    # Import UI automation tools
    try:
        from vision_tools import WindowsUIAutomation, SmartElementFinder
        _uia_available = True
        _uia = WindowsUIAutomation()
        _finder = SmartElementFinder(use_uia=True, use_ocr=False, use_florence=False)
    except ImportError as e:
        print(f"Warning: UI automation not available: {e}")
        _uia_available = False
        _uia = None
        _finder = None

    # Import Florence-2 vision tools (for OCR and screen understanding)
    try:
        import florence_vision
        _florence_available = True
    except ImportError as e:
        print(f"Warning: Florence-2 vision not available: {e}")
        _florence_available = False
        florence_vision = None

# Global state for enhance mode
_enhance_enabled = False

# Default florence availability (set in import block)
if "--help" in sys.argv or "-h" in sys.argv:
    _florence_available = False
    florence_vision = None

# Create MCP server
mcp = FastMCP(
    "Computer Use",
    instructions="""This MCP server provides computer control capabilities.
You can take screenshots, click, type, press keys, and move the mouse.

IMPORTANT - USE UI AUTOMATION FIRST:
Before trying to visually identify icons or elements in screenshots, use these tools:
- get_ui_state() - Get all windows and taskbar apps with exact coordinates
- get_taskbar_apps() - List all taskbar items with click coordinates
- click_taskbar_app("name") - Click a taskbar app by name (much more reliable!)
- get_all_windows() - List all open windows
- find_and_click_window("title") - Find and focus a window by title

These use the Windows UI Automation API and are FAR more reliable than trying
to visually identify icons in screenshots. Use them whenever possible!

AI VISION TOOLS (Florence-2):
- ocr_screen() - Extract all text visible on screen (great for verification!)
- describe_screen() - Get AI description of what's on screen
- verify_text_on_screen("expected text") - Check if specific text is visible

Use these to VERIFY actions worked instead of just assuming!

COORDINATE SYSTEM:
- All coordinates are in native screen pixels
- Origin (0, 0) is at the TOP-LEFT corner
- X increases going RIGHT
- Y increases going DOWN
- Screen is 3840x2160 (4K)

SCREENSHOT vs ZOOM vs ENHANCE:
- Use 'screenshot' for overview of the full screen (scaled down for efficiency)
- Use 'zoom' when you need to read small text or precisely identify click targets
- Use 'enhance' when visual elements are hard to distinguish due to low contrast

COORDINATE CONSISTENCY:
- screenshot: Image is scaled but coordinates map 1:1 to native screen
- zoom/enhance: Returns exact region you request. To click something at pixel (px, py)
  within the zoomed/enhanced image, calculate: click_x = region_x + px, click_y = region_y + py

EFFECTIVE USE OF ZOOM - SEARCHING FOR THINGS:
- Don't assume you'll find what you need in one zoom - be prepared to pan around
- If you don't see what you're looking for, try zooming a DIFFERENT region
- Start wide (e.g., 1500-2000px) to get context, then narrow down
- If text is still hard to read after zooming, try a smaller region or enable enhance
- THINK ABOUT THE SHAPE: taskbar is horizontal (use wide rectangles like 2000x70),
  menus are vertical (use tall rectangles), dialogs are roughly square
- TO ZOOM OUT: use full screenshot() to regain context and orientation
- Alternate between screenshot (overview) and zoom (detail) as needed

SYSTEMATIC PANNING - DON'T GIVE UP:
- If you don't find what you're looking for, PAN in all directions: LEFT, RIGHT, UP, DOWN
- Keep panning until you've covered the entire possible area
- Never assume you've looked "far enough" - keep panning until you find it or exhaust the space
- If lost, take a full screenshot() to "zoom out" and reorient yourself
- Track which regions you've already checked to avoid redundant searches

BEFORE CLICKING - VERIFY YOUR TARGET:
- Use get_mouse_position() to check current position
- Move mouse to target with mouse_move(), then take screenshot to verify position
- After clicking: zoom on affected area to VERIFY the action worked (don't assume)

COMMON MISTAKES TO AVOID:
- Clicking near the top of a window title bar can hit minimize/maximize/close buttons
- The minimize button is usually a dash/line icon near the top-right of windows
- If you meant to click a tab but the window vanished, check the taskbar
- Looking only at the LEFT side of taskbar - running apps are often further RIGHT
- Using the same zoom region repeatedly - pan around to explore"""
)


def get_screen_info() -> dict:
    """Get information about the primary monitor."""
    with mss.mss() as sct:
        monitor = sct.monitors[1]
        return {
            "width": monitor["width"],
            "height": monitor["height"],
            "left": monitor["left"],
            "top": monitor["top"]
        }


def apply_enhancement(img: "Image.Image") -> "Image.Image":
    """Apply contrast enhancement pipeline to an image."""
    # 1. Auto-contrast: stretches histogram to use full range
    img = ImageOps.autocontrast(img, cutoff=0.5)
    # 2. Boost contrast for extra punch
    img = ImageEnhance.Contrast(img).enhance(1.3)
    # 3. Sharpen to make edges and text crisper
    img = img.filter(ImageFilter.SHARPEN)
    # 4. Slight saturation boost to make colors more distinct
    img = ImageEnhance.Color(img).enhance(1.2)
    return img


def draw_cursor_marker(img: "Image.Image", x: int, y: int, scale: float = 1.0) -> "Image.Image":
    """Draw a visible cursor marker on the image.

    Args:
        img: PIL Image to draw on
        x: Native screen X coordinate of cursor
        y: Native screen Y coordinate of cursor
        scale: Scale factor if image was resized (e.g., 0.5 if scaled to half)

    Returns:
        Image with cursor marker drawn
    """
    # Adjust coordinates for scaling
    draw_x = int(x * scale)
    draw_y = int(y * scale)

    # Make a copy to draw on
    img = img.copy()
    draw = ImageDraw.Draw(img)

    # Draw a bright green filled square for maximum visibility
    size = int(25 * scale) if scale < 1 else 25

    # Bright green filled square - maximally visible
    draw.rectangle(
        [(draw_x - size, draw_y - size), (draw_x + size, draw_y + size)],
        fill="#00FF00",  # Pure bright green
        outline="#00FF00"
    )

    return img


def save_screenshot(img: "Image.Image", mode: str, enhanced: bool, region: dict = None) -> tuple:
    """Save screenshot to disk and return the filename and relative path.

    Args:
        img: PIL Image to save
        mode: 'full' or 'zoom'
        enhanced: Whether enhancement was applied
        region: For zoom mode, dict with x, y, width, height

    Returns:
        Tuple of (filename, relative_path including session)
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    enhance_suffix = "_enhanced" if enhanced else ""

    if mode == "zoom" and region:
        filename = f"{timestamp}_zoom_{region['x']}x{region['y']}_{region['width']}x{region['height']}{enhance_suffix}.png"
    else:
        filename = f"{timestamp}_full{enhance_suffix}.png"

    filepath = SCREENSHOTS_DIR / filename
    img.save(filepath, format="PNG", optimize=True)
    relative_path = f"screenshots/{SESSION_ID}/{filename}"
    return (filename, relative_path)


def capture_screenshot(max_dimension: int = 1920, force_enhance: bool = None, draw_cursor: bool = True) -> tuple:
    """Capture screenshot and return image + base64 PNG.

    Args:
        max_dimension: Max width/height before scaling
        force_enhance: Override enhance mode (None = use global setting)
        draw_cursor: Whether to draw cursor marker on image

    Returns:
        Tuple of (PIL Image, base64 string, enhanced bool)
    """
    global _enhance_enabled
    should_enhance = force_enhance if force_enhance is not None else _enhance_enabled

    # Get cursor position before capturing
    cursor_pos = pyautogui.position()

    with mss.mss() as sct:
        monitor = sct.monitors[1]
        screenshot = sct.grab(monitor)
        img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

        # Apply enhancement if enabled
        if should_enhance:
            img = apply_enhancement(img)

        # Scale down if needed
        scale = 1.0
        max_actual = max(img.width, img.height)
        if max_actual > max_dimension:
            scale = max_dimension / max_actual
            new_size = (int(img.width * scale), int(img.height * scale))
            img = img.resize(new_size, Image.Resampling.LANCZOS)

        # Draw cursor marker
        if draw_cursor:
            img = draw_cursor_marker(img, cursor_pos.x, cursor_pos.y, scale)

        buffer = io.BytesIO()
        img.save(buffer, format="PNG", optimize=True)
        img_b64 = base64.standard_b64encode(buffer.getvalue()).decode("utf-8")
        return (img, img_b64, should_enhance)


@mcp.tool()
def screenshot() -> str:
    """
    Take a screenshot of the current screen.
    Returns base64-encoded PNG image data.

    The image is scaled down for efficiency but coordinates map 1:1 to native screen.
    Use this for getting an overview of the screen state.
    Use 'zoom' instead when you need to read small text or precisely identify targets.

    Note: If enhance mode is ON (via set_enhance_mode), contrast enhancement is applied.
    Screenshots are also saved to the screenshots/ directory for user inspection.
    """
    try:
        screen = get_screen_info()
        img, img_b64, enhanced = capture_screenshot()

        # Save to disk for user inspection
        filename, relative_path = save_screenshot(img, "full", enhanced)

        enhance_status = " [ENHANCED]" if enhanced else ""
        return (
            f"Screenshot captured ({screen['width']}x{screen['height']}){enhance_status}.\n"
            f"Saved: {relative_path}\n"
            f"Base64 PNG data: {img_b64[:100]}... (truncated, {len(img_b64)} chars total)"
        )
    except Exception as e:
        return f"Error capturing screenshot: {e}"


@mcp.tool()
def zoom(x: int, y: int, width: int, height: int) -> str:
    """
    Capture a specific region of the screen at NATIVE resolution (no scaling).

    Use this when you need to:
    - Read small text or UI labels precisely
    - Identify exact click targets in dense/crowded UIs
    - Verify element positions before clicking

    Args:
        x: Left edge of region (native screen pixels, 0 = left edge)
        y: Top edge of region (native screen pixels, 0 = top edge)
        width: Width of region to capture
        height: Height of region to capture

    COORDINATE MAPPING:
    The returned image shows exactly the region you requested at native resolution.
    If you see something at pixel position (px, py) within the zoomed image,
    its native screen coordinate is: (x + px, y + py)

    Example: zoom(100, 50, 400, 300) captures a 400x300 region.
    Something at pixel (150, 100) in that image is at screen coordinate (250, 150).

    Note: If enhance mode is ON (via set_enhance_mode), contrast enhancement is applied.
    Screenshots are also saved to the screenshots/ directory for user inspection.
    """
    try:
        global _enhance_enabled
        screen = get_screen_info()

        # Clamp region to screen bounds
        x = max(0, min(x, screen["width"] - 1))
        y = max(0, min(y, screen["height"] - 1))
        width = min(width, screen["width"] - x)
        height = min(height, screen["height"] - y)

        if width <= 0 or height <= 0:
            return f"Error: Invalid region dimensions after clamping"

        # Get cursor position before capturing
        cursor_pos = pyautogui.position()

        # Capture the specific region
        with mss.mss() as sct:
            monitor = sct.monitors[1]
            region = {
                "left": monitor["left"] + x,
                "top": monitor["top"] + y,
                "width": width,
                "height": height
            }
            screenshot_data = sct.grab(region)
            img = Image.frombytes("RGB", screenshot_data.size, screenshot_data.bgra, "raw", "BGRX")

        # Apply enhancement if enabled
        if _enhance_enabled:
            img = apply_enhancement(img)

        # Draw cursor marker if cursor is within the zoomed region
        cursor_in_region = (x <= cursor_pos.x < x + width and y <= cursor_pos.y < y + height)
        if cursor_in_region:
            # Adjust cursor position to be relative to the region
            relative_cursor_x = cursor_pos.x - x
            relative_cursor_y = cursor_pos.y - y
            img = draw_cursor_marker(img, relative_cursor_x, relative_cursor_y, scale=1.0)

        # Save to disk for user inspection
        filename, relative_path = save_screenshot(img, "zoom", _enhance_enabled, {"x": x, "y": y, "width": width, "height": height})

        # No scaling - return at native resolution
        buffer = io.BytesIO()
        img.save(buffer, format="PNG", optimize=True)
        img_b64 = base64.standard_b64encode(buffer.getvalue()).decode("utf-8")

        enhance_status = " [ENHANCED]" if _enhance_enabled else ""
        return (
            f"Zoomed region captured at native resolution{enhance_status}.\n"
            f"Region: x={x}, y={y}, width={width}, height={height}\n"
            f"Saved: {relative_path}\n"
            f"To click something at pixel (px, py) in this image, use: click({x} + px, {y} + py)\n"
            f"Base64 PNG data: {img_b64[:100]}... (truncated, {len(img_b64)} chars total)"
        )
    except Exception as e:
        return f"Error capturing zoom region: {e}"


@mcp.tool()
def set_enhance_mode(enabled: bool) -> str:
    """
    Turn contrast enhancement mode ON or OFF.

    When enhance mode is ON, all screenshots and zoomed regions will have
    contrast enhancement applied automatically. This persists until you
    turn it OFF.

    TURN ON (enabled=True) WHEN YOU EXPERIENCE:
    - Text that's hard to read (gray on white, light on light, dark on dark)
    - UI elements that seem to blend into the background
    - Borders, dividers, or separators you can barely see
    - Dark mode interfaces where you can't distinguish elements
    - Disabled/inactive states that look almost invisible
    - Selection highlights or focus indicators you might be missing
    - Subtle icons, checkmarks, or visual indicators

    TURN OFF (enabled=False) WHEN:
    - You can see everything clearly now
    - The enhancement is making things look weird or over-processed
    - You're done with the low-contrast area
    - Colors look too saturated or unnatural

    Args:
        enabled: True to turn ON enhancement, False to turn OFF

    Enhancement applies automatic contrast stretching, sharpening, and
    slight saturation boost to make visual differences more pronounced.
    """
    global _enhance_enabled
    _enhance_enabled = enabled
    status = "ON" if enabled else "OFF"
    return f"Enhance mode is now {status}. All subsequent screenshots and zooms will {'have' if enabled else 'NOT have'} contrast enhancement applied."


@mcp.tool()
def list_screenshots(limit: int = 10) -> str:
    """
    List recent screenshots from the current session.

    Use this when you want to:
    - Review what you've captured so far
    - Find a previous screenshot to re-examine
    - Check what regions you've already zoomed into

    Args:
        limit: Maximum number of screenshots to list (default: 10, most recent first)

    Returns list of filenames with metadata (timestamp, type, region if zoom, enhanced status).
    """
    try:
        files = sorted(SCREENSHOTS_DIR.glob("*.png"), key=lambda f: f.stat().st_mtime, reverse=True)
        files = files[:limit]

        if not files:
            return f"No screenshots in current session ({SESSION_ID})"

        result = [f"Recent screenshots (session: {SESSION_ID}):"]
        for f in files:
            name = f.name
            # Parse filename for metadata
            parts = name.replace(".png", "").split("_")
            timestamp = f"{parts[0]}_{parts[1]}" if len(parts) >= 2 else "unknown"
            enhanced = "_enhanced" in name
            is_zoom = "_zoom_" in name

            if is_zoom:
                # Extract region info from filename like: timestamp_zoom_100x200_400x300_enhanced.png
                result.append(f"  {name} [ZOOM]{' [ENHANCED]' if enhanced else ''}")
            else:
                result.append(f"  {name} [FULL]{' [ENHANCED]' if enhanced else ''}")

        return "\n".join(result)
    except Exception as e:
        return f"Error listing screenshots: {e}"


@mcp.tool()
def view_screenshot(filename: str) -> str:
    """
    Re-load and return a previous screenshot from this session.

    Use this when you want to:
    - Re-examine a screenshot you took earlier without taking a new one
    - Compare what the screen looked like before vs after an action
    - Look at a zoomed region again to find something you missed

    Args:
        filename: The filename of the screenshot (from list_screenshots output)

    Returns the base64-encoded PNG data of the screenshot.
    """
    try:
        filepath = SCREENSHOTS_DIR / filename
        if not filepath.exists():
            # Try other sessions if not in current
            for session_dir in SCREENSHOTS_BASE.iterdir():
                if session_dir.is_dir():
                    alt_path = session_dir / filename
                    if alt_path.exists():
                        filepath = alt_path
                        break

        if not filepath.exists():
            return f"Screenshot not found: {filename}"

        with open(filepath, "rb") as f:
            img_data = f.read()
        img_b64 = base64.standard_b64encode(img_data).decode("utf-8")

        # Parse region info from filename if it's a zoom
        region_info = ""
        if "_zoom_" in filename:
            # Parse: timestamp_zoom_Xx_Y_WxH_enhanced.png
            try:
                parts = filename.replace(".png", "").replace("_enhanced", "").split("_zoom_")[1]
                coords = parts.split("_")
                if len(coords) >= 2:
                    xy = coords[0].split("x")
                    wh = coords[1].split("x")
                    x, y = int(xy[0]), int(xy[1])
                    w, h = int(wh[0]), int(wh[1])
                    region_info = f"\nRegion: x={x}, y={y}, width={w}, height={h}\nTo click at (px,py) in image: click({x}+px, {y}+py)"
            except:
                pass

        return (
            f"Loaded screenshot: {filename}{region_info}\n"
            f"Base64 PNG data: {img_b64[:100]}... (truncated, {len(img_b64)} chars total)"
        )
    except Exception as e:
        return f"Error loading screenshot: {e}"


@mcp.tool()
def get_screen_size() -> dict:
    """
    Get the dimensions of the primary screen.
    Returns width and height in pixels.
    """
    return get_screen_info()


@mcp.tool()
def left_click(x: int, y: int) -> str:
    """
    Perform a left mouse click at the specified coordinates.

    Args:
        x: X coordinate (pixels from left edge)
        y: Y coordinate (pixels from top edge)
    """
    try:
        pyautogui.click(x, y)
        return f"Left clicked at ({x}, {y})"
    except Exception as e:
        return f"Error clicking: {e}"


@mcp.tool()
def right_click(x: int, y: int) -> str:
    """
    Perform a right mouse click at the specified coordinates.

    Args:
        x: X coordinate (pixels from left edge)
        y: Y coordinate (pixels from top edge)
    """
    try:
        pyautogui.rightClick(x, y)
        return f"Right clicked at ({x}, {y})"
    except Exception as e:
        return f"Error right-clicking: {e}"


@mcp.tool()
def double_click(x: int, y: int) -> str:
    """
    Perform a double left click at the specified coordinates.

    Args:
        x: X coordinate (pixels from left edge)
        y: Y coordinate (pixels from top edge)
    """
    try:
        pyautogui.doubleClick(x, y)
        return f"Double clicked at ({x}, {y})"
    except Exception as e:
        return f"Error double-clicking: {e}"


@mcp.tool()
def mouse_move(x: int, y: int) -> str:
    """
    Move the mouse cursor to the specified coordinates.

    Args:
        x: X coordinate (pixels from left edge)
        y: Y coordinate (pixels from top edge)
    """
    try:
        pyautogui.moveTo(x, y)
        return f"Moved mouse to ({x}, {y})"
    except Exception as e:
        return f"Error moving mouse: {e}"


@mcp.tool()
def drag(start_x: int, start_y: int, end_x: int, end_y: int) -> str:
    """
    Click and drag from start coordinates to end coordinates.

    Args:
        start_x: Starting X coordinate
        start_y: Starting Y coordinate
        end_x: Ending X coordinate
        end_y: Ending Y coordinate
    """
    try:
        pyautogui.moveTo(start_x, start_y)
        pyautogui.drag(end_x - start_x, end_y - start_y, duration=0.3)
        return f"Dragged from ({start_x}, {start_y}) to ({end_x}, {end_y})"
    except Exception as e:
        return f"Error dragging: {e}"


@mcp.tool()
def type_text(text: str) -> str:
    """
    Type the specified text using the keyboard.
    Note: For special characters or non-ASCII, use key() instead.

    Args:
        text: The text to type
    """
    try:
        pyautogui.typewrite(text, interval=0.02)
        return f"Typed: {text[:50]}{'...' if len(text) > 50 else ''}"
    except Exception as e:
        return f"Error typing: {e}"


@mcp.tool()
def type_unicode(text: str) -> str:
    """
    Type text that may contain unicode/special characters.
    Slower but supports all characters.

    Args:
        text: The text to type (can include unicode)
    """
    try:
        pyautogui.write(text)
        return f"Typed (unicode): {text[:50]}{'...' if len(text) > 50 else ''}"
    except Exception as e:
        return f"Error typing unicode: {e}"


@mcp.tool()
def key(keys: str) -> str:
    """
    Press a key or key combination.

    Args:
        keys: Key name or combination (e.g., "enter", "ctrl+s", "alt+tab", "win+r")

    Common keys: enter, tab, escape, backspace, delete, space,
                 up, down, left, right, home, end, pageup, pagedown,
                 f1-f12, ctrl, alt, shift, win
    """
    try:
        key_list = keys.lower().split("+")
        if len(key_list) > 1:
            pyautogui.hotkey(*key_list)
        else:
            pyautogui.press(keys)
        return f"Pressed key(s): {keys}"
    except Exception as e:
        return f"Error pressing key: {e}"


@mcp.tool()
def scroll(x: int, y: int, direction: str, amount: int = 3) -> str:
    """
    Scroll at the specified coordinates.

    Args:
        x: X coordinate to scroll at
        y: Y coordinate to scroll at
        direction: Direction to scroll ("up", "down", "left", "right")
        amount: Number of scroll units (default: 3)
    """
    try:
        pyautogui.moveTo(x, y)
        if direction in ("up", "down"):
            scroll_amount = amount if direction == "up" else -amount
            pyautogui.scroll(scroll_amount)
        else:
            scroll_amount = amount if direction == "right" else -amount
            pyautogui.hscroll(scroll_amount)
        return f"Scrolled {direction} by {amount} at ({x}, {y})"
    except Exception as e:
        return f"Error scrolling: {e}"


@mcp.tool()
def wait(seconds: float = 1.0) -> str:
    """
    Wait for the specified number of seconds.
    Useful after actions that trigger UI changes.

    Args:
        seconds: Number of seconds to wait (default: 1.0)
    """
    time.sleep(seconds)
    return f"Waited {seconds} seconds"


@mcp.tool()
def get_mouse_position() -> dict:
    """
    Get the current mouse cursor position.
    Returns x and y coordinates.
    """
    pos = pyautogui.position()
    return {"x": pos.x, "y": pos.y}


# =============================================================================
# UI AUTOMATION TOOLS - More reliable than visual detection
# =============================================================================

@mcp.tool()
def get_taskbar_apps() -> str:
    """
    Get all applications shown in the Windows taskbar.

    This uses the Windows UI Automation API to directly query the taskbar,
    providing exact names and click coordinates for each app.

    Returns a list of all taskbar items (Start, Search, running apps, system tray).
    Each item includes the name and exact coordinates to click.

    USE THIS instead of trying to visually identify icons in screenshots!
    """
    if not _uia_available or not _uia:
        return "Error: UI Automation not available. Install pywinauto: pip install pywinauto"

    try:
        apps = _uia.get_taskbar_apps()
        if not apps:
            return "No taskbar apps found"

        result = ["Taskbar Applications:"]
        for app in apps:
            result.append(f"  - '{app.name}' -> click at ({app.center_x}, {app.center_y})")

        return "\n".join(result)
    except Exception as e:
        return f"Error getting taskbar apps: {e}"


@mcp.tool()
def click_taskbar_app(app_name: str) -> str:
    """
    Click on a taskbar application by name.

    This is MUCH more reliable than trying to visually identify and click icons.
    Uses Windows UI Automation to find the exact app location.

    Args:
        app_name: Part of the app name to search for (case-insensitive).
                  Examples: "Terminal", "Firefox", "Chrome", "Explorer"

    The app_name can be partial - it will match if the taskbar button text
    contains your search term. For example, "Terminal" will match
    "Terminal - 2 running windows".
    """
    if not _uia_available or not _uia:
        return "Error: UI Automation not available. Install pywinauto: pip install pywinauto"

    try:
        app = _uia.find_taskbar_app(app_name)
        if not app:
            # List available apps to help user
            apps = _uia.get_taskbar_apps()
            app_names = [a.name for a in apps if a.name]
            return f"App '{app_name}' not found in taskbar.\nAvailable apps: {app_names}"

        # Click on the app
        pyautogui.click(app.center_x, app.center_y)
        return f"Clicked on '{app.name}' at ({app.center_x}, {app.center_y})"

    except Exception as e:
        return f"Error clicking taskbar app: {e}"


@mcp.tool()
def get_all_windows() -> str:
    """
    Get all visible windows on the desktop.

    Returns window titles and their screen positions.
    Use this to understand what's currently open and where windows are located.
    """
    if not _uia_available or not _uia:
        return "Error: UI Automation not available. Install pywinauto: pip install pywinauto"

    try:
        windows = _uia.get_all_windows()
        if not windows:
            return "No windows found"

        result = ["Open Windows:"]
        for win in windows:
            # Skip tiny windows or off-screen windows
            if win.width > 50 and win.height > 50 and win.x >= -2000:
                result.append(f"  - '{win.name[:60]}' at ({win.x}, {win.y}) size {win.width}x{win.height}")

        return "\n".join(result)
    except Exception as e:
        return f"Error getting windows: {e}"


@mcp.tool()
def find_and_click_window(title: str) -> str:
    """
    Find a window by title and click to focus it.

    Args:
        title: Part of the window title to search for (case-insensitive)

    This brings the window to the foreground by clicking on it.
    """
    if not _uia_available or not _uia:
        return "Error: UI Automation not available. Install pywinauto: pip install pywinauto"

    try:
        window = _uia.find_window(title_contains=title)
        if not window:
            # List available windows to help
            windows = _uia.get_all_windows()
            titles = [w.name[:50] for w in windows if w.name]
            return f"Window with title containing '{title}' not found.\nAvailable windows: {titles}"

        # Click to focus
        pyautogui.click(window.center_x, window.center_y)
        return f"Clicked on window '{window.name}' at ({window.center_x}, {window.center_y})"

    except Exception as e:
        return f"Error finding window: {e}"


@mcp.tool()
def get_ui_state() -> str:
    """
    Get a comprehensive UI state including all windows and taskbar apps.

    This provides structured information about the current desktop state
    without needing to analyze screenshots visually.

    Returns:
        - All visible windows with titles and positions
        - All taskbar applications with click coordinates

    Use this FIRST when you need to understand what's on screen,
    instead of trying to parse screenshots visually.
    """
    if not _uia_available or not _uia:
        return "Error: UI Automation not available. Install pywinauto: pip install pywinauto"

    try:
        result = ["=" * 50, "CURRENT UI STATE", "=" * 50]

        # Windows
        result.append("\n--- OPEN WINDOWS ---")
        windows = _uia.get_all_windows()
        for win in windows:
            if win.width > 50 and win.height > 50 and win.x >= -2000:
                result.append(f"  [{win.name[:50]}]")
                result.append(f"    Position: ({win.x}, {win.y}) Size: {win.width}x{win.height}")
                result.append(f"    Click center: ({win.center_x}, {win.center_y})")

        # Taskbar
        result.append("\n--- TASKBAR APPS ---")
        apps = _uia.get_taskbar_apps()
        for app in apps:
            result.append(f"  '{app.name}' -> ({app.center_x}, {app.center_y})")

        result.append("\n" + "=" * 50)
        return "\n".join(result)

    except Exception as e:
        return f"Error getting UI state: {e}"


# =============================================================================
# FLORENCE-2 VISION TOOLS - AI-powered screen understanding
# =============================================================================

@mcp.tool()
def ocr_screen() -> str:
    """
    Extract all visible text from the current screen using Florence-2 OCR.

    This uses a local AI model (Florence-2) to perform optical character recognition.
    Much more accurate than BLIP for reading UI text, webpage content, etc.

    Use this to:
    - Verify what page/application is currently displayed
    - Read text that might be hard to see in screenshots
    - Check if expected text is visible on screen after an action

    Returns the extracted text from the screen.
    Note: First call will take longer as the model loads into GPU memory.
    """
    if not _florence_available or not florence_vision:
        return "Error: Florence-2 vision not available. Check that the venv is active and dependencies are installed."

    try:
        # Capture screen
        img = florence_vision.capture_screen()
        # Resize for faster processing (Florence-2 doesn't need full 4K)
        img.thumbnail((1280, 1280), Image.Resampling.LANCZOS)

        # Run OCR
        text = florence_vision.ocr_screenshot(img)
        return f"OCR Result:\n{text}"
    except Exception as e:
        return f"Error performing OCR: {e}"


@mcp.tool()
def describe_screen() -> str:
    """
    Get an AI-generated description of what's currently on screen.

    Uses Florence-2 to analyze and describe the screen content.
    Useful for understanding the overall state of the UI.

    Returns a text description of the screen.
    Note: First call will take longer as the model loads into GPU memory.
    """
    if not _florence_available or not florence_vision:
        return "Error: Florence-2 vision not available. Check that the venv is active and dependencies are installed."

    try:
        # Capture screen
        img = florence_vision.capture_screen()
        img.thumbnail((1024, 1024), Image.Resampling.LANCZOS)

        # Get detailed caption
        description = florence_vision.detailed_caption(img)
        return f"Screen Description:\n{description}"
    except Exception as e:
        return f"Error describing screen: {e}"


@mcp.tool()
def verify_text_on_screen(expected_text: str) -> str:
    """
    Check if specific text is visible on the current screen.

    Uses Florence-2 OCR to extract all text, then searches for the expected text.
    Case-insensitive search.

    Args:
        expected_text: The text to look for on screen

    Returns whether the text was found and context around it if present.

    Use this to verify:
    - Page navigation succeeded (check for page title/URL)
    - Dialog appeared (check for dialog text)
    - Action completed (check for confirmation message)
    """
    if not _florence_available or not florence_vision:
        return "Error: Florence-2 vision not available. Check that the venv is active and dependencies are installed."

    try:
        # Capture screen
        img = florence_vision.capture_screen()
        img.thumbnail((1280, 1280), Image.Resampling.LANCZOS)

        # Run OCR
        ocr_text = florence_vision.ocr_screenshot(img)

        # Search for expected text (case-insensitive)
        if expected_text.lower() in ocr_text.lower():
            # Find context around the match
            lower_ocr = ocr_text.lower()
            lower_expected = expected_text.lower()
            pos = lower_ocr.find(lower_expected)
            start = max(0, pos - 50)
            end = min(len(ocr_text), pos + len(expected_text) + 50)
            context = ocr_text[start:end]

            return f"FOUND: '{expected_text}' is visible on screen.\nContext: ...{context}..."
        else:
            return f"NOT FOUND: '{expected_text}' was not detected on screen.\nOCR extracted {len(ocr_text)} characters. Sample: {ocr_text[:200]}..."

    except Exception as e:
        return f"Error verifying text: {e}"


def main():
    parser = argparse.ArgumentParser(description="Computer Use MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "streamable-http"],
        default="stdio",
        help="Transport to use (default: stdio)"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8080,
        help="Port for HTTP transport (default: 8080)"
    )
    args = parser.parse_args()

    if args.transport == "streamable-http":
        mcp.settings.host = "127.0.0.1"
        mcp.settings.port = args.port

    mcp.run(transport=args.transport)


if __name__ == "__main__":
    main()
