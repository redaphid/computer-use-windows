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

    # Import Win32 APIs for better Windows integration
    try:
        import win32gui
        import win32con
        import win32process
        import subprocess
        _win32_available = True
    except ImportError as e:
        print(f"Warning: Win32 API not available: {e}")
        _win32_available = False

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
    instructions="""Computer control via screenshots, mouse, keyboard, and OCR.

WORKFLOW:
1. Use ocr_screen() to find text and get click coordinates
2. Use left_click(x, y) with the coordinates from OCR
3. Use verify_text_on_screen("text") to confirm actions worked

OCR TOOLS:
- ocr_screen() - Returns all visible text with click coordinates
- verify_text_on_screen("text") - Find specific text and get its coordinates
- set_enhance_mode(true) - Enable contrast enhancement for better OCR on dark UIs

VISUAL TOOLS:
- screenshot() - Full screen overview (scaled)
- zoom(x, y, w, h) - Native resolution region capture
- describe_screen() - AI description of screen content

UI AUTOMATION:
- get_all_windows() - List open windows with positions
- find_and_click_window("title") - Focus a window by title
- get_ui_state() - Get all windows with coordinates

INPUT:
- left_click(x, y), right_click(x, y), double_click(x, y)
- type_text("text"), key("ctrl+s")
- scroll(x, y, "up"/"down")

COORDINATES: Origin (0,0) is top-left. X increases right, Y increases down."""
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
    Get all open windows with their positions and click coordinates.

    Returns window titles, positions, sizes, and center coordinates for clicking.
    Use this to find and interact with application windows.
    """
    if not _uia_available or not _uia:
        return "Error: UI Automation not available. Install pywinauto: pip install pywinauto"

    try:
        result = ["Open Windows:"]
        windows = _uia.get_all_windows()
        for win in windows:
            if win.width > 50 and win.height > 50 and win.x >= -2000:
                result.append(f"  '{win.name[:60]}' at ({win.x}, {win.y}) {win.width}x{win.height} -> click({win.center_x}, {win.center_y})")

        return "\n".join(result)

    except Exception as e:
        return f"Error getting UI state: {e}"


# =============================================================================
# FLORENCE-2 VISION TOOLS - AI-powered screen understanding
# =============================================================================

@mcp.tool()
def ocr_screen() -> str:
    """
    Extract all visible text from the current screen using PaddleOCR.

    Returns text with click coordinates. Each line shows:
      'detected text' -> click(x, y)

    To click on detected text, use the coordinates directly with left_click(x, y).

    If enhance mode is ON, contrast enhancement is applied before OCR
    which can help with low-contrast text.
    """
    global _enhance_enabled

    if not _florence_available or not florence_vision:
        return "Error: Vision module not available. Check that the venv is active and dependencies are installed."

    try:
        # Capture screen at full resolution for accurate coordinates
        img = florence_vision.capture_screen()

        # Apply enhancement if enabled (helps with low-contrast text)
        if _enhance_enabled:
            img = apply_enhancement(img)

        # Run OCR with regions to get coordinates
        regions = florence_vision.ocr_with_regions(img)

        if not regions:
            return "OCR Result: No text detected on screen"

        # Format output with click coordinates
        result = ["OCR Result (text -> click coordinates):"]
        for r in regions:
            bbox = r['bbox']
            # Calculate center point for clicking
            center_x = int((bbox[0][0] + bbox[2][0]) / 2)
            center_y = int((bbox[0][1] + bbox[2][1]) / 2)
            text = r['text']
            result.append(f"  '{text}' -> click({center_x}, {center_y})")

        return "\n".join(result)
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
    Check if specific text is visible on the current screen and return its click coordinates.

    Uses PaddleOCR to find text. Case-insensitive partial match.

    Args:
        expected_text: The text to look for on screen

    Returns whether the text was found, with click coordinates if found.

    Use this to verify actions worked (page loaded, dialog appeared, etc.)
    """
    global _enhance_enabled

    if not _florence_available or not florence_vision:
        return "Error: Vision module not available. Check that the venv is active."

    try:
        # Capture screen at full resolution
        img = florence_vision.capture_screen()

        # Apply enhancement if enabled
        if _enhance_enabled:
            img = apply_enhancement(img)

        # Run OCR with regions
        regions = florence_vision.ocr_with_regions(img)

        # Search for expected text (case-insensitive)
        for r in regions:
            if expected_text.lower() in r['text'].lower():
                bbox = r['bbox']
                center_x = int((bbox[0][0] + bbox[2][0]) / 2)
                center_y = int((bbox[0][1] + bbox[2][1]) / 2)
                return f"FOUND: '{r['text']}' -> click({center_x}, {center_y})"

        # Not found - show what was detected
        all_text = [r['text'] for r in regions[:10]]
        return f"NOT FOUND: '{expected_text}' not detected.\nVisible text: {all_text}"

    except Exception as e:
        return f"Error verifying text: {e}"


# =============================================================================
# WINDOWS INTEGRATION - Win32 API for proper window control
# =============================================================================

@mcp.tool()
def close_window(title: str) -> str:
    """
    Close a window by title using Win32 API.

    More reliable than clicking the X button. Sends WM_CLOSE message.

    Args:
        title: Part of the window title (case-insensitive)

    Returns success/failure message.
    """
    if not _win32_available:
        return "Error: Win32 API not available"

    try:
        def callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    results.append((hwnd, window_title))
            return True

        results = []
        win32gui.EnumWindows(callback, results)

        if not results:
            return f"No window found containing '{title}'"

        hwnd, window_title = results[0]
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        return f"Sent close message to '{window_title}'"

    except Exception as e:
        return f"Error closing window: {e}"


@mcp.tool()
def focus_window(title: str) -> str:
    """
    Bring a window to the foreground by title.

    Args:
        title: Part of the window title (case-insensitive)
    """
    if not _win32_available:
        return "Error: Win32 API not available"

    try:
        def callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    results.append((hwnd, window_title))
            return True

        results = []
        win32gui.EnumWindows(callback, results)

        if not results:
            return f"No window found containing '{title}'"

        hwnd, window_title = results[0]
        win32gui.SetForegroundWindow(hwnd)
        return f"Focused '{window_title}'"

    except Exception as e:
        return f"Error focusing window: {e}"


@mcp.tool()
def minimize_window(title: str) -> str:
    """
    Minimize a window by title.

    Args:
        title: Part of the window title (case-insensitive)
    """
    if not _win32_available:
        return "Error: Win32 API not available"

    try:
        def callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    results.append((hwnd, window_title))
            return True

        results = []
        win32gui.EnumWindows(callback, results)

        if not results:
            return f"No window found containing '{title}'"

        hwnd, window_title = results[0]
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        return f"Minimized '{window_title}'"

    except Exception as e:
        return f"Error minimizing window: {e}"


@mcp.tool()
def maximize_window(title: str) -> str:
    """
    Maximize a window by title.

    Args:
        title: Part of the window title (case-insensitive)
    """
    if not _win32_available:
        return "Error: Win32 API not available"

    try:
        def callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    results.append((hwnd, window_title))
            return True

        results = []
        win32gui.EnumWindows(callback, results)

        if not results:
            return f"No window found containing '{title}'"

        hwnd, window_title = results[0]
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        return f"Maximized '{window_title}'"

    except Exception as e:
        return f"Error maximizing window: {e}"


@mcp.tool()
def launch_app(app_name: str) -> str:
    """
    Launch an application using Windows Search/Start menu.

    Args:
        app_name: Name of the app to launch (e.g., "Firefox", "Notepad", "Steam")

    Uses Windows search to find and launch the app.
    """
    try:
        # Use PowerShell Start-Process which handles app resolution
        result = subprocess.run(
            ["powershell", "-Command", f"Start-Process '{app_name}'"],
            capture_output=True,
            text=True,
            timeout=10
        )

        if result.returncode == 0:
            return f"Launched '{app_name}'"
        else:
            # Try shell execute as fallback
            os.startfile(app_name)
            return f"Launched '{app_name}' via shell"

    except FileNotFoundError:
        return f"App '{app_name}' not found. Try the full path or exact name."
    except Exception as e:
        return f"Error launching app: {e}"


@mcp.tool()
def windows_search(query: str) -> str:
    """
    Open Windows Search and type a query.

    Args:
        query: What to search for

    Opens the Start menu search and types the query.
    """
    try:
        # Press Windows key to open Start/Search
        pyautogui.press('win')
        time.sleep(0.3)

        # Type the search query
        pyautogui.typewrite(query, interval=0.02)

        return f"Opened Windows search with query: '{query}'. Press Enter to select first result."

    except Exception as e:
        return f"Error with Windows search: {e}"


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
