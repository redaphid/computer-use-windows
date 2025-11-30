# Computer Use MCP Server

An MCP server that provides computer control capabilities for Claude Code. Enables Claude to interact with the desktop through screenshots, mouse, and keyboard control.

## IMPORTANT: Use UI Automation First!

Before trying to visually identify icons in screenshots, **use the UI Automation tools**. They use the Windows accessibility API to get exact element names and coordinates - much more reliable than visual identification!

```python
# Instead of trying to find the Terminal icon visually:
click_taskbar_app("Terminal")  # Just works!

# Or get all apps first:
get_taskbar_apps()  # Returns names and coordinates for everything
```

## Tools Available

### UI Automation (Most Reliable - Use These First!)
- `get_ui_state()` - Get all windows and taskbar apps with exact coordinates
- `get_taskbar_apps()` - List all taskbar items with names and click coordinates
- `click_taskbar_app(name)` - Click a taskbar app by name (partial match works!)
- `get_all_windows()` - List all open windows with titles and positions
- `find_and_click_window(title)` - Find and focus a window by title

### Visual (Use When UI Automation Isn't Enough)
- `screenshot` - Capture full screen (scaled to max 1920px for efficiency)
- `zoom` - Capture a specific region at native resolution for precise UI inspection
- `set_enhance_mode` - Toggle contrast enhancement ON/OFF (persists across captures)

### Input
- `left_click`, `right_click`, `double_click` - Mouse clicks at coordinates
- `mouse_move` - Move cursor to coordinates
- `drag` - Click and drag between coordinates
- `type_text`, `type_unicode` - Keyboard input
- `key` - Press keys/key combinations
- `scroll` - Scroll at coordinates

### AI Vision (Florence-2 - Use to VERIFY actions!)
- `ocr_screen()` - Extract all visible text using AI OCR (great for verification!)
- `describe_screen()` - Get AI-generated description of screen content
- `verify_text_on_screen(text)` - Check if specific text is visible on screen

**IMPORTANT**: Use these tools to VERIFY that actions worked, instead of just assuming!

### Utility
- `get_mouse_position` - Get current cursor position
- `get_screen_size` - Get screen dimensions
- `wait` - Pause for UI to update
- `list_screenshots` - List recent screenshots from current session
- `view_screenshot` - Re-load a previous screenshot to examine again

## Screenshot Sessions

Screenshots are saved in `screenshots/<session_id>/` folders:
- Each MCP server run gets a unique session ID (timestamp + random)
- Use `list_screenshots()` to see what you've captured
- Use `view_screenshot(filename)` to re-examine without taking a new capture
- Useful for comparing before/after states or finding something you missed

## Coordinate System

All coordinates are in **native screen pixels**:
- Origin (0, 0) is at the **top-left** corner
- X increases going **right**
- Y increases going **down**
- Screen size is 3840x2160 (4K)

**Important**: The `screenshot` tool scales images down for efficiency, but coordinates in the scaled image map directly to native coordinates. When you see something at position (x, y) in a scaled screenshot, use those same coordinates for clicking.

## Using the Zoom Tool

Use `zoom` when you need to:
- Read small text or UI elements precisely
- Identify exact click targets in dense UIs
- Verify the exact position of elements before clicking

The zoom tool returns a region at native resolution without scaling. Specify the region using native screen coordinates (top-left x, y and width, height).

**Coordinate consistency**: The zoom tool returns the exact region you request in native coordinates. If you zoom into region (100, 50, 400, 300), the top-left of the returned image corresponds to screen coordinate (100, 50). To click on something visible at position (px, py) within the zoomed image, calculate: `click_x = region_x + px`, `click_y = region_y + py`.

## Using Enhance Mode

Enhance mode is a **stateful toggle** that applies contrast enhancement to ALL subsequent screenshots and zooms until turned off.

**Turn ON** (`set_enhance_mode(true)`) when:
- Text is hard to read due to low contrast
- UI elements blend into the background
- Working with dark mode interfaces
- Borders/dividers are barely visible
- Selection highlights are hard to see

**Turn OFF** (`set_enhance_mode(false)`) when:
- You can see everything clearly
- Enhancement looks over-processed
- Colors appear too saturated

The enhancement applies: auto-contrast stretching, sharpening, and slight saturation boost.

## Best Practices for Computer Use

### Before Clicking
- **Always verify mouse position** before clicking using `get_mouse_position()`
- Take a screenshot or zoom AFTER moving mouse to confirm you're targeting the right spot

### Verification After Actions
- After clicking to minimize: zoom on the main screen area to VERIFY the window is gone
- After clicking to restore: zoom to VERIFY the window is back
- Don't assume actions worked based on file sizes - actually look at the content

### Full Screenshots vs Zoom
- Use full `screenshot()` to get overall context and orientation
- Use `zoom()` when you need to read text or identify precise click targets
- Alternate between them as needed

## Running the Server

```bash
# For Claude Code (stdio transport)
python computer_use_mcp.py

# For testing (HTTP transport)
python computer_use_mcp.py --transport streamable-http --port 8080
```

## Dependencies

- mss (screenshots)
- pyautogui (mouse/keyboard control)
- Pillow (image processing)
- mcp (MCP server framework)
- pywinauto (Windows UI Automation - highly recommended!)

### For Florence-2 Vision (in .venv)
- torch with CUDA
- transformers<4.46 (pinned for Florence-2 compatibility)
- timm, einops

To use Florence-2 tools, run the MCP server from the venv:
```bash
.venv\Scripts\python.exe computer_use_mcp.py
```
