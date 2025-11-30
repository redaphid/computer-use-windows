# Computer Use MCP Server

An MCP server that provides computer control capabilities for Claude Code. Enables Claude to interact with the desktop through screenshots, mouse, keyboard, OCR, and Win32 APIs.

## Workflow

1. Use `ocr_screen()` to find text and get click coordinates
2. Use `left_click(x, y)` with the coordinates from OCR
3. Use `verify_text_on_screen("text")` to confirm actions worked

For window management, prefer the Win32 tools over clicking X buttons.

## Tools Available

### OCR Tools (Primary Method)
- `ocr_screen()` - Extract visible text with click coordinates
- `verify_text_on_screen(text)` - Find specific text and get coordinates
- `set_enhance_mode(true/false)` - Enable/disable contrast enhancement for dark UIs

### Win32 Window Management (Most Reliable)
- `close_window(title)` - Close window via WM_CLOSE message
- `focus_window(title)` - Bring window to foreground
- `minimize_window(title)` - Minimize window
- `maximize_window(title)` - Maximize window
- `launch_app(name)` - Launch app via PowerShell
- `windows_search(query)` - Open Windows search and type query

### Visual
- `screenshot` - Capture full screen (scaled, enhanced if enabled)
- `zoom(x, y, w, h)` - Capture region at native resolution
- `describe_screen()` - AI description of screen content

### Input
- `left_click`, `right_click`, `double_click` - Mouse clicks at coordinates
- `mouse_move` - Move cursor to coordinates
- `drag` - Click and drag between coordinates
- `type_text`, `type_unicode` - Keyboard input
- `key` - Press keys/key combinations (e.g., "ctrl+s", "alt+f4")
- `scroll` - Scroll at coordinates

### Windows Info
- `get_all_windows()` - List open windows with positions
- `find_and_click_window(title)` - Focus window by clicking
- `get_ui_state()` - Get all windows with coordinates
- `get_mouse_position` - Current cursor position
- `get_screen_size` - Screen dimensions

### Journal (Cross-Session Learning)
- `journal_write(observation, tags)` - Record what worked/didn't work
- `journal_query(search_term)` - Query past observations

## Coordinate System

- Origin (0, 0) is top-left
- X increases right, Y increases down
- Screen: 3840x2160 (4K)
- `screenshot` scales but coordinates map 1:1 to native

## Best Practices

### Use Win32 for Window Management
```
close_window("Firefox")      # More reliable than clicking X
focus_window("Steam")        # Bring to front
minimize_window("Discord")   # Minimize
```

### Use OCR for Finding UI Elements
```
ocr_screen()                 # Get all text with coordinates
# Output: 'Save' -> click(1234, 567)
left_click(1234, 567)        # Click on Save
verify_text_on_screen("Saved")  # Confirm it worked
```

### Record What Works
```
journal_write("Alt+F4 reliably closes most windows", "close,keyboard")
journal_query("close")       # Find past learnings
```

## Dependencies

- mss, pyautogui, Pillow, mcp
- pywin32 (Win32 API)
- PaddleOCR (text extraction)
- torch, transformers (Florence-2 vision)

Run with: `.venv\Scripts\python.exe computer_use_mcp.py`
