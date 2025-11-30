# Computer Use MCP Server for Windows

An MCP (Model Context Protocol) server that enables Claude to control your Windows desktop through screenshots, mouse, keyboard, OCR, and Win32 APIs.

## Features

- **Win32 Integration** - Native window management (close, focus, minimize, maximize)
- **Screen Capture** - Screenshots with zoom and contrast enhancement
- **Input Control** - Mouse clicks, keyboard input, drag operations
- **OCR** - PaddleOCR for accurate text extraction with click coordinates
- **AI Vision** - Florence-2 for screen descriptions
- **Journal** - Cross-session learning from past observations

## Quick Start

### 1. Install uv (if not already installed)

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 2. Clone and set up

```powershell
git clone https://github.com/anthropics/computer-use-windows.git
cd computer-use-windows

# Create venv and install dependencies
uv venv --python 3.11
uv pip install torch torchvision --index-url https://download.pytorch.org/whl/cu121
uv pip install "transformers<4.46" timm einops mss pyautogui pillow mcp pywinauto paddleocr paddlepaddle pywin32
```

### 3. Configure Claude Code

The `.mcp.json` file configures Claude Code to use this server:

```json
{
  "mcpServers": {
    "computer-use": {
      "command": ".venv/Scripts/python.exe",
      "args": ["computer_use_mcp.py"]
    }
  }
}
```

### 4. Restart Claude Code

The computer-use tools will now be available.

## Workflow

```
1. ocr_screen()                  # Get text with click coordinates
2. left_click(x, y)              # Click using coordinates from OCR
3. verify_text_on_screen("text") # Confirm action worked
```

For window management, use Win32 tools instead of clicking X buttons.

## Available Tools

### Win32 Window Management (Most Reliable)
| Tool | Description |
|------|-------------|
| `close_window(title)` | Close window via WM_CLOSE |
| `focus_window(title)` | Bring window to foreground |
| `minimize_window(title)` | Minimize window |
| `maximize_window(title)` | Maximize window |
| `launch_app(name)` | Launch app via PowerShell |
| `windows_search(query)` | Open Windows search |

### OCR
| Tool | Description |
|------|-------------|
| `ocr_screen()` | Extract text with click coordinates |
| `verify_text_on_screen(text)` | Find text and get coordinates |
| `set_enhance_mode(bool)` | Toggle contrast enhancement |

### Screen Capture
| Tool | Description |
|------|-------------|
| `screenshot()` | Capture full screen |
| `zoom(x, y, w, h)` | Capture region at native resolution |
| `describe_screen()` | AI description (Florence-2) |

### Input
| Tool | Description |
|------|-------------|
| `left_click(x, y)` | Left click at coordinates |
| `right_click(x, y)` | Right click at coordinates |
| `double_click(x, y)` | Double click at coordinates |
| `type_text(text)` | Type ASCII text |
| `key(keys)` | Press key combination (e.g., "ctrl+s") |
| `scroll(x, y, direction)` | Scroll at position |

### Windows Info
| Tool | Description |
|------|-------------|
| `get_all_windows()` | List open windows with positions |
| `find_and_click_window(title)` | Focus window by clicking |
| `get_ui_state()` | Get all windows with coordinates |

### Journal (Cross-Session Learning)
| Tool | Description |
|------|-------------|
| `journal_write(observation, tags)` | Record what worked/didn't |
| `journal_query(search_term)` | Query past observations |

## Requirements

- Windows 10/11
- Python 3.11
- NVIDIA GPU with CUDA (for Florence-2)
- [uv](https://github.com/astral-sh/uv) package manager

## Architecture

```
computer_use_mcp.py   - Main MCP server with all tools
florence_vision.py    - PaddleOCR + Florence-2 integration
vision_tools.py       - Windows UI Automation helpers
journal.md            - Persistent journal file
```

## License

MIT
