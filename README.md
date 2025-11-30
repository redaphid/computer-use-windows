# Computer Use MCP Server for Windows

An MCP (Model Context Protocol) server that enables Claude to control your Windows desktop through screenshots, mouse, keyboard, and AI-powered screen understanding.

## Features

- **UI Automation** - Windows accessibility API for reliable element detection
- **Screen Capture** - Screenshots with zoom and contrast enhancement
- **Input Control** - Mouse clicks, keyboard input, drag operations
- **OCR** - PaddleOCR for accurate text extraction from screen
- **AI Vision** - Florence-2 for screen descriptions

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
uv pip install "transformers<4.46" timm einops mss pyautogui pillow mcp pywinauto paddleocr paddlepaddle
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

## Available Tools

### UI Automation (Most Reliable)
| Tool | Description |
|------|-------------|
| `get_ui_state()` | Get all windows and taskbar apps with coordinates |
| `get_taskbar_apps()` | List taskbar items with click coordinates |
| `click_taskbar_app(name)` | Click a taskbar app by name |
| `get_all_windows()` | List all open windows |
| `find_and_click_window(title)` | Find and focus a window |

### Screen Capture
| Tool | Description |
|------|-------------|
| `screenshot()` | Capture full screen |
| `zoom(x, y, w, h)` | Capture region at native resolution |
| `set_enhance_mode(bool)` | Toggle contrast enhancement |

### Input
| Tool | Description |
|------|-------------|
| `left_click(x, y)` | Left click at coordinates |
| `right_click(x, y)` | Right click at coordinates |
| `double_click(x, y)` | Double click at coordinates |
| `type_text(text)` | Type ASCII text |
| `key(keys)` | Press key combination (e.g., "ctrl+s") |
| `scroll(x, y, direction)` | Scroll at position |

### AI Vision
| Tool | Description |
|------|-------------|
| `ocr_screen()` | Extract all text from screen (PaddleOCR) |
| `verify_text_on_screen(text)` | Check if text is visible |
| `describe_screen()` | AI description of screen (Florence-2) |

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
```

## License

MIT
