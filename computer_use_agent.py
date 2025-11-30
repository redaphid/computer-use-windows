"""
Computer Use Agent for Windows

This module implements Claude's computer use capability, allowing Claude to
control the desktop through screenshots, mouse, and keyboard actions.

Requires: anthropic, pyautogui, mss, Pillow
"""

import base64
import io
import time
from dataclasses import dataclass
from typing import Callable

import anthropic
import mss
import pyautogui
from PIL import Image

# Safety settings for pyautogui
pyautogui.FAILSAFE = True  # Move mouse to corner to abort
pyautogui.PAUSE = 0.1  # Small pause between actions


@dataclass
class DisplayConfig:
    """Configuration for the display being controlled."""
    width: int
    height: int
    display_number: int = 1

    @classmethod
    def from_primary_monitor(cls) -> "DisplayConfig":
        """Create config from the primary monitor's actual resolution."""
        with mss.mss() as sct:
            monitor = sct.monitors[1]  # Primary monitor
            return cls(
                width=monitor["width"],
                height=monitor["height"],
                display_number=1
            )


class ScreenCapture:
    """Handles screenshot capture on Windows."""

    def __init__(self, config: DisplayConfig, max_dimension: int = 1280):
        self.config = config
        self.max_dimension = max_dimension
        self._scale_factor = self._calculate_scale_factor()

    def _calculate_scale_factor(self) -> float:
        """Calculate scale factor to keep screenshots under max dimension."""
        max_actual = max(self.config.width, self.config.height)
        if max_actual <= self.max_dimension:
            return 1.0
        return self.max_dimension / max_actual

    def capture(self) -> str:
        """Capture screenshot and return as base64-encoded PNG."""
        with mss.mss() as sct:
            monitor = sct.monitors[self.config.display_number]
            screenshot = sct.grab(monitor)

            # Convert to PIL Image
            img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

            # Resize if needed for API efficiency
            if self._scale_factor < 1.0:
                new_size = (
                    int(img.width * self._scale_factor),
                    int(img.height * self._scale_factor)
                )
                img = img.resize(new_size, Image.Resampling.LANCZOS)

            # Convert to base64 PNG
            buffer = io.BytesIO()
            img.save(buffer, format="PNG", optimize=True)
            return base64.standard_b64encode(buffer.getvalue()).decode("utf-8")

    def scale_coordinates(self, x: int, y: int) -> tuple[int, int]:
        """Scale coordinates from API space to actual screen space."""
        if self._scale_factor < 1.0:
            return (
                int(x / self._scale_factor),
                int(y / self._scale_factor)
            )
        return (x, y)


class ActionExecutor:
    """Executes mouse and keyboard actions on Windows."""

    def __init__(self, screen_capture: ScreenCapture):
        self.screen = screen_capture

    def execute(self, action: str, **params) -> str:
        """Execute a computer use action and return result."""
        try:
            handler = getattr(self, f"_action_{action}", None)
            if handler is None:
                return f"Unknown action: {action}"
            return handler(**params)
        except Exception as e:
            return f"Error executing {action}: {e}"

    def _action_screenshot(self, **_) -> str:
        """Take a screenshot (handled separately in agent loop)."""
        return "Screenshot captured"

    def _action_left_click(self, coordinate: list[int], **_) -> str:
        x, y = self.screen.scale_coordinates(coordinate[0], coordinate[1])
        pyautogui.click(x, y)
        return f"Left clicked at ({x}, {y})"

    def _action_right_click(self, coordinate: list[int], **_) -> str:
        x, y = self.screen.scale_coordinates(coordinate[0], coordinate[1])
        pyautogui.rightClick(x, y)
        return f"Right clicked at ({x}, {y})"

    def _action_middle_click(self, coordinate: list[int], **_) -> str:
        x, y = self.screen.scale_coordinates(coordinate[0], coordinate[1])
        pyautogui.middleClick(x, y)
        return f"Middle clicked at ({x}, {y})"

    def _action_double_click(self, coordinate: list[int], **_) -> str:
        x, y = self.screen.scale_coordinates(coordinate[0], coordinate[1])
        pyautogui.doubleClick(x, y)
        return f"Double clicked at ({x}, {y})"

    def _action_triple_click(self, coordinate: list[int], **_) -> str:
        x, y = self.screen.scale_coordinates(coordinate[0], coordinate[1])
        pyautogui.tripleClick(x, y)
        return f"Triple clicked at ({x}, {y})"

    def _action_mouse_move(self, coordinate: list[int], **_) -> str:
        x, y = self.screen.scale_coordinates(coordinate[0], coordinate[1])
        pyautogui.moveTo(x, y)
        return f"Moved mouse to ({x}, {y})"

    def _action_left_click_drag(self, start_coordinate: list[int], end_coordinate: list[int], **_) -> str:
        sx, sy = self.screen.scale_coordinates(start_coordinate[0], start_coordinate[1])
        ex, ey = self.screen.scale_coordinates(end_coordinate[0], end_coordinate[1])
        pyautogui.moveTo(sx, sy)
        pyautogui.drag(ex - sx, ey - sy)
        return f"Dragged from ({sx}, {sy}) to ({ex}, {ey})"

    def _action_type(self, text: str, **_) -> str:
        pyautogui.typewrite(text, interval=0.02)
        return f"Typed: {text[:50]}{'...' if len(text) > 50 else ''}"

    def _action_key(self, key: str, **_) -> str:
        # Handle key combinations like "ctrl+s"
        keys = key.lower().split("+")
        if len(keys) > 1:
            pyautogui.hotkey(*keys)
        else:
            pyautogui.press(key)
        return f"Pressed key: {key}"

    def _action_scroll(self, coordinate: list[int], direction: str, amount: int = 3, **_) -> str:
        x, y = self.screen.scale_coordinates(coordinate[0], coordinate[1])
        pyautogui.moveTo(x, y)

        scroll_amount = amount if direction == "up" else -amount
        if direction in ("up", "down"):
            pyautogui.scroll(scroll_amount)
        else:
            pyautogui.hscroll(scroll_amount if direction == "right" else -scroll_amount)

        return f"Scrolled {direction} by {amount} at ({x}, {y})"

    def _action_wait(self, duration: int = 1, **_) -> str:
        time.sleep(duration)
        return f"Waited {duration} seconds"

    def _action_hold_key(self, key: str, **_) -> str:
        # This is typically used in combination with other actions
        pyautogui.keyDown(key)
        return f"Holding key: {key}"


class ComputerUseAgent:
    """
    Agent that allows Claude to control the computer.

    Example usage:
        agent = ComputerUseAgent()
        result = agent.run("Open notepad and type 'Hello World'")
        print(result)
    """

    def __init__(
        self,
        model: str = "claude-sonnet-4-20250514",
        display_config: DisplayConfig | None = None,
        max_iterations: int = 50,
        callback: Callable[[str, dict], None] | None = None
    ):
        self.client = anthropic.Anthropic()
        self.model = model
        self.config = display_config or DisplayConfig.from_primary_monitor()
        self.max_iterations = max_iterations
        self.callback = callback or (lambda action, params: None)

        self.screen = ScreenCapture(self.config)
        self.executor = ActionExecutor(self.screen)

        # Determine tool version based on model
        if "opus-4-5" in model:
            self.tool_type = "computer_20251124"
            self.beta_flag = "computer-use-2025-11-24"
        else:
            self.tool_type = "computer_20250124"
            self.beta_flag = "computer-use-2025-01-24"

    def _get_computer_tool(self) -> dict:
        """Get the computer tool definition for the API."""
        return {
            "type": self.tool_type,
            "name": "computer",
            "display_width_px": self.config.width,
            "display_height_px": self.config.height,
            "display_number": self.config.display_number
        }

    def _create_screenshot_content(self) -> dict:
        """Create a screenshot content block for the API."""
        screenshot_b64 = self.screen.capture()
        return {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": screenshot_b64
            }
        }

    def run(self, task: str, system_prompt: str | None = None) -> str:
        """
        Run a task using computer use.

        Args:
            task: The task to accomplish (e.g., "Open notepad and type hello")
            system_prompt: Optional system prompt override

        Returns:
            Final text response from Claude
        """
        default_system = """You are a computer use agent that can control a Windows desktop.
You can see the screen through screenshots and interact using mouse and keyboard.

Guidelines:
- Take a screenshot first to see the current state
- Use keyboard shortcuts when possible (e.g., Win+R for Run dialog)
- Verify actions succeeded by taking screenshots
- Be precise with click coordinates
- Wait briefly after actions that trigger UI changes"""

        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": task},
                    self._create_screenshot_content()
                ]
            }
        ]

        for iteration in range(self.max_iterations):
            print(f"\n--- Iteration {iteration + 1} ---")

            response = self.client.beta.messages.create(
                model=self.model,
                max_tokens=4096,
                system=system_prompt or default_system,
                tools=[self._get_computer_tool()],
                messages=messages,
                betas=[self.beta_flag]
            )

            # Check if we're done
            if response.stop_reason == "end_turn":
                # Extract final text response
                for block in response.content:
                    if hasattr(block, "text"):
                        return block.text
                return "Task completed"

            # Process tool use
            tool_results = []
            for block in response.content:
                if block.type == "tool_use":
                    action = block.input.get("action", "screenshot")
                    params = {k: v for k, v in block.input.items() if k != "action"}

                    print(f"Action: {action} | Params: {params}")
                    self.callback(action, params)

                    # Execute the action
                    result = self.executor.execute(action, **params)
                    print(f"Result: {result}")

                    # Build tool result
                    tool_result = {
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": []
                    }

                    # Always include a screenshot after actions
                    tool_result["content"].append({
                        "type": "text",
                        "text": result
                    })
                    tool_result["content"].append(self._create_screenshot_content())

                    tool_results.append(tool_result)

            # Add assistant response and tool results to messages
            messages.append({"role": "assistant", "content": response.content})
            messages.append({"role": "user", "content": tool_results})

        return "Max iterations reached"


def main():
    """Example usage of the computer use agent."""
    import sys

    if len(sys.argv) < 2:
        print("Usage: python computer_use_agent.py <task>")
        print('Example: python computer_use_agent.py "Open notepad"')
        sys.exit(1)

    task = " ".join(sys.argv[1:])

    print(f"Display config: {DisplayConfig.from_primary_monitor()}")
    print(f"Task: {task}")
    print("-" * 50)

    agent = ComputerUseAgent()
    result = agent.run(task)

    print("\n" + "=" * 50)
    print("Final result:")
    print(result)


if __name__ == "__main__":
    main()
