"""
Example usage of the Computer Use Agent.

This demonstrates how to use Claude to control your Windows desktop.
"""

from computer_use_agent import ComputerUseAgent, DisplayConfig


def simple_task():
    """Run a simple task like opening notepad."""
    agent = ComputerUseAgent()

    # Simple task
    result = agent.run("Open the Windows Start menu")
    print(f"Result: {result}")


def task_with_callback():
    """Run a task with a callback to monitor actions."""

    def on_action(action: str, params: dict):
        print(f"  -> Executing: {action}")
        if params:
            print(f"     Params: {params}")

    agent = ComputerUseAgent(callback=on_action)
    result = agent.run("Open Notepad and type 'Hello from Claude!'")
    print(f"Result: {result}")


def custom_model():
    """Use a specific model (e.g., Opus for more complex tasks)."""
    agent = ComputerUseAgent(
        model="claude-opus-4-5-20251101",
        max_iterations=30
    )
    result = agent.run("Find and open the Calculator app, then compute 123 * 456")
    print(f"Result: {result}")


def check_display():
    """Check the display configuration."""
    config = DisplayConfig.from_primary_monitor()
    print(f"Primary monitor: {config.width}x{config.height}")


if __name__ == "__main__":
    import sys

    print("Computer Use Agent Examples")
    print("=" * 40)

    # Show display info
    check_display()
    print()

    # Run based on argument
    if len(sys.argv) > 1:
        example = sys.argv[1]
        if example == "simple":
            simple_task()
        elif example == "callback":
            task_with_callback()
        elif example == "opus":
            custom_model()
        else:
            print(f"Unknown example: {example}")
    else:
        print("Available examples:")
        print("  python example.py simple   - Open start menu")
        print("  python example.py callback - Open notepad with action logging")
        print("  python example.py opus     - Use Opus model for calculator")
