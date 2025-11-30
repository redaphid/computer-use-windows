"""
ComfyUI Vision Integration for Computer Use

Uses ComfyUI's BLIP nodes to analyze screenshots and answer questions about them.
This enables asking questions like "What website is this?" or "What text is visible?"

Requirements:
    - ComfyUI running on localhost:8188
    - BLIP nodes installed in ComfyUI
"""

import base64
import io
import json
import time
import uuid
from typing import Optional
from urllib import request, error
from PIL import Image

COMFYUI_URL = "http://localhost:8188"


def check_comfyui_available() -> bool:
    """Check if ComfyUI is running and accessible."""
    try:
        req = request.urlopen(f"{COMFYUI_URL}/system_stats", timeout=5)
        return req.status == 200
    except:
        return False


def upload_image(image: Image.Image, name: str = None) -> str:
    """Upload an image to ComfyUI and return the filename."""
    if name is None:
        name = f"screenshot_{uuid.uuid4().hex[:8]}.png"

    # Convert image to bytes
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    image_bytes = buffer.getvalue()

    # Create multipart form data
    boundary = uuid.uuid4().hex

    body = (
        f'--{boundary}\r\n'
        f'Content-Disposition: form-data; name="image"; filename="{name}"\r\n'
        f'Content-Type: image/png\r\n\r\n'
    ).encode('utf-8') + image_bytes + f'\r\n--{boundary}--\r\n'.encode('utf-8')

    headers = {
        'Content-Type': f'multipart/form-data; boundary={boundary}',
        'Content-Length': str(len(body))
    }

    req = request.Request(
        f"{COMFYUI_URL}/upload/image",
        data=body,
        headers=headers,
        method='POST'
    )

    try:
        response = request.urlopen(req, timeout=30)
        result = json.loads(response.read().decode('utf-8'))
        return result.get('name', name)
    except Exception as e:
        print(f"Upload error: {e}")
        return name


def queue_prompt(workflow: dict) -> str:
    """Queue a workflow and return the prompt_id."""
    data = json.dumps({"prompt": workflow}).encode('utf-8')
    req = request.Request(
        f"{COMFYUI_URL}/prompt",
        data=data,
        headers={'Content-Type': 'application/json'},
        method='POST'
    )
    response = request.urlopen(req, timeout=30)
    result = json.loads(response.read().decode('utf-8'))
    return result['prompt_id']


def get_history(prompt_id: str) -> Optional[dict]:
    """Get the execution history for a prompt."""
    try:
        req = request.urlopen(f"{COMFYUI_URL}/history/{prompt_id}", timeout=10)
        history = json.loads(req.read().decode('utf-8'))
        return history.get(prompt_id)
    except:
        return None


def wait_for_completion(prompt_id: str, timeout: float = 60) -> Optional[dict]:
    """Wait for a prompt to complete and return results."""
    start = time.time()
    while time.time() - start < timeout:
        history = get_history(prompt_id)
        if history and history.get('status', {}).get('completed', False):
            return history
        time.sleep(0.5)
    return None


def create_blip_workflow(image_filename: str, mode: str = "caption", question: str = "") -> dict:
    """Create a BLIP analysis workflow.

    Args:
        image_filename: Name of uploaded image file
        mode: "caption" for general description, "interrogate" for detailed
        question: Question context (used with interrogate mode for VQA-style prompting)

    Returns:
        Workflow dict ready to queue
    """
    # Use interrogate mode for questions, caption for general description
    actual_mode = "interrogate" if question else mode

    workflow = {
        # Load the image
        "1": {
            "class_type": "LoadImage",
            "inputs": {
                "image": image_filename
            }
        },
        # Load BLIP model (with correct Salesforce/ prefix)
        "2": {
            "class_type": "BLIP Model Loader",
            "inputs": {
                "blip_model": "Salesforce/blip-image-captioning-large",
                "vqa_model_id": "Salesforce/blip-vqa-base",
                "device": "cuda"
            }
        },
        # Analyze with BLIP
        "3": {
            "class_type": "BLIP Analyze Image",
            "inputs": {
                "images": ["1", 0],
                "mode": actual_mode,
                "question": question if question else "What does this image show?",
                "blip_model": ["2", 0]
            }
        },
        # Output node (required for ComfyUI to execute)
        "4": {
            "class_type": "DisplayAny",
            "inputs": {
                "input": ["3", 0],
                "mode": "raw value"
            }
        }
    }
    return workflow


def analyze_image(image: Image.Image, mode: str = "caption", question: str = "", timeout: float = 60) -> Optional[str]:
    """Analyze an image using BLIP via ComfyUI.

    Args:
        image: PIL Image to analyze
        mode: "caption", "interrogate", or "question"
        question: Question for VQA mode
        timeout: Max seconds to wait

    Returns:
        Analysis result string, or None if failed
    """
    if not check_comfyui_available():
        return "Error: ComfyUI not available at " + COMFYUI_URL

    try:
        # Upload image
        filename = upload_image(image)

        # Create and queue workflow
        workflow = create_blip_workflow(filename, mode, question)
        prompt_id = queue_prompt(workflow)

        # Wait for result
        result = wait_for_completion(prompt_id, timeout)

        if result:
            # Extract output from DisplayAny node (node 4)
            outputs = result.get('outputs', {})
            display_output = outputs.get('4', {})

            # DisplayAny outputs text as character array - join them
            if 'text' in display_output:
                text_list = display_output['text']
                if isinstance(text_list, list):
                    # Join character array into string
                    return ''.join(text_list).strip()
                return str(text_list)

            # Try other output formats
            for key, value in outputs.items():
                if isinstance(value, dict) and 'text' in value:
                    text_data = value['text']
                    if isinstance(text_data, list):
                        return ''.join(text_data).strip()
                    return str(text_data)

            return f"Completed but no text output found. Outputs: {outputs}"

        return "Timeout waiting for ComfyUI response"

    except Exception as e:
        return f"Error: {e}"


def caption_screenshot(image: Image.Image) -> str:
    """Get a general caption/description of a screenshot."""
    return analyze_image(image, mode="caption")


def interrogate_screenshot(image: Image.Image) -> str:
    """Get a detailed description of a screenshot."""
    return analyze_image(image, mode="interrogate")


def ask_about_screenshot(image: Image.Image, question: str) -> str:
    """Ask a specific question about a screenshot."""
    return analyze_image(image, mode="question", question=question)


# ============================================================================
# Test functions
# ============================================================================

def capture_screen() -> Image.Image:
    """Capture the current screen."""
    import mss
    with mss.mss() as sct:
        monitor = sct.monitors[1]
        screenshot = sct.grab(monitor)
        return Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")


def test_comfyui_connection():
    """Test basic ComfyUI connectivity."""
    print("Testing ComfyUI connection...")

    if check_comfyui_available():
        print("  ComfyUI is available!")

        # Get system info
        req = request.urlopen(f"{COMFYUI_URL}/system_stats")
        info = json.loads(req.read().decode('utf-8'))
        print(f"  Version: {info['system']['comfyui_version']}")
        print(f"  GPU: {info['devices'][0]['name']}")
        print(f"  VRAM Free: {info['devices'][0]['vram_free'] / 1e9:.1f} GB")
        return True
    else:
        print("  ComfyUI not available!")
        return False


def test_image_upload():
    """Test uploading a screenshot to ComfyUI."""
    print("\nTesting image upload...")

    # Create a simple test image
    img = Image.new('RGB', (100, 100), color='red')
    filename = upload_image(img, "test_image.png")
    print(f"  Uploaded as: {filename}")
    return filename


def test_blip_caption():
    """Test BLIP captioning on current screen."""
    print("\nTesting BLIP captioning...")

    # Capture screen
    print("  Capturing screen...")
    img = capture_screen()

    # Resize for faster processing (BLIP doesn't need full 4K)
    img.thumbnail((1024, 1024), Image.Resampling.LANCZOS)
    print(f"  Resized to: {img.size}")

    # Get caption
    print("  Sending to ComfyUI for analysis...")
    result = caption_screenshot(img)
    print(f"  Caption: {result}")
    return result


def test_blip_question():
    """Test BLIP VQA on current screen."""
    print("\nTesting BLIP Visual Question Answering...")

    # Capture screen
    img = capture_screen()
    img.thumbnail((1024, 1024), Image.Resampling.LANCZOS)

    # Ask a question
    question = "What website or application is shown in this image?"
    print(f"  Question: {question}")

    result = ask_about_screenshot(img, question)
    print(f"  Answer: {result}")
    return result


if __name__ == "__main__":
    print("=" * 60)
    print("ComfyUI Vision Integration Test")
    print("=" * 60)

    if test_comfyui_connection():
        test_image_upload()
        test_blip_caption()
        test_blip_question()

    print("\n" + "=" * 60)
    print("Test complete!")
