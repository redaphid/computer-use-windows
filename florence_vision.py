"""
Vision Integration for Computer Use

- PaddleOCR for accurate text extraction
- Florence-2 for image descriptions and captions

Requirements:
    - PyTorch with CUDA
    - transformers, timm, einops (for Florence-2)
    - paddleocr (for OCR)
"""

import torch
from PIL import Image
from typing import Optional
import mss
import numpy as np

# Lazy load models to save memory
_florence_model = None
_florence_processor = None
_paddle_ocr = None


def load_paddle_ocr():
    """Load PaddleOCR (lazy loading)."""
    global _paddle_ocr

    if _paddle_ocr is None:
        print("Loading PaddleOCR...")
        from paddleocr import PaddleOCR
        _paddle_ocr = PaddleOCR(lang='en')
        print("PaddleOCR loaded!")

    return _paddle_ocr


def load_florence():
    """Load Florence-2 model (lazy loading)."""
    global _florence_model, _florence_processor

    if _florence_model is None:
        print("Loading Florence-2 model...")
        from transformers import AutoProcessor, AutoModelForCausalLM

        model_id = "microsoft/Florence-2-large"

        _florence_processor = AutoProcessor.from_pretrained(model_id, trust_remote_code=True)
        _florence_model = AutoModelForCausalLM.from_pretrained(
            model_id,
            trust_remote_code=True,
            torch_dtype=torch.float16
        ).to("cuda")

        print("Florence-2 loaded!")

    return _florence_model, _florence_processor


def run_florence(image: Image.Image, task: str, text_input: str = "") -> str:
    """Run Florence-2 on an image."""
    model, processor = load_florence()

    prompt = task if not text_input else f"{task} {text_input}"
    inputs = processor(text=prompt, images=image, return_tensors="pt").to("cuda", torch.float16)

    with torch.no_grad():
        generated_ids = model.generate(
            input_ids=inputs["input_ids"],
            pixel_values=inputs["pixel_values"],
            max_new_tokens=1024,
            num_beams=3,
            do_sample=False
        )

    generated_text = processor.batch_decode(generated_ids, skip_special_tokens=False)[0]
    result = processor.post_process_generation(
        generated_text,
        task=task,
        image_size=(image.width, image.height)
    )

    return result


def ocr_screenshot(image: Image.Image) -> str:
    """Extract all text from screenshot using PaddleOCR.

    Returns text organized by position (top to bottom, left to right).
    """
    ocr = load_paddle_ocr()

    # Convert PIL Image to numpy array
    img_array = np.array(image)

    # Run OCR using new predict API
    result = ocr.predict(img_array)

    if not result:
        return ""

    # Extract text with positions for sorting
    text_items = []
    for item in result:
        if 'rec_texts' in item and 'rec_polys' in item:
            texts = item['rec_texts']
            polys = item['rec_polys']
            for text, poly in zip(texts, polys):
                # poly is [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
                y_pos = poly[0][1]
                x_pos = poly[0][0]
                text_items.append((y_pos, x_pos, text))

    # Sort by y position (top to bottom), then x position (left to right)
    text_items.sort(key=lambda x: (x[0] // 30, x[1]))  # Group by ~30px rows

    # Join text
    return ' '.join(item[2] for item in text_items)


def ocr_with_regions(image: Image.Image) -> list:
    """Extract text with bounding box regions using PaddleOCR.

    Returns list of dicts with 'text', 'bbox', 'confidence'.
    """
    ocr = load_paddle_ocr()
    img_array = np.array(image)
    result = ocr.predict(img_array)

    if not result:
        return []

    regions = []
    for item in result:
        if 'rec_texts' in item and 'rec_polys' in item and 'rec_scores' in item:
            texts = item['rec_texts']
            polys = item['rec_polys']
            scores = item['rec_scores']
            for text, poly, score in zip(texts, polys, scores):
                regions.append({
                    'text': text,
                    'confidence': score,
                    'bbox': poly
                })

    return regions


def caption_screenshot(image: Image.Image) -> str:
    """Get a caption/description of the screenshot using Florence-2."""
    result = run_florence(image, "<CAPTION>")
    return result.get("<CAPTION>", "")


def detailed_caption(image: Image.Image) -> str:
    """Get a detailed caption of the screenshot using Florence-2."""
    result = run_florence(image, "<DETAILED_CAPTION>")
    return result.get("<DETAILED_CAPTION>", "")


def detect_objects(image: Image.Image) -> dict:
    """Detect objects and their locations using Florence-2."""
    result = run_florence(image, "<OD>")
    return result.get("<OD>", {})


def caption_to_grounding(image: Image.Image, caption: str) -> dict:
    """Find regions matching a text description using Florence-2."""
    result = run_florence(image, "<CAPTION_TO_PHRASE_GROUNDING>", caption)
    return result.get("<CAPTION_TO_PHRASE_GROUNDING>", {})


def capture_screen() -> Image.Image:
    """Capture the current screen."""
    with mss.mss() as sct:
        monitor = sct.monitors[1]
        screenshot = sct.grab(monitor)
        return Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")


# ============================================================================
# Test functions
# ============================================================================

def test_ocr():
    """Test PaddleOCR on current screen."""
    print("=" * 60)
    print("PaddleOCR Test")
    print("=" * 60)

    print("\nCapturing screen...")
    img = capture_screen()
    print(f"Screen size: {img.size}")

    print("\n--- Running OCR ---")
    text = ocr_screenshot(img)
    print(f"Text found ({len(text)} chars):")
    print(text[:1000] if len(text) > 1000 else text)

    print("\n" + "=" * 60)


def test_florence():
    """Test Florence-2 on current screen."""
    print("=" * 60)
    print("Florence-2 Vision Test")
    print("=" * 60)

    print("\nCapturing screen...")
    img = capture_screen()
    img.thumbnail((1024, 1024), Image.Resampling.LANCZOS)
    print(f"Resized to: {img.size}")

    print("\n--- Caption ---")
    caption = caption_screenshot(img)
    print(f"Caption: {caption}")

    print("\n--- Detailed Caption ---")
    detailed = detailed_caption(img)
    print(f"Detailed: {detailed}")

    print("\n" + "=" * 60)


if __name__ == "__main__":
    test_ocr()
