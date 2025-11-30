"""
Advanced Vision Tools for Computer Use

Provides multiple approaches to understand and locate UI elements:
1. Windows UI Automation (pywinauto) - Direct access to accessibility tree
2. Florence-2 Visual Grounding - AI-powered element location by description
3. EasyOCR - Text detection and location
4. Combined smart finder - Uses all approaches intelligently

Requirements:
    pip install pywinauto easyocr torch torchvision transformers pillow

For Florence-2 (optional, for visual grounding):
    pip install transformers accelerate
"""

import io
import json
import time
from dataclasses import dataclass, asdict
from typing import List, Optional, Tuple, Dict, Any
from pathlib import Path

# Lazy imports for optional dependencies
_pywinauto_available = False
_easyocr_available = False
_florence_available = False
_mss_available = False

try:
    import mss
    from PIL import Image
    _mss_available = True
except ImportError:
    pass

try:
    from pywinauto import Desktop, Application
    from pywinauto.findwindows import ElementNotFoundError
    import pywinauto.controls.uia_controls
    _pywinauto_available = True
except ImportError:
    pass

try:
    import easyocr
    _easyocr_available = True
except ImportError:
    pass

try:
    import torch
    from transformers import AutoProcessor, AutoModelForCausalLM
    _florence_available = True
except ImportError:
    pass


@dataclass
class UIElement:
    """Represents a UI element found on screen."""
    name: str
    element_type: str
    x: int
    y: int
    width: int
    height: int
    center_x: int
    center_y: int
    confidence: float = 1.0
    source: str = "unknown"  # "uia", "ocr", "florence", "template"
    extra: Dict[str, Any] = None

    def to_dict(self) -> dict:
        d = asdict(self)
        if d['extra'] is None:
            del d['extra']
        return d

    def __str__(self):
        return f"{self.name} ({self.element_type}) at ({self.center_x}, {self.center_y}) [{self.source}]"


class WindowsUIAutomation:
    """
    Access Windows UI elements through the UI Automation API.
    This is the most reliable way to find elements - no vision needed.
    """

    def __init__(self):
        if not _pywinauto_available:
            raise ImportError("pywinauto not installed. Run: pip install pywinauto")
        self.desktop = Desktop(backend="uia")

    def get_all_windows(self) -> List[UIElement]:
        """Get all visible windows."""
        windows = []
        try:
            for win in self.desktop.windows():
                try:
                    rect = win.rectangle()
                    windows.append(UIElement(
                        name=win.window_text() or "Untitled",
                        element_type="Window",
                        x=rect.left,
                        y=rect.top,
                        width=rect.width(),
                        height=rect.height(),
                        center_x=rect.left + rect.width() // 2,
                        center_y=rect.top + rect.height() // 2,
                        source="uia"
                    ))
                except Exception:
                    continue
        except Exception as e:
            print(f"Error getting windows: {e}")
        return windows

    def find_window(self, title_contains: str = None, class_name: str = None) -> Optional[UIElement]:
        """Find a window by title or class name."""
        try:
            kwargs = {}
            if title_contains:
                kwargs['title_re'] = f".*{title_contains}.*"
            if class_name:
                kwargs['class_name'] = class_name

            win = self.desktop.window(**kwargs)
            rect = win.rectangle()
            return UIElement(
                name=win.window_text() or "Untitled",
                element_type="Window",
                x=rect.left,
                y=rect.top,
                width=rect.width(),
                height=rect.height(),
                center_x=rect.left + rect.width() // 2,
                center_y=rect.top + rect.height() // 2,
                source="uia"
            )
        except ElementNotFoundError:
            return None
        except Exception as e:
            print(f"Error finding window: {e}")
            return None

    def get_taskbar_apps(self) -> List[UIElement]:
        """Get all running apps shown in the taskbar."""
        apps = []
        try:
            # Connect to explorer for taskbar access
            explorer = Application(backend="uia").connect(path="explorer.exe")
            taskbar = explorer.window(class_name="Shell_TrayWnd")

            # Try to find the running apps area
            # Windows 11 structure
            try:
                # Running applications pane
                running_apps = taskbar.child_window(auto_id="TaskListThumbnailWnd", control_type="Pane")
                for btn in running_apps.children(control_type="Button"):
                    try:
                        rect = btn.rectangle()
                        name = btn.window_text() or "Unknown App"
                        apps.append(UIElement(
                            name=name,
                            element_type="TaskbarButton",
                            x=rect.left,
                            y=rect.top,
                            width=rect.width(),
                            height=rect.height(),
                            center_x=rect.left + rect.width() // 2,
                            center_y=rect.top + rect.height() // 2,
                            source="uia"
                        ))
                    except Exception:
                        continue
            except Exception:
                pass

            # Alternative: try to get all buttons in taskbar
            if not apps:
                try:
                    for btn in taskbar.descendants(control_type="Button"):
                        try:
                            rect = btn.rectangle()
                            name = btn.window_text()
                            if name and rect.width() > 20:  # Filter out tiny buttons
                                apps.append(UIElement(
                                    name=name,
                                    element_type="TaskbarButton",
                                    x=rect.left,
                                    y=rect.top,
                                    width=rect.width(),
                                    height=rect.height(),
                                    center_x=rect.left + rect.width() // 2,
                                    center_y=rect.top + rect.height() // 2,
                                    source="uia"
                                ))
                        except Exception:
                            continue
                except Exception:
                    pass

        except Exception as e:
            print(f"Error getting taskbar apps: {e}")

        return apps

    def find_taskbar_app(self, app_name: str) -> Optional[UIElement]:
        """Find a specific app in the taskbar by name."""
        apps = self.get_taskbar_apps()
        app_name_lower = app_name.lower()

        # First try exact match
        for app in apps:
            if app_name_lower == app.name.lower():
                return app

        # Then try contains match
        for app in apps:
            if app_name_lower in app.name.lower():
                return app

        return None

    def get_window_elements(self, window_title: str, element_type: str = None) -> List[UIElement]:
        """Get all elements in a specific window."""
        elements = []
        try:
            win = self.desktop.window(title_re=f".*{window_title}.*")

            control_type = element_type if element_type else None
            descendants = win.descendants(control_type=control_type) if control_type else win.descendants()

            for elem in descendants:
                try:
                    rect = elem.rectangle()
                    if rect.width() > 0 and rect.height() > 0:
                        elements.append(UIElement(
                            name=elem.window_text() or "",
                            element_type=elem.element_info.control_type or "Unknown",
                            x=rect.left,
                            y=rect.top,
                            width=rect.width(),
                            height=rect.height(),
                            center_x=rect.left + rect.width() // 2,
                            center_y=rect.top + rect.height() // 2,
                            source="uia"
                        ))
                except Exception:
                    continue
        except Exception as e:
            print(f"Error getting window elements: {e}")

        return elements

    def dump_taskbar_tree(self) -> str:
        """Debug: Dump the taskbar UI tree structure."""
        try:
            explorer = Application(backend="uia").connect(path="explorer.exe")
            taskbar = explorer.window(class_name="Shell_TrayWnd")
            return taskbar.dump_tree()
        except Exception as e:
            return f"Error: {e}"


class OCREngine:
    """
    Text detection and location using EasyOCR.
    GPU-accelerated when available.
    """

    def __init__(self, languages: List[str] = None, gpu: bool = True):
        if not _easyocr_available:
            raise ImportError("easyocr not installed. Run: pip install easyocr")

        self.languages = languages or ['en']
        self.reader = easyocr.Reader(self.languages, gpu=gpu)

    def find_text_in_image(self, image: "Image.Image", min_confidence: float = 0.3) -> List[UIElement]:
        """Find all text in an image and return locations."""
        import numpy as np

        # Convert PIL image to numpy array
        img_array = np.array(image)

        # Run OCR
        results = self.reader.readtext(img_array)

        elements = []
        for (bbox, text, confidence) in results:
            if confidence >= min_confidence and text.strip():
                # bbox is [[x1,y1], [x2,y1], [x2,y2], [x1,y2]]
                x1 = int(min(p[0] for p in bbox))
                y1 = int(min(p[1] for p in bbox))
                x2 = int(max(p[0] for p in bbox))
                y2 = int(max(p[1] for p in bbox))

                elements.append(UIElement(
                    name=text,
                    element_type="Text",
                    x=x1,
                    y=y1,
                    width=x2 - x1,
                    height=y2 - y1,
                    center_x=(x1 + x2) // 2,
                    center_y=(y1 + y2) // 2,
                    confidence=confidence,
                    source="ocr"
                ))

        return elements

    def find_specific_text(self, image: "Image.Image", search_text: str,
                          case_sensitive: bool = False) -> List[UIElement]:
        """Find specific text in an image."""
        all_text = self.find_text_in_image(image)

        if case_sensitive:
            return [e for e in all_text if search_text in e.name]
        else:
            search_lower = search_text.lower()
            return [e for e in all_text if search_lower in e.name.lower()]


class Florence2Grounding:
    """
    Visual grounding using Microsoft's Florence-2 model.
    Can find objects by natural language description.
    """

    def __init__(self, model_name: str = "microsoft/Florence-2-large", device: str = None):
        if not _florence_available:
            raise ImportError("transformers/torch not installed. Run: pip install transformers torch accelerate")

        self.device = device or ("cuda" if torch.cuda.is_available() else "cpu")
        self.torch_dtype = torch.float16 if self.device == "cuda" else torch.float32

        print(f"Loading Florence-2 on {self.device}...")
        self.model = AutoModelForCausalLM.from_pretrained(
            model_name,
            torch_dtype=self.torch_dtype,
            trust_remote_code=True
        ).to(self.device)

        self.processor = AutoProcessor.from_pretrained(
            model_name,
            trust_remote_code=True
        )
        print("Florence-2 loaded!")

    def _run_task(self, image: "Image.Image", task: str, text_input: str = None) -> dict:
        """Run a Florence-2 task on an image."""
        prompt = task if text_input is None else task + text_input

        inputs = self.processor(
            text=prompt,
            images=image,
            return_tensors="pt"
        ).to(self.device, self.torch_dtype)

        generated_ids = self.model.generate(
            input_ids=inputs["input_ids"],
            pixel_values=inputs["pixel_values"],
            max_new_tokens=1024,
            num_beams=3
        )

        generated_text = self.processor.batch_decode(generated_ids, skip_special_tokens=False)[0]

        return self.processor.post_process_generation(
            generated_text,
            task=task,
            image_size=(image.width, image.height)
        )

    def find_by_description(self, image: "Image.Image", description: str) -> List[UIElement]:
        """
        Find elements matching a natural language description.
        Uses <CAPTION_TO_PHRASE_GROUNDING> task.
        """
        try:
            # Use phrase grounding to find the described element
            result = self._run_task(image, "<CAPTION_TO_PHRASE_GROUNDING>", description)

            elements = []
            if "<CAPTION_TO_PHRASE_GROUNDING>" in result:
                data = result["<CAPTION_TO_PHRASE_GROUNDING>"]
                bboxes = data.get("bboxes", [])
                labels = data.get("labels", [])

                for bbox, label in zip(bboxes, labels):
                    x1, y1, x2, y2 = [int(v) for v in bbox]
                    elements.append(UIElement(
                        name=label,
                        element_type="GroundedObject",
                        x=x1,
                        y=y1,
                        width=x2 - x1,
                        height=y2 - y1,
                        center_x=(x1 + x2) // 2,
                        center_y=(y1 + y2) // 2,
                        source="florence"
                    ))

            return elements
        except Exception as e:
            print(f"Florence-2 grounding error: {e}")
            return []

    def detect_all_objects(self, image: "Image.Image") -> List[UIElement]:
        """Detect all objects in the image."""
        try:
            result = self._run_task(image, "<OD>")

            elements = []
            if "<OD>" in result:
                data = result["<OD>"]
                bboxes = data.get("bboxes", [])
                labels = data.get("labels", [])

                for bbox, label in zip(bboxes, labels):
                    x1, y1, x2, y2 = [int(v) for v in bbox]
                    elements.append(UIElement(
                        name=label,
                        element_type="DetectedObject",
                        x=x1,
                        y=y1,
                        width=x2 - x1,
                        height=y2 - y1,
                        center_x=(x1 + x2) // 2,
                        center_y=(y1 + y2) // 2,
                        source="florence"
                    ))

            return elements
        except Exception as e:
            print(f"Florence-2 detection error: {e}")
            return []

    def ocr(self, image: "Image.Image") -> List[UIElement]:
        """Extract text with locations using Florence-2's OCR capability."""
        try:
            result = self._run_task(image, "<OCR_WITH_REGION>")

            elements = []
            if "<OCR_WITH_REGION>" in result:
                data = result["<OCR_WITH_REGION>"]
                quad_boxes = data.get("quad_boxes", [])
                labels = data.get("labels", [])

                for quad, label in zip(quad_boxes, labels):
                    # quad is [x1,y1, x2,y1, x2,y2, x1,y2]
                    xs = [quad[i] for i in [0, 2, 4, 6]]
                    ys = [quad[i] for i in [1, 3, 5, 7]]
                    x1, x2 = int(min(xs)), int(max(xs))
                    y1, y2 = int(min(ys)), int(max(ys))

                    elements.append(UIElement(
                        name=label,
                        element_type="Text",
                        x=x1,
                        y=y1,
                        width=x2 - x1,
                        height=y2 - y1,
                        center_x=(x1 + x2) // 2,
                        center_y=(y1 + y2) // 2,
                        source="florence-ocr"
                    ))

            return elements
        except Exception as e:
            print(f"Florence-2 OCR error: {e}")
            return []


class SmartElementFinder:
    """
    Unified element finder that combines multiple approaches.
    Tries the most reliable method first, falls back to others.
    """

    def __init__(self, use_uia: bool = True, use_ocr: bool = True,
                 use_florence: bool = False, ocr_gpu: bool = True):
        """
        Initialize the smart finder.

        Args:
            use_uia: Enable Windows UI Automation (most reliable)
            use_ocr: Enable OCR text detection
            use_florence: Enable Florence-2 visual grounding (slower but powerful)
            ocr_gpu: Use GPU for OCR
        """
        self.uia = WindowsUIAutomation() if use_uia and _pywinauto_available else None
        self.ocr = OCREngine(gpu=ocr_gpu) if use_ocr and _easyocr_available else None
        self.florence = None

        # Lazy load Florence-2 only when needed
        self._use_florence = use_florence and _florence_available
        self._florence_loaded = False

    def _ensure_florence(self):
        """Lazy load Florence-2 model."""
        if self._use_florence and not self._florence_loaded:
            try:
                self.florence = Florence2Grounding()
                self._florence_loaded = True
            except Exception as e:
                print(f"Failed to load Florence-2: {e}")
                self._use_florence = False

    def capture_screen(self) -> "Image.Image":
        """Capture the current screen."""
        if not _mss_available:
            raise ImportError("mss not installed. Run: pip install mss")

        with mss.mss() as sct:
            monitor = sct.monitors[1]
            screenshot = sct.grab(monitor)
            return Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

    def find_taskbar_app(self, app_name: str) -> Optional[UIElement]:
        """Find an app in the taskbar. Uses UI Automation (most reliable)."""
        if self.uia:
            return self.uia.find_taskbar_app(app_name)
        return None

    def get_all_taskbar_apps(self) -> List[UIElement]:
        """Get all apps in the taskbar."""
        if self.uia:
            return self.uia.get_taskbar_apps()
        return []

    def find_window(self, title: str) -> Optional[UIElement]:
        """Find a window by title."""
        if self.uia:
            return self.uia.find_window(title_contains=title)
        return None

    def get_all_windows(self) -> List[UIElement]:
        """Get all visible windows."""
        if self.uia:
            return self.uia.get_all_windows()
        return []

    def find_text_on_screen(self, text: str, screenshot: "Image.Image" = None) -> List[UIElement]:
        """Find text on screen using OCR."""
        if not self.ocr:
            return []

        if screenshot is None:
            screenshot = self.capture_screen()

        return self.ocr.find_specific_text(screenshot, text)

    def get_all_text_on_screen(self, screenshot: "Image.Image" = None) -> List[UIElement]:
        """Get all text visible on screen."""
        if not self.ocr:
            return []

        if screenshot is None:
            screenshot = self.capture_screen()

        return self.ocr.find_text_in_image(screenshot)

    def find_by_description(self, description: str, screenshot: "Image.Image" = None) -> List[UIElement]:
        """Find elements by natural language description using Florence-2."""
        self._ensure_florence()

        if not self.florence:
            return []

        if screenshot is None:
            screenshot = self.capture_screen()

        return self.florence.find_by_description(screenshot, description)

    def smart_find(self, query: str, screenshot: "Image.Image" = None) -> List[UIElement]:
        """
        Smart element finder that tries multiple approaches.

        Args:
            query: What to find (app name, text, or description)
            screenshot: Optional screenshot (will capture if not provided)

        Returns:
            List of found elements from all sources, best matches first
        """
        results = []

        # 1. Try UI Automation first (most reliable for apps/windows)
        if self.uia:
            # Check if it's a window
            win = self.uia.find_window(title_contains=query)
            if win:
                results.append(win)

            # Check taskbar
            app = self.uia.find_taskbar_app(query)
            if app:
                results.append(app)

        # 2. Try OCR for text
        if self.ocr:
            if screenshot is None:
                screenshot = self.capture_screen()
            text_matches = self.ocr.find_specific_text(screenshot, query)
            results.extend(text_matches)

        # 3. Try Florence-2 for visual grounding (only if no results yet)
        if not results and self._use_florence:
            self._ensure_florence()
            if self.florence:
                if screenshot is None:
                    screenshot = self.capture_screen()
                visual_matches = self.florence.find_by_description(screenshot, query)
                results.extend(visual_matches)

        # Sort by confidence
        results.sort(key=lambda x: x.confidence, reverse=True)

        return results


def test_uia():
    """Test Windows UI Automation."""
    print("=" * 60)
    print("Testing Windows UI Automation")
    print("=" * 60)

    uia = WindowsUIAutomation()

    print("\n--- All Windows ---")
    windows = uia.get_all_windows()
    for w in windows[:10]:  # First 10
        print(f"  {w}")

    print("\n--- Taskbar Apps ---")
    apps = uia.get_taskbar_apps()
    for app in apps:
        print(f"  {app}")

    print("\n--- Finding Terminal ---")
    terminal = uia.find_taskbar_app("Terminal")
    if terminal:
        print(f"  Found: {terminal}")
        print(f"  Click at: ({terminal.center_x}, {terminal.center_y})")
    else:
        print("  Not found. Trying other names...")
        for name in ["PowerShell", "cmd", "Console", "Windows Terminal"]:
            result = uia.find_taskbar_app(name)
            if result:
                print(f"  Found '{name}': {result}")
                break

    return apps


def test_ocr():
    """Test OCR text detection."""
    print("\n" + "=" * 60)
    print("Testing OCR Text Detection")
    print("=" * 60)

    # Capture screen
    with mss.mss() as sct:
        monitor = sct.monitors[1]
        screenshot = sct.grab(monitor)
        image = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

    ocr = OCREngine(gpu=True)

    print("\n--- All Text Found ---")
    texts = ocr.find_text_in_image(image)
    for t in texts[:20]:  # First 20
        print(f"  '{t.name}' at ({t.center_x}, {t.center_y}) conf={t.confidence:.2f}")

    print(f"\n  Total: {len(texts)} text elements found")

    return texts


def test_florence():
    """Test Florence-2 visual grounding."""
    print("\n" + "=" * 60)
    print("Testing Florence-2 Visual Grounding")
    print("=" * 60)

    # Capture screen
    with mss.mss() as sct:
        monitor = sct.monitors[1]
        screenshot = sct.grab(monitor)
        image = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

    florence = Florence2Grounding()

    print("\n--- Finding 'terminal icon' ---")
    results = florence.find_by_description(image, "terminal icon")
    for r in results:
        print(f"  {r}")

    print("\n--- Finding 'taskbar' ---")
    results = florence.find_by_description(image, "taskbar at bottom of screen")
    for r in results:
        print(f"  {r}")

    print("\n--- All Detected Objects ---")
    objects = florence.detect_all_objects(image)
    for obj in objects[:20]:
        print(f"  {obj}")

    return objects


def test_smart_finder():
    """Test the unified smart finder."""
    print("\n" + "=" * 60)
    print("Testing Smart Element Finder")
    print("=" * 60)

    finder = SmartElementFinder(use_uia=True, use_ocr=True, use_florence=False)

    print("\n--- Finding 'Terminal' ---")
    results = finder.smart_find("Terminal")
    for r in results[:5]:
        print(f"  {r}")

    print("\n--- All Taskbar Apps ---")
    apps = finder.get_all_taskbar_apps()
    for app in apps:
        print(f"  {app}")

    print("\n--- All Windows ---")
    windows = finder.get_all_windows()
    for w in windows[:5]:
        print(f"  {w}")

    return results


if __name__ == "__main__":
    import sys

    print("Vision Tools Test Suite")
    print("=" * 60)
    print(f"pywinauto available: {_pywinauto_available}")
    print(f"easyocr available: {_easyocr_available}")
    print(f"florence available: {_florence_available}")
    print(f"mss available: {_mss_available}")

    if len(sys.argv) > 1:
        test = sys.argv[1]
        if test == "uia":
            test_uia()
        elif test == "ocr":
            test_ocr()
        elif test == "florence":
            test_florence()
        elif test == "smart":
            test_smart_finder()
        else:
            print(f"Unknown test: {test}")
            print("Available: uia, ocr, florence, smart")
    else:
        # Run all tests
        test_uia()
        # test_ocr()  # Uncomment to test OCR
        # test_florence()  # Uncomment to test Florence-2
        test_smart_finder()
