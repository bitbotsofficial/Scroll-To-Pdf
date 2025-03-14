import os
import time
from PIL import ImageGrab, Image
import pyautogui
import img2pdf
try:
    import pygetwindow as gw
except ImportError:
    gw = None  # Fallback if pygetwindow isn’t installed
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QLabel, QDoubleSpinBox, QSpinBox, QPushButton, QFileDialog, 
                            QMessageBox, QFrame, QGridLayout, QProgressBar, QCheckBox)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QIcon, QPalette, QColor

class CaptureThread(QThread):
    screenshot_taken = pyqtSignal(int)
    capture_complete = pyqtSignal(bool, int)
    status_update = pyqtSignal(str)

    def __init__(self, delay, max_scrolls, manual_height, is_fullscreen):
        super().__init__()
        self.delay = delay
        self.max_scrolls = max_scrolls
        self.manual_height = manual_height
        self.is_fullscreen = is_fullscreen
        self.capturing = True
        self.screenshots = []

    def run(self):
        time.sleep(3)  # Additional wait after fullscreen adjustment (total 5 seconds from start)
        scroll_count = 0
        previous_screenshot = None

        if self.manual_height > 0:
            scroll_height = self.manual_height
            self.status_update.emit(f"Using manual height: {scroll_height}px")
        else:
            scroll_height = 1300 if self.is_fullscreen else 1245
            mode = "fullscreen" if self.is_fullscreen else "non-fullscreen"
            self.status_update.emit(f"Using default {mode} height: {scroll_height}px")

        while self.capturing and (self.max_scrolls == 0 or scroll_count < self.max_scrolls):
            self.status_update.emit(f"Capturing screenshot {scroll_count + 1}...")
            current_screenshot = ImageGrab.grab()

            if previous_screenshot is not None:
                similarity, remaining_height = self.check_page_end(current_screenshot, previous_screenshot, scroll_height)
                if similarity > 0.98:
                    self.status_update.emit(f"High similarity detected: {similarity:.3f}")
                    if remaining_height < 35:
                        self.status_update.emit(f"Page end detected - remaining content: {remaining_height}px")
                        break
                    else:
                        self.status_update.emit(f"Continuing - remaining content: {remaining_height}px")
                else:
                    self.status_update.emit(f"Content differs - similarity: {similarity:.3f}")

            self.screenshots.append(current_screenshot)
            self.screenshot_taken.emit(len(self.screenshots))
            
            previous_screenshot = current_screenshot
            pyautogui.scroll(-scroll_height)
            scroll_count += 1
            time.sleep(self.delay)

        self.capture_complete.emit(scroll_count < self.max_scrolls or self.max_scrolls == 0, len(self.screenshots))

    def stop(self):
        self.capturing = False

    def check_page_end(self, img1, img2, scroll_height):
        img1 = img1.resize((100, 100)).convert('L')
        img2 = img2.resize((100, 100)).convert('L')
        
        pixels1 = list(img1.getdata())
        pixels2 = list(img2.getdata())
        diff_pixels = sum(1 for p1, p2 in zip(pixels1, pixels2) if abs(p1 - p2) > 40)
        similarity = 1 - (diff_pixels / len(pixels1))

        original_height = img2.size[1]
        scale_factor = original_height / 100
        
        bottom_diff = 0
        for y in range(99, -1, -1):
            row1 = [img1.getpixel((x, y)) for x in range(100)]
            row2 = [img2.getpixel((x, y)) for x in range(100)]
            if any(abs(p1 - p2) > 40 for p1, p2 in zip(row1, row2)):
                bottom_diff = 100 - y
                break
        
        remaining_height = (bottom_diff * scale_factor) if bottom_diff > 0 else 0
        
        return similarity, remaining_height

class AutoScrollCapturePDF(QMainWindow):
    def __init__(self):
        super().__init__()
        self.screenshots = []
        self.capturing = False
        self.capture_thread = None
        self.setWindowTitle("Scroll2Pdf")
        self.resize(600, 800)
        self.setup_ui()
        icon_path = os.path.join(os.path.dirname(__file__), "app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.update_stylesheet_based_on_theme()

    def setup_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(25, 25, 25, 25)
        main_layout.setSpacing(20)
        self.setCentralWidget(main_widget)

        title_frame = QFrame()
        title_frame.setObjectName("neumorphic")
        title_layout = QVBoxLayout(title_frame)
        title_layout.setContentsMargins(15, 20, 15, 20)
        title_label = QLabel("Auto-Scrolling Webpage Capture")
        title_label.setObjectName("title")
        title_layout.addWidget(title_label)
        main_layout.addWidget(title_frame)

        settings_frame = QFrame()
        settings_frame.setObjectName("neumorphic")
        settings_layout = QGridLayout(settings_frame)
        settings_layout.setContentsMargins(20, 20, 20, 20)
        settings_layout.setSpacing(15)
        main_layout.addWidget(settings_frame)

        delay_label = QLabel("Delay between scrolls (seconds):")
        settings_layout.addWidget(delay_label, 0, 0)
        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0.1, 5.0)
        self.delay_spin.setValue(0.5)
        self.delay_spin.setSingleStep(0.1)
        self.delay_spin.setFixedWidth(120)
        settings_layout.addWidget(self.delay_spin, 0, 1)

        max_label = QLabel("Max scrolls (0 = unlimited):")
        settings_layout.addWidget(max_label, 1, 0)
        self.max_spin = QDoubleSpinBox()
        self.max_spin.setRange(0, 500)
        self.max_spin.setValue(10)
        self.max_spin.setFixedWidth(120)
        settings_layout.addWidget(self.max_spin, 1, 1)

        height_label = QLabel("Scroll height (0 = auto, pixels):")
        settings_layout.addWidget(height_label, 2, 0)
        self.height_spin = QSpinBox()
        self.height_spin.setRange(0, 4000)
        self.height_spin.setValue(0)
        self.height_spin.setFixedWidth(120)
        settings_layout.addWidget(self.height_spin, 2, 1)

        fullscreen_label = QLabel("Fullscreen mode (auto-toggle):")
        settings_layout.addWidget(fullscreen_label, 3, 0)
        self.fullscreen_check = QCheckBox()
        self.fullscreen_check.setChecked(False)
        settings_layout.addWidget(self.fullscreen_check, 3, 1)

        status_frame = QFrame()
        status_frame.setObjectName("neumorphic")
        status_layout = QVBoxLayout(status_frame)
        status_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.addWidget(status_frame)

        self.status_label = QLabel("Ready to capture")
        self.status_label.setObjectName("status")
        status_layout.addWidget(self.status_label)

        self.counter_label = QLabel("Screenshots: 0")
        self.counter_label.setObjectName("counter")
        status_layout.addWidget(self.counter_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        status_layout.addWidget(self.progress_bar)

        buttons_frame = QFrame()
        buttons_frame.setObjectName("neumorphic")
        buttons_layout = QVBoxLayout(buttons_frame)
        buttons_layout.setContentsMargins(20, 20, 20, 20)
        buttons_layout.setSpacing(15)
        main_layout.addWidget(buttons_frame)

        self.start_button = QPushButton("Start Capture")
        self.start_button.setObjectName("start")
        self.start_button.clicked.connect(self.start_capture)
        buttons_layout.addWidget(self.start_button)

        self.stop_button = QPushButton("Stop Capture")
        self.stop_button.setObjectName("stop")
        self.stop_button.clicked.connect(self.stop_capture)
        self.stop_button.setEnabled(False)
        buttons_layout.addWidget(self.stop_button)

        self.save_button = QPushButton("Save as PDF")
        self.save_button.setObjectName("save")
        self.save_button.clicked.connect(self.save_pdf)
        self.save_button.setEnabled(False)
        buttons_layout.addWidget(self.save_button)

        self.preview_button = QPushButton("Preview Screenshots")
        self.preview_button.setObjectName("preview")
        self.preview_button.clicked.connect(self.preview_screenshots)
        self.preview_button.setEnabled(False)
        buttons_layout.addWidget(self.preview_button)

        self.clear_button = QPushButton("Clear Screenshots")
        self.clear_button.setObjectName("clear")
        self.clear_button.clicked.connect(self.clear_screenshots)
        self.clear_button.setEnabled(False)
        buttons_layout.addWidget(self.clear_button)

    def update_stylesheet_based_on_theme(self):
        palette = QApplication.palette()
        is_dark_mode = palette.color(QPalette.ColorRole.Window).lightness() < 128
        
        if is_dark_mode:
            self.setStyleSheet("""
                QMainWindow { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2C3E50, stop:1 #4CA1AF); }
                QWidget { color: #E0E0E0; font-family: 'Segoe UI', Arial; }
                QLabel { color: #E0E0E0; }
                QLabel#title { color: white; font-size: 22px; font-weight: bold; qproperty-alignment: AlignCenter; }
                QLabel#counter { font-size: 14px; qproperty-alignment: AlignCenter; }
                QLabel#status { color: #64B5F6; font-size: 13px; font-style: italic; qproperty-alignment: AlignCenter; }
                QPushButton { border: none; border-radius: 8px; padding: 12px; font-weight: bold; color: white; font-size: 13px; }
                QPushButton:disabled { background-color: #7f8c8d; }
                QPushButton#start { background-color: #2ecc71; }
                QPushButton#start:hover:!disabled { background-color: #27ae60; }
                QPushButton#stop { background-color: #e74c3c; }
                QPushButton#stop:hover:!disabled { background-color: #c0392b; }
                QPushButton#save { background-color: #3498db; }
                QPushButton#save:hover:!disabled { background-color: #2980b9; }
                QPushButton#clear { background-color: #f39c12; color: #333; }
                QPushButton#clear:hover:!disabled { background-color: #d35400; color: white; }
                QPushButton#preview { background-color: #9b59b6; }
                QPushButton#preview:hover:!disabled { background-color: #8e44ad; }
                QDoubleSpinBox, QSpinBox, QCheckBox { border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 6px; padding: 6px; background-color: rgba(255, 255, 255, 0.05); color: #E0E0E0; }
                QFrame#neumorphic { background-color: rgba(44, 62, 80, 0.7); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 16px; }
                QProgressBar { border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 5px; background-color: rgba(255, 255, 255, 0.05); }
                QProgressBar::chunk { background-color: #3498db; border-radius: 3px; }
                QMessageBox { background-color: #2C3E50; color: #E0E0E0; }
                QMessageBox QPushButton { background-color: #3498db; color: white; border: none; padding: 5px; }
                QMessageBox QPushButton:hover { background-color: #2980b9; }
            """)
        else:
            self.setStyleSheet("""
                QMainWindow { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #E0E0E0, stop:1 #FFFFFF); }
                QWidget { color: #333333; font-family: 'Segoe UI', Arial; }
                QLabel { color: #333333; }
                QLabel#title { color: #2C3E50; font-size: 22px; font-weight: bold; qproperty-alignment: AlignCenter; }
                QLabel#counter { font-size: 14px; qproperty-alignment: AlignCenter; }
                QLabel#status { color: #1976D2; font-size: 13px; font-style: italic; qproperty-alignment: AlignCenter; }
                QPushButton { border: none; border-radius: 8px; padding: 12px; font-weight: bold; color: white; font-size: 13px; }
                QPushButton:disabled { background-color: #B0BEC5; }
                QPushButton#start { background-color: #2ecc71; }
                QPushButton#start:hover:!disabled { background-color: #27ae60; }
                QPushButton#stop { background-color: #e74c3c; }
                QPushButton#stop:hover:!disabled { background-color: #c0392b; }
                QPushButton#save { background-color: #3498db; }
                QPushButton#save:hover:!disabled { background-color: #2980b9; }
                QPushButton#clear { background-color: #f39c12; color: #333; }
                QPushButton#clear:hover:!disabled { background-color: #d35400; color: white; }
                QPushButton#preview { background-color: #9b59b6; }
                QPushButton#preview:hover:!disabled { background-color: #8e44ad; }
                QDoubleSpinBox, QSpinBox, QCheckBox { border: 1px solid rgba(0, 0, 0, 0.2); border-radius: 6px; padding: 6px; background-color: rgba(255, 255, 255, 0.8); color: #333333; }
                QFrame#neumorphic { background-color: rgba(255, 255, 255, 0.7); border: 1px solid rgba(0, 0, 0, 0.1); border-radius: 16px; }
                QProgressBar { border: 1px solid rgba(0, 0, 0, 0.2); border-radius: 5px; background-color: rgba(255, 255, 255, 0.8); }
                QProgressBar::chunk { background-color: #3498db; border-radius: 3px; }
                QMessageBox { background-color: #FFFFFF; color: #333333; }
                QMessageBox QPushButton { background-color: #3498db; color: white; border: none; padding: 5px; }
                QMessageBox QPushButton:hover { background-color: #2980b9; }
            """)

    def changeEvent(self, event):
        if event.type() == event.Type.StyleChange:
            self.update_stylesheet_based_on_theme()
        super().changeEvent(event)

    def is_fullscreen_active(self):
        """Check if the current screen matches the full screen resolution."""
        screen_width, screen_height = pyautogui.size()
        screenshot = ImageGrab.grab()
        screenshot_width, screenshot_height = screenshot.size
        
        # Stricter threshold: 98% of screen size
        is_fullscreen = (screenshot_width >= screen_width * 0.98 and 
                         screenshot_height >= screen_height * 0.98)
        
        # Additional check with pygetwindow if available
        if gw:
            browser_window = self.find_browser_window()
            if browser_window:
                try:
                    win_width, win_height = browser_window.width, browser_window.height
                    win_x, win_y = browser_window.left, browser_window.top
                    print(f"Window size: {win_width}x{win_height}, Position: ({win_x}, {win_y})")
                    # True fullscreen should have minimal offset and match screen size closely
                    is_fullscreen = (is_fullscreen and 
                                   win_width >= screen_width * 0.98 and 
                                   win_height >= screen_height * 0.98 and 
                                   win_x <= 5 and win_y <= 5)
                except Exception as e:
                    print(f"Error checking window size with pygetwindow: {str(e)}")
        
        print(f"Screen: {screen_width}x{screen_height}, Screenshot: {screenshot_width}x{screenshot_height}, "
              f"Fullscreen detected: {is_fullscreen}")
        return is_fullscreen

    def find_browser_window(self):
        """Find and return an active browser window using pygetwindow."""
        if not gw:
            print("pygetwindow not installed, falling back to pyautogui")
            return None
        
        browser_titles = ["Chrome", "Firefox", "Edge", "Opera", "Safari"]
        windows = gw.getAllWindows()
        for window in windows:
            if not window.title.strip():
                continue
            if any(browser in window.title for browser in browser_titles) and window.visible:
                try:
                    print(f"Found browser window: {window.title} (Active: {window.isActive})")
                    if not window.isActive:
                        window.activate()
                        time.sleep(1)  # Wait for activation
                    if window.isActive:
                        return window
                except Exception as e:
                    print(f"Error activating window {window.title}: {str(e)}")
        print("No suitable browser window found")
        return None

    def adjust_fullscreen(self, desired_fullscreen):
        """Adjust fullscreen state with retries and improved detection."""
        self.status_label.setText("Adjusting window mode...")
        time.sleep(2)  # Initial delay to allow GUI minimization and browser focus
        
        browser_window = self.find_browser_window()
        max_attempts = 3
        attempt = 0
        
        while attempt < max_attempts:
            attempt += 1
            print(f"Fullscreen adjustment attempt {attempt}/{max_attempts}")
            
            if browser_window:
                try:
                    browser_window.activate()
                    time.sleep(1)  # Ensure focus
                    current_fullscreen = self.is_fullscreen_active()
                    print(f"Desired fullscreen: {desired_fullscreen}, Current fullscreen: {current_fullscreen}")
                    
                    if desired_fullscreen != current_fullscreen:
                        print(f"Toggling fullscreen with F11 (desired: {desired_fullscreen})")
                        pyautogui.hotkey('f11')
                        time.sleep(2)  # Wait for transition
                        if self.is_fullscreen_active() == desired_fullscreen:
                            self.status_label.setText(f"Switched to {'fullscreen' if desired_fullscreen else 'normal'} mode")
                            return
                        else:
                            print(f"Fullscreen toggle failed on attempt {attempt}")
                    else:
                        # Force one toggle attempt if detection might be wrong
                        if attempt == 1 and desired_fullscreen:
                            print("Forcing fullscreen toggle due to potential detection error")
                            pyautogui.hotkey('f11')
                            time.sleep(2)
                            if self.is_fullscreen_active():
                                self.status_label.setText("Switched to fullscreen mode (forced)")
                                return
                            else:
                                print("Forced toggle failed, continuing attempts")
                        else:
                            self.status_label.setText(f"Already in {'fullscreen' if current_fullscreen else 'normal'} mode")
                            return
                except Exception as e:
                    print(f"Error with browser window adjustment: {str(e)}")
            else:
                # Fallback with pyautogui
                print("Falling back to pyautogui")
                pyautogui.moveTo(50, 50)
                pyautogui.click()
                time.sleep(1)
                current_fullscreen = self.is_fullscreen_active()
                print(f"Desired fullscreen: {desired_fullscreen}, Current fullscreen: {current_fullscreen}")
                
                if desired_fullscreen != current_fullscreen:
                    print(f"Toggling fullscreen with F11 (fallback, attempt {attempt})")
                    pyautogui.hotkey('f11')
                    time.sleep(2)
                    if self.is_fullscreen_active() == desired_fullscreen:
                        self.status_label.setText(f"Switched to {'fullscreen' if desired_fullscreen else 'normal'} mode (fallback)")
                        return
                    else:
                        print(f"Fullscreen toggle failed (fallback) on attempt {attempt}")
                else:
                    # Force one toggle attempt in fallback mode too
                    if attempt == 1 and desired_fullscreen:
                        print("Forcing fullscreen toggle due to potential detection error (fallback)")
                        pyautogui.hotkey('f11')
                        time.sleep(2)
                        if self.is_fullscreen_active():
                            self.status_label.setText("Switched to fullscreen mode (forced, fallback)")
                            return
                        else:
                            print("Forced toggle failed (fallback), continuing attempts")
                    else:
                        self.status_label.setText(f"Already in {'fullscreen' if current_fullscreen else 'normal'} mode (fallback)")
                        return
        
        self.status_label.setText(f"Failed to {'enter' if desired_fullscreen else 'exit'} fullscreen - press F11 manually")
        print("Warning: All fullscreen adjustment attempts failed")

    def start_capture(self):
        if self.capturing:
            return
        self.capturing = True
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.save_button.setEnabled(False)
        self.preview_button.setEnabled(False)
        self.clear_button.setEnabled(False)

        delay = self.delay_spin.value()
        max_scrolls = int(self.max_spin.value())
        manual_height = self.height_spin.value()
        is_fullscreen = self.fullscreen_check.isChecked()

        QMessageBox.information(self, "Auto-Scrolling Capture", 
            "Focus your browser window now. Fullscreen adjustment starts in 2 seconds if checked.\nCapture begins 5 seconds after adjustment.")
        self.showMinimized()
        self.status_label.setText("Preparing to capture...")

        self.adjust_fullscreen(is_fullscreen)
        time.sleep(3)  # Additional wait after adjustment (total 5 seconds from message)

        self.capture_thread = CaptureThread(delay, max_scrolls, manual_height, self.is_fullscreen_active())
        self.capture_thread.screenshot_taken.connect(self.update_counter)
        self.capture_thread.capture_complete.connect(self.capture_finished)
        self.capture_thread.status_update.connect(self.update_status)
        self.capture_thread.start()

    def update_counter(self, count):
        self.counter_label.setText(f"Screenshots: {count}")
        self.screenshots = self.capture_thread.screenshots
        if self.capture_thread.max_scrolls > 0:
            progress = min(100, int((count / self.capture_thread.max_scrolls) * 100))
            self.progress_bar.setValue(progress)

    def update_status(self, status):
        self.status_label.setText(status)

    def capture_finished(self, page_end_detected, count):
        self.capturing = False
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(True)
        self.preview_button.setEnabled(True)
        self.clear_button.setEnabled(True)
        self.showNormal()
        self.status_label.setText("Capture complete" if page_end_detected else "Max scrolls reached")
        self.progress_bar.setValue(100 if page_end_detected else self.progress_bar.value())

    def stop_capture(self):
        if self.capture_thread:
            self.capture_thread.stop()
        self.capturing = False
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(True)
        self.preview_button.setEnabled(True)
        self.clear_button.setEnabled(True)
        self.status_label.setText("Capture stopped")
        self.showNormal()

    def clear_screenshots(self):
        self.screenshots = []
        self.counter_label.setText("Screenshots: 0")
        self.status_label.setText("Screenshots cleared")
        self.save_button.setEnabled(False)
        self.preview_button.setEnabled(False)
        self.clear_button.setEnabled(False)
        self.progress_bar.setValue(0)

    def preview_screenshots(self):
        if not self.screenshots:
            QMessageBox.warning(self, "No Screenshots", "No screenshots to preview.")
            return
        total_height = sum(img.height for img in self.screenshots)
        combined = Image.new('RGB', (self.screenshots[0].width, total_height))
        y_offset = 0
        for img in self.screenshots:
            combined.paste(img, (0, y_offset))
            y_offset += img.height
        combined.show()

    def save_pdf(self):
        if not self.screenshots:
            QMessageBox.warning(self, "No Screenshots", "No screenshots to save.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Save PDF As", "", "PDF files (*.pdf)")
        if not file_path:
            return

        try:
            self.status_label.setText("Saving PDF...")
            rgb_screenshots = [img.convert('RGB') for img in self.screenshots]
            rgb_screenshots[0].save(file_path, save_all=True, append_images=rgb_screenshots[1:], resolution=100.0)
            QMessageBox.information(self, "Success", f"PDF saved at: {file_path}")
            self.status_label.setText("PDF saved successfully")
        except Exception as e:
            try:
                with open(file_path, "wb") as pdf_file:
                    pdf_file.write(img2pdf.convert([img.convert('RGB') for img in self.screenshots]))
                QMessageBox.information(self, "Success", f"PDF saved using fallback method at: {file_path}")
                self.status_label.setText("PDF saved successfully (fallback)")
            except Exception as e2:
                QMessageBox.critical(self, "Error", f"Failed to save PDF:\nPrimary error: {str(e)}\nFallback error: {str(e2)}")
                self.status_label.setText("Error saving PDF")
                print(f"Primary error: {str(e)}\nFallback error: {str(e2)}")

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = AutoScrollCapturePDF()
    window.show()
    sys.exit(app.exec())