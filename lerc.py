import keyboard
from PIL import Image, ImageGrab
import pytesseract
import ctypes
import csv
import os
import logging
from pathlib import Path
import subprocess
import sys

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='ocr_tool.log'
)

class ItemLookupTool:
    def __init__(self, database_path="items.csv"):
        self.database_path = database_path
        self.setup_tesseract()
        self.verify_database()

    def setup_tesseract(self):
        """Configure and verify Tesseract installation"""
        try:
            # First, try to find Tesseract automatically
            import shutil
            tesseract_path = shutil.which('tesseract')

            if tesseract_path:
                # Found tesseract in PATH
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                logging.info(f"Found Tesseract in PATH: {tesseract_path}")
            else:
                # Common installation paths to check
                possible_paths = [
                    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                ]

                for path in possible_paths:
                    if os.path.exists(path):
                        pytesseract.pytesseract.tesseract_cmd = path
                        logging.info(f"Using Tesseract from: {path}")
                        break
                else:
                    # If not found, check if we need to install it
                    logging.warning("Tesseract path not found. Checking installation...")
                    self.show_dialog("Tesseract OCR not found. Please install Tesseract OCR from:\n"
                                    "https://github.com/UB-Mannheim/tesseract/wiki\n\n"
                                    "After installation, run this script again.")
                    import webbrowser
                    webbrowser.open("https://github.com/UB-Mannheim/tesseract/wiki")
                    sys.exit(1)

            # Test Tesseract with a simple image
            try:
                version = pytesseract.get_tesseract_version()
                logging.info(f"Tesseract version: {version}")
            except Exception as e:
                logging.error(f"Tesseract test failed: {str(e)}")
                self.show_dialog(f"Tesseract OCR test failed: {str(e)}\n\n"
                               "Please make sure Tesseract is properly installed.")
                sys.exit(1)

        except Exception as e:
            logging.error(f"Tesseract setup error: {str(e)}")
            self.show_dialog(f"Error setting up Tesseract: {str(e)}")
            sys.exit(1)

    def verify_database(self):
        """Verify that the database file exists"""
        if not os.path.exists(self.database_path):
            logging.error(f"Database file not found: {self.database_path}")
            self.show_dialog(f"Database file not found: {self.database_path}\nPlease ensure the file exists.")
            return False
        return True

    def get_clipboard_image(self):
        """Grab the image from the clipboard with better handling"""
        try:
            clipboard_content = ImageGrab.grabclipboard()
            logging.info(f"Clipboard content type: {type(clipboard_content)}")

            # Direct image in clipboard
            if isinstance(clipboard_content, Image.Image):
                logging.info(f"Found image in clipboard: {clipboard_content.size} {clipboard_content.mode}")
                return clipboard_content

            # List in clipboard (could contain images or file paths)
            elif isinstance(clipboard_content, list):
                logging.info(f"Found list in clipboard with {len(clipboard_content)} items")

                # Check each item in the list
                for i, item in enumerate(clipboard_content):
                    logging.info(f"  Item {i}: {type(item)}")

                    # If it's an image, use it
                    if isinstance(item, Image.Image):
                        logging.info(f"  -> Using image: {item.size} {item.mode}")
                        return item

                    # If it's a string, it might be a file path
                    elif isinstance(item, str) and (item.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'))):
                        try:
                            image = Image.open(item)
                            logging.info(f"  -> Opened image from path: {item}")
                            return image
                        except Exception as e:
                            logging.error(f"  -> Failed to open image from path: {str(e)}")

            # Try to handle an alternative approach - sometimes Windows puts screenshot in a temp file
            if clipboard_content is None:
                # Check if there's an image on clipboard that PIL can't directly access
                try:
                    from io import BytesIO
                    import win32clipboard

                    win32clipboard.OpenClipboard()
                    try:
                        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
                            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
                            win32clipboard.CloseClipboard()

                            # Convert DIB to PIL Image
                            try:
                                # Skip the BITMAPINFOHEADER (40 bytes) to get to the pixel data
                                # This is a simplified approach and might not work for all DIB formats
                                bytes_io = BytesIO(data[40:])
                                image = Image.open(bytes_io)
                                logging.info("Successfully retrieved image from win32clipboard")
                                return image
                            except Exception as e:
                                logging.error(f"Failed to convert clipboard DIB to image: {str(e)}")
                        else:
                            win32clipboard.CloseClipboard()
                    except Exception as e:
                        win32clipboard.CloseClipboard()
                        logging.error(f"Win32clipboard access error: {str(e)}")
                except ImportError:
                    logging.warning("win32clipboard module not available")

            return None

        except Exception as e:
            logging.error(f"Error accessing clipboard: {str(e)}")
            return None

    def ocr_image(self, image):
        """Perform OCR on the given image with improved configuration"""
        try:
            # Save image to a temporary file because pytesseract might be failing with memory objects
            temp_image_path = "temp_image_for_ocr.png"
            image.save(temp_image_path)

            # Configure tesseract for better accuracy with game text
            custom_config = r'--oem 3 --psm 6'

            # Use the file instead of the image object
            text = pytesseract.image_to_string(temp_image_path, config=custom_config)

            # Delete the temporary file
            try:
                os.remove(temp_image_path)
            except:
                pass

            return text
        except Exception as e:
            logging.error(f"OCR processing error: {str(e)}")
            # Show more detailed error information
            import traceback
            logging.error(f"OCR traceback: {traceback.format_exc()}")
            return ""

    def filter_capitalized_words(self, text):
        """Filter capitalized words from text with improved parsing"""
        if not text:
            return ""

        lines = text.splitlines()
        result = []

        for line in lines:
            words = line.split()
            capitalized_words = []

            for word in words:
                # Clean the word of non-alphanumeric characters
                clean_word = ''.join(c for c in word if c.isalnum() or c in "-' ")
                if clean_word and clean_word.isupper():
                    capitalized_words.append(clean_word)

            if capitalized_words:
                # Remove single-letter words from the beginning
                while capitalized_words and len(capitalized_words[0]) == 1:
                    capitalized_words.pop(0)

                if capitalized_words:
                    result.append(" ".join(capitalized_words))

        return " ".join(result)

    def get_reroll_chance(self, item_name):
        """Retrieve reroll chance for the given item name from database"""
        if not item_name:
            return None

        try:
            with open(self.database_path, mode="r", newline='', encoding='utf-8') as file:
                csv_reader = csv.reader(file)
                next(csv_reader)  # Skip header row

                for row in csv_reader:
                    # Check if row has enough columns and compare with item name
                    if row and len(row) >= 6 and row[0].upper() == item_name.upper():
                        # Return the Rarity column (index 5)
                        return row[5] if row[5] else None
        except Exception as e:
            logging.error(f"Database read error: {str(e)}")
            self.show_dialog(f"Error reading database: {str(e)}")

        return None

    def show_dialog(self, text):
        """Display a dialog box with the given text"""
        try:
            ctypes.windll.user32.MessageBoxW(0, text, "Item Lookup Result", 0x1000)
        except Exception as e:
            logging.error(f"Dialog display error: {str(e)}")

    def process_image(self, image):
        """Process a given image (for direct processing without clipboard)"""
        if image is None:
            self.show_dialog("No valid image provided.")
            return

        # Process the image with OCR
        text = self.ocr_image(image)
        if not text:
            self.show_dialog("No text could be extracted from the image.")
            return

        logging.info(f"OCR extracted text: {text}")

        # Filter and process the text
        filtered_text = self.filter_capitalized_words(text)
        logging.info(f"Filtered text: {filtered_text}")

        if not filtered_text:
            self.show_dialog("No capitalized item name found in the image.")
            return

        # Look up the item in the database
        reroll_chance = self.get_reroll_chance(filtered_text)
        if reroll_chance is not None:
            if reroll_chance == "0% (Common)":
                reroll_text = "Nothing special (Common item)"
            elif reroll_chance == "–" or reroll_chance == "-":
                reroll_text = "No weight assigned"
            else:
                reroll_text = f"Reroll Chance: {reroll_chance}"

            self.show_dialog(f"Item: {filtered_text}\n{reroll_text}")
        else:
            self.show_dialog(f"Item: {filtered_text}\n(Item not found in database)")

        logging.info(f"Processed item: {filtered_text}")

    def process_clipboard(self):
        """Process the clipboard image and display results"""
        logging.info("Processing clipboard image")

        # Get image from clipboard
        image = self.get_clipboard_image()
        if image is None:
            self.show_dialog("No image found on clipboard. Please copy an image first or use Alt+Shift+S to take a screenshot.")
            return

        self.process_image(image)

def main():
    try:
        # Install required modules if missing
        required_modules = ['pywin32', 'Pillow', 'pytesseract', 'keyboard']
        for module in required_modules:
            try:
                __import__(module.split('[')[0])  # Handle modules with extras like 'package[extra]'
            except ImportError:
                logging.info(f"Installing {module} module...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', module])
                logging.info(f"{module} module installed successfully")

        tool = ItemLookupTool()

        # Register hotkeys
        keyboard.add_hotkey("ctrl+shift+d", tool.process_clipboard)

        # Show startup message
        tool.show_dialog("LERC STARTED.\n\n" +
                        "Ctrl+Shift+D: Process image from clipboard\n")

        # Keep the script running
        keyboard.wait()
    except Exception as e:
        logging.critical(f"Critical error: {str(e)}")
        import traceback
        logging.critical(f"Traceback: {traceback.format_exc()}")
        ctypes.windll.user32.MessageBoxW(0, f"Critical error: {str(e)}", "Error", 0x1000)

if __name__ == "__main__":
    main()
