import pyautogui
import pandas as pd
import time
import os
import subprocess
import pyperclip
import pygetwindow as gw
from PIL import Image


class OutlookBotWithImageRecognition:
    def __init__(self, excel_file_path, images_folder="outlook_images"):
        self.excel_file_path = excel_file_path
        self.images_folder = images_folder
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 1.0
        
        if not os.path.exists(self.images_folder):
            os.makedirs(self.images_folder)
            print(f"Created images folder: {self.images_folder}")
            print("Please add the following image files to this folder:")
            print("- new_email_button.png (screenshot of the 'New Email' button)")
            print("- to_field.png (screenshot of the 'To:' field)")
            print("- subject_field.png (screenshot of the 'Subject:' field)")
            print("- body_field.png (screenshot of the email body area)")
            print("- send_button.png (screenshot of the 'Send' button)")
            print("- search_box.png (screenshot of the search box)")
        
    def read_excel_data(self):
        try:
            df = pd.read_excel(self.excel_file_path)
            print(f"Successfully loaded {len(df)} rows from Excel file")
            return df
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return None
    
    def focus_outlook_window_and_maximize(self):
        try:
            windows = [w for w in gw.getAllWindows() if 'Outlook' in w.title or 'Microsoft Outlook' in w.title]
            if windows:
                outlook_window = windows[0]
                print(f"Found Outlook window: {outlook_window.title}")
                outlook_window.activate()
                outlook_window.maximize()
                time.sleep(3)
                return True
            else:
                print("Outlook window not found!")
                return False
        except Exception as e:
            print(f"Error focusing Outlook window: {e}")
            return False
    
    def open_outlook(self):
        try:
            print("Opening Outlook...")
            subprocess.Popen(["start", "outlook"], shell=True)
            time.sleep(10)
            return True
        except Exception as e:
            print(f"Error opening Outlook: {e}")
            return False
    
    def open_outlook_via_run(self):
        try:
            print("Opening Outlook via Run dialog...")
            pyautogui.hotkey('win', 'r')
            time.sleep(1)
            pyautogui.write('outlook')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(10)
            return True
        except Exception as e:
            print(f"Error opening Outlook via Run: {e}")
            return False
    
    def close_outlook(self):
        try:
            print("Closing Outlook...")
            windows = [w for w in gw.getAllWindows() if 'Outlook' in w.title or 'Microsoft Outlook' in w.title]
            if windows:
                outlook_window = windows[0]
                outlook_window.activate()
                time.sleep(1)
                pyautogui.hotkey('alt', 'f4')
                time.sleep(2)
                print("Outlook closed successfully")
                return True
            else:
                print("Outlook window not found for closing")
                return False
        except Exception as e:
            print(f"Error closing Outlook: {e}")
            try:
                subprocess.run(["taskkill", "/f", "/im", "OUTLOOK.EXE"], shell=True)
                print("Outlook force closed via taskkill")
                return True
            except:
                return False
    
    def find_and_click_image(self, image_name, confidence=0.8, timeout=10):
        image_path = os.path.join(self.images_folder, image_name)
        
        if not os.path.exists(image_path):
            print(f"Image file not found: {image_path}")
            return False
        
        print(f"Looking for image: {image_name}")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                location = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if location:
                    center = pyautogui.center(location)
                    print(f"Found {image_name} at {center}")
                    pyautogui.click(center)
                    time.sleep(1)
                    return True
            except pyautogui.ImageNotFoundException:
                pass
            except Exception as e:
                print(f"Error looking for {image_name}: {e}")
            
            time.sleep(0.5)
        
        print(f"Could not find {image_name} within {timeout} seconds")
        return False
    
    def search_for_contact(self, email_address):
        try:
            print(f"Searching for contact: {email_address}")
            
            if self.find_and_click_image("search_box.png"):
                time.sleep(1)
                pyautogui.write(f'from:{email_address}')
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(3)
                return True
            else:
                print("Using keyboard shortcut for search...")
                pyautogui.hotkey('ctrl', 'e')
                time.sleep(2)
                pyautogui.write(f'from:{email_address}')
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(3)
                return True
                
        except Exception as e:
            print(f"Error searching for contact: {e}")
            return False
    
    def click_new_email_button(self):
        try:
            print("Looking for New Email button...")
            
            if self.find_and_click_image("new_email_button.png"):
                print("Successfully clicked New Email button")
                time.sleep(3)
                return True
            else:
                print("Using keyboard shortcut Ctrl+N...")
                pyautogui.hotkey('ctrl', 'n')
                time.sleep(3)
                return True
                
        except Exception as e:
            print(f"Error clicking New Email button: {e}")
            return False
    
    def fill_to_field(self, email_address):
        try:
            print(f"Filling To field with: {email_address}")
            pyperclip.copy(email_address)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)

            pyautogui.moveTo(100, 300)
            pyautogui.rightClick()
            time.sleep(0.5)

            return True

        except Exception as e:
            print(f"Error filling To field: {e}")
            return False

    def fill_subject_field(self, subject):
        try:
            print(f"Filling Subject field with: {subject}")

            pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(0.5)

            pyperclip.copy(subject)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            return True

        except Exception as e:
            print(f"Error using TAB to fill Subject: {e}")
            try:
                if self.find_and_click_image("subject_field.png"):
                    pyperclip.copy(subject)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(1)
                    return True
            except Exception as e2:
                print(f"Error with image fallback: {e2}")
            return False

    def fill_body_field(self, body):
        try:
            print("Filling email body...")

            pyautogui.press('tab')
            time.sleep(0.5)

            pyperclip.copy(body)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            return True

        except Exception as e:
            print(f"Error using TAB to fill body: {e}")
            try:
                if self.find_and_click_image("body_field.png"):
                    pyperclip.copy(body)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(1)
                    return True
            except Exception as e2:
                print(f"Error with image fallback: {e2}")
            return False

    def click_send_button(self):
        try:
            print("Looking for Send button...")
            
            if self.find_and_click_image("send_button.png"):
                print("Successfully clicked Send button")
                time.sleep(3)
                return True
            else:
                print("Using keyboard shortcut Ctrl+Enter...")
                pyautogui.hotkey('ctrl', 'enter')
                time.sleep(3)
                return True
                
        except Exception as e:
            print(f"Error clicking Send button: {e}")
            try:
                pyautogui.hotkey('alt', 's')
                time.sleep(3)
                return True
            except:
                return False
    
    def close_search_results(self):
        try:
            pyautogui.press('escape')
            time.sleep(1)
            return True
        except:
            return True
    
    def create_sample_images_info(self):
        info = """
        Required Image Files (save in '{folder}' folder):
        
        1. new_email_button.png - Screenshot of the "New Email" or "New" button in Outlook
        2. to_field.png - Screenshot of the "To:" field in compose window
        3. subject_field.png - Screenshot of the "Subject:" field
        4. body_field.png - Screenshot of the email body/content area
        5. send_button.png - Screenshot of the "Send" button
        6. search_box.png - Screenshot of the search box in Outlook
        
        Tips for taking screenshots:
        - Use high resolution images
        - Capture only the specific UI element
        - Ensure good contrast and clarity
        - Test with different Outlook themes if needed
        """.format(folder=self.images_folder)
        
        print(info)
    
    def process_all_emails(self):
        print("=" * 50)
        print("OUTLOOK EMAIL AUTOMATION STARTED - LIMIT: 2 EMAILS")
        print("=" * 50)
        
        df = self.read_excel_data()
        if df is None or df.empty:
            print("Failed to read Excel file. Exiting...")
            return False
        
        if not os.path.exists(self.images_folder) or len(os.listdir(self.images_folder)) == 0:
            self.create_sample_images_info()
            return False
        
        if not self.open_outlook():
            if not self.open_outlook_via_run():
                print("Failed to open Outlook. Exiting...")
                return False
        
        if not self.focus_outlook_window_and_maximize():
            print("Failed to focus Outlook window. Exiting...")
            return False
        
        max_emails = min(2, len(df))
        print(f"Processing only {max_emails} emails (limited to 2)...")
        successful_emails = 0
        failed_emails = 0
        
        for index in range(max_emails):
            try:
                row = df.iloc[index]
                email = str(row.iloc[0]).strip()
                email = email.replace(";", "").strip()
                subject = str(row.iloc[1]).strip()
                body = str(row.iloc[2]).strip()
                
                print(f"\n--- Processing Email {index + 1}/{max_emails} ---")
                print(f"To: {email}")
                print(f"Subject: {subject[:50]}...")
                
                if not self.click_new_email_button():
                    print(f"Failed to open new email window for {email}")
                    failed_emails += 1
                    continue
                
                if not self.fill_to_field(email):
                    print(f"Failed to fill To field for {email}")
                    failed_emails += 1
                    continue
                
                if not self.fill_subject_field(subject):
                    print(f"Failed to fill Subject field for {email}")
                    failed_emails += 1
                    continue
                
                if not self.fill_body_field(body):
                    print(f"Failed to fill body for {email}")
                    failed_emails += 1
                    continue
                
                if not self.click_send_button():
                    print(f"Failed to send email to {email}")
                    failed_emails += 1
                    continue
                
                print(f"✓ Email sent successfully to: {email}")
                successful_emails += 1
                
                self.close_search_results()
                time.sleep(2)
                
            except Exception as e:
                print(f"Error processing email {index + 1}: {e}")
                failed_emails += 1
                continue
        
        print("\n" + "=" * 50)
        print("AUTOMATION SUMMARY")
        print("=" * 50)
        print(f"Total emails processed: {max_emails}")
        print(f"Successful: {successful_emails}")
        print(f"Failed: {failed_emails}")
        print(f"Success rate: {(successful_emails/max_emails*100):.1f}%")
        
        print("\nClosing Outlook...")
        self.close_outlook()
        
        return successful_emails > 0


def main():
    excel_file_path = "C:\\Users\\SWE - Mostafa Gamal\\Desktop\\GUIBOT\\outlook_Bot\\outlook Message.xlsx"
    images_folder = "outlook_images"
    
    if not os.path.exists(excel_file_path):
        print(f"Excel file not found: {excel_file_path}")
        print("Please make sure the Excel file exists and the path is correct.")
        return
    
    bot = OutlookBotWithImageRecognition(excel_file_path, images_folder)
    
    print("Enhanced Outlook Email Automation Bot - LIMITED TO 2 EMAILS")
    print("=========================================================")
    
    success = bot.process_all_emails()
    
    if success:
        print("\n Outlook automation completed successfully!")
        print("Sent 2 emails and closed Outlook as requested.")
    else:
        print("\n Outlook automation encountered issues.")
        print("Please check the error messages above and try again.")


if __name__ == "__main__":
    main()