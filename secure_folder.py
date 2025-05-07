import cv2
import pyttsx3
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox, Listbox, Entry, Label, Button, Frame, Toplevel
import json
import os
import winreg
import sys
import pythoncom
import win32com.client
import win32serviceutil
import win32service
import win32event
import servicemanager
from deepface import DeepFace
import time
import ctypes
import logging
import subprocess
import win32con
import win32api
import hashlib
import tempfile

# Data directory and file paths
DATA_DIR = "C:\\FolderProtectorData"
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
PROTECTED_FOLDERS_FILE = os.path.join(DATA_DIR, "protected_folders.json")
REFERENCE_IMAGE = os.path.join(DATA_DIR, "reference_face.jpg")
LOG_FILE = os.path.join(DATA_DIR, "folder_protector.log")

# Initialize data directory for first run
def initialize_data_dir(first_run=False):
    try:
        if not os.path.exists(DATA_DIR):
            os.makedirs(DATA_DIR)
            logging.info("Created data directory: %s", DATA_DIR)
        if first_run:
            # Delete old files only on first run
            for file in [CONFIG_FILE, PROTECTED_FOLDERS_FILE, REFERENCE_IMAGE, LOG_FILE]:
                if os.path.exists(file):
                    os.remove(file)
                    logging.info("Deleted old file: %s", file)
        return True
    except Exception as e:
        logging.error("Error initializing data directory: %s", str(e))
        messagebox.showerror("Error", f"Cannot create {DATA_DIR}: {e}")
        return False

# Configure logging with fallback
def configure_logging():
    try:
        # Check if directory exists to determine first run
        first_run = not os.path.exists(DATA_DIR)
        if initialize_data_dir(first_run):
            logging.basicConfig(filename=LOG_FILE, level=logging.DEBUG,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.info("Logging configured successfully. First run: %s", first_run)
        else:
            # Fallback to temp directory
            fallback_log = os.path.join(tempfile.gettempdir(), "folder_protector_fallback.log")
            logging.basicConfig(filename=fallback_log, level=logging.DEBUG,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.error("Failed to create data directory, using fallback log: %s", fallback_log)
            messagebox.showerror("Error", f"Cannot create {DATA_DIR}. Using fallback log at {fallback_log}.")
    except Exception as e:
        fallback_log = os.path.join(tempfile.gettempdir(), "folder_protector_fallback.log")
        logging.basicConfig(filename=fallback_log, level=logging.DEBUG,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        logging.error("Error configuring logging: %s", str(e))
        messagebox.showerror("Error", f"Logging setup failed: {e}")

# Configure logging
configure_logging()

# Check for admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    logging.error("Program requires administrator privileges.")
    messagebox.showerror("Error", "Please run the program with administrator privileges (Run as Administrator).")
    input("Press Enter to exit...")
    exit()

# Hash password
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# Initial setup
def setup_application():
    def save_setup():
        password = password_entry.get()
        confirm_password = confirm_entry.get()
        if not password or not confirm_password:
            messagebox.showerror("Error", "Please enter password and confirm password!")
            return
        if password != confirm_password:
            messagebox.showerror("Error", "Passwords do not match!")
            return
        
        # Capture face image
        try:
            video_capture = cv2.VideoCapture(0)
            if not video_capture.isOpened():
                logging.error("Cannot open camera.")
                messagebox.showerror("Error", "Cannot access camera.")
                return
            
            messagebox.showinfo("Capture Image", "Look at the camera, image will be captured in 5 seconds.")
            start_time = time.time()
            while True:
                ret, frame = video_capture.read()
                if not ret:
                    logging.error("Cannot capture image from camera.")
                    messagebox.showerror("Error", "Cannot capture image from camera.")
                    video_capture.release()
                    return
                cv2.imshow('Capture Face Image', frame)
                if time.time() - start_time > 5:
                    cv2.imwrite(REFERENCE_IMAGE, frame)
                    video_capture.release()
                    cv2.destroyAllWindows()
                    break
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    video_capture.release()
                    cv2.destroyAllWindows()
                    return
        except Exception as e:
            logging.error("Error capturing face image: %s", str(e))
            messagebox.showerror("Error", f"Face capture failed: {e}")
            return
        
        # Save configuration
        try:
            config = {
                "password_hash": hash_password(password),
                "reference_image": REFERENCE_IMAGE
            }
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f)
            logging.info("Setup completed, configuration saved.")
            messagebox.showinfo("Success", "Setup completed!")
            setup_window.destroy()
        except Exception as e:
            logging.error("Error saving configuration: %s", str(e))
            messagebox.showerror("Error", f"Failed to save configuration: {e}")

    setup_window = tk.Tk()
    setup_window.title("Application Setup")
    setup_window.geometry("300x200")
    
    Label(setup_window, text="Enter Password:").pack(pady=5)
    password_entry = Entry(setup_window, show='*')
    password_entry.pack(pady=5)
    
    Label(setup_window, text="Confirm Password:").pack(pady=5)
    confirm_entry = Entry(setup_window, show='*')
    confirm_entry.pack(pady=5)
    
    Button(setup_window, text="Save and Capture Image", command=save_setup).pack(pady=10)
    setup_window.mainloop()

# Function to change password
def change_password():
    def save_new_password():
        current_password = current_password_entry.get()
        new_password = new_password_entry.get()
        confirm_password = confirm_password_entry.get()
        
        # Verify current password
        config = load_config()
        if not config:
            messagebox.showerror("Error", "Configuration not found!")
            return
        
        # Check current password
        if hash_password(current_password) != config.get("password_hash"):
            messagebox.showerror("Error", "Current password is incorrect!")
            return
            
        # Validate new password
        if not new_password or not confirm_password:
            messagebox.showerror("Error", "Please enter new password and confirmation!")
            return
        if new_password != confirm_password:
            messagebox.showerror("Error", "New passwords do not match!")
            return
            
        # Update password
        try:
            config["password_hash"] = hash_password(new_password)
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f)
            logging.info("Password updated successfully.")
            messagebox.showinfo("Success", "Password updated successfully!")
            change_pwd_window.destroy()
        except Exception as e:
            logging.error("Error updating password: %s", str(e))
            messagebox.showerror("Error", f"Failed to update password: {e}")
    
    # Create password change window
    change_pwd_window = Toplevel()
    change_pwd_window.title("Change Password")
    change_pwd_window.geometry("300x250")
    
    Label(change_pwd_window, text="Current Password:").pack(pady=5)
    current_password_entry = Entry(change_pwd_window, show='*')
    current_password_entry.pack(pady=5)
    
    Label(change_pwd_window, text="New Password:").pack(pady=5)
    new_password_entry = Entry(change_pwd_window, show='*')
    new_password_entry.pack(pady=5)
    
    Label(change_pwd_window, text="Confirm New Password:").pack(pady=5)
    confirm_password_entry = Entry(change_pwd_window, show='*')
    confirm_password_entry.pack(pady=5)
    
    Button(change_pwd_window, text="Update Password", command=save_new_password).pack(pady=10)
    Button(change_pwd_window, text="Cancel", command=change_pwd_window.destroy).pack(pady=5)

# Function to update face recognition image
def update_face_image():
    def capture_new_image():
        # Verify password first
        password = password_entry.get()
        if not password:
            messagebox.showerror("Error", "Please enter your password!")
            return
            
        # Check password
        config = load_config()
        if not config:
            messagebox.showerror("Error", "Configuration not found!")
            return
            
        if hash_password(password) != config.get("password_hash"):
            messagebox.showerror("Error", "Incorrect password!")
            return
            
        # Capture new face image
        try:
            video_capture = cv2.VideoCapture(0)
            if not video_capture.isOpened():
                logging.error("Cannot open camera.")
                messagebox.showerror("Error", "Cannot access camera.")
                return
            
            messagebox.showinfo("Update Face Image", "Look at the camera, image will be captured in 3 seconds.")
            start_time = time.time()
            
            while True:
                ret, frame = video_capture.read()
                if not ret:
                    logging.error("Cannot capture image from camera.")
                    messagebox.showerror("Error", "Cannot capture image from camera.")
                    video_capture.release()
                    return
                    
                cv2.imshow('Capture New Face Image', frame)
                
                if time.time() - start_time > 3:
                    # Backup old image
                    backup_image = os.path.join(DATA_DIR, "reference_face_backup.jpg")
                    if os.path.exists(REFERENCE_IMAGE):
                        os.rename(REFERENCE_IMAGE, backup_image)
                        
                    # Save new image
                    cv2.imwrite(REFERENCE_IMAGE, frame)
                    video_capture.release()
                    cv2.destroyAllWindows()
                    
                    messagebox.showinfo("Success", "Face image updated successfully!")
                    update_face_window.destroy()
                    break
                    
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    video_capture.release()
                    cv2.destroyAllWindows()
                    return
                    
        except Exception as e:
            logging.error("Error updating face image: %s", str(e))
            messagebox.showerror("Error", f"Face image update failed: {e}")
            return
    
    # Create face update window
    update_face_window = Toplevel()
    update_face_window.title("Update Face Image")
    update_face_window.geometry("300x180")
    
    Label(update_face_window, text="Please verify your password:").pack(pady=10)
    password_entry = Entry(update_face_window, show='*')
    password_entry.pack(pady=5)
    
    Button(update_face_window, text="Capture New Image", command=capture_new_image).pack(pady=15)
    Button(update_face_window, text="Cancel", command=update_face_window.destroy).pack(pady=5)

# Check if setup is complete
def is_setup_complete():
    result = os.path.exists(CONFIG_FILE) and os.path.exists(REFERENCE_IMAGE)
    logging.info("Setup complete check: %s", result)
    return result

# Load configuration
def load_config():
    try:
        if is_setup_complete():
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                logging.info("Configuration loaded successfully.")
                return config
        logging.warning("Configuration file not found.")
        return {}
    except Exception as e:
        logging.error("Error loading configuration: %s", str(e))
        return {}

# Authenticate password
def authenticate_password():
    config = load_config()
    if not config:
        logging.error("Authentication failed: Application not set up.")
        messagebox.showerror("Error", "Application not set up!")
        return False
    
    try:
        root = tk.Tk()
        root.withdraw()
        password = simpledialog.askstring("Password", "Enter password:", show='*')
        root.destroy()
        
        if password is None:
            logging.warning("Password input cancelled.")
            return False
        result = hash_password(password) == config.get("password_hash")
        logging.info("Password authentication result: %s", result)
        return result
    except Exception as e:
        logging.error("Error during password authentication: %s", str(e))
        messagebox.showerror("Error", f"Password authentication failed: {e}")
        return False

# Authenticate face
def authenticate_face():
    config = load_config()
    if not config:
        logging.error("Face authentication failed: Application not set up.")
        messagebox.showerror("Error", "Application not set up!")
        return False
    
    try:
        video_capture = cv2.VideoCapture(0)
        if not video_capture.isOpened():
            logging.error("Cannot open camera for face authentication.")
            messagebox.showerror("Error", "Cannot access camera.")
            return False
        
        start_time = time.time()
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Face Authentication", "Preparing to scan face in 3 seconds...")
        root.destroy()
        
        while True:
            ret, frame = video_capture.read()
            if not ret:
                logging.error("Cannot capture image from camera.")
                messagebox.showerror("Error", "Cannot access camera.")
                video_capture.release()
                return False
            cv2.imshow('Video', frame)

            if time.time() - start_time > 3:
                try:
                    result = DeepFace.verify(frame, config.get("reference_image"), model_name="VGG-Face")
                    logging.info("Face authentication result: %s", result)
                    if result["verified"]:
                        video_capture.release()
                        cv2.destroyAllWindows()
                        return True
                    else:
                        logging.warning("Face does not match.")
                        messagebox.showerror("Error", "Face does not match!")
                except Exception as e:
                    logging.error("Face recognition error: %s", str(e))
                    messagebox.showerror("Error", f"Face recognition error: {e}")
                    video_capture.release()
                    cv2.destroyAllWindows()
                    return False

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        video_capture.release()
        cv2.destroyAllWindows()
        return False
    except Exception as e:
        logging.error("Error during face authentication: %s", str(e))
        messagebox.showerror("Error", f"Face authentication failed: {e}")
        return False

# Lock and unlock folder
def protect_folder(folder_path):
    try:
        os.system(f'icacls "{folder_path}" /inheritance:d')
        os.system(f'icacls "{folder_path}" /deny "Everyone:(OI)(CI)(RX)"')
        logging.info("Locked folder: %s", folder_path)
    except Exception as e:
        logging.error("Error locking folder %s: %s", folder_path, str(e))

def unprotect_folder(folder_path):
    try:
        os.system(f'icacls "{folder_path}" /reset')
        time.sleep(2.0)  # Increase delay to ensure permissions update
        # Verify permissions
        result = subprocess.run(f'icacls "{folder_path}"', capture_output=True, text=True, shell=True)
        logging.info("Permissions after unlocking %s: %s", folder_path, result.stdout)
        if "DENY" in result.stdout.upper():
            logging.warning("DENY permission still present after reset: %s", folder_path)
            os.system(f'icacls "{folder_path}" /grant "Everyone:(OI)(CI)(RX)"')
            result = subprocess.run(f'icacls "{folder_path}"', capture_output=True, text=True, shell=True)
            logging.info("Permissions after granting %s: %s", folder_path, result.stdout)
        logging.info("Unlocked folder: %s", folder_path)
    except Exception as e:
        logging.error("Error unlocking folder %s: %s", folder_path, str(e))

# Refresh File Explorer
def refresh_explorer(folder_path):
    try:
        folder_path = os.path.normpath(folder_path)
        win32api.SHChangeNotify(win32con.SHCNE_UPDATEDIR, win32con.SHCNF_PATH, folder_path)
        logging.info("Refreshed File Explorer for: %s", folder_path)
    except Exception as e:
        logging.error("Error refreshing File Explorer: %s", str(e))

# Add shell open command
def add_shell_open_command():
    try:
        # Use .exe path if packaged
        if hasattr(sys, '_MEIPASS'):
            script_path = sys.executable
        else:
            script_path = os.path.abspath(__file__)
        command = f'"{script_path}" "%1"'
        key_path = r"Directory\shell\open\command"
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
        winreg.CloseKey(key)
        logging.info("Customized shell open command.")
    except Exception as e:
        logging.error("Error customizing shell open command: %s", str(e))

# Restore default shell open command
def restore_shell_open_command():
    try:
        key_path = r"Directory\shell\open\command"
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "")
        winreg.CloseKey(key)
        logging.info("Restored default shell open command.")
    except Exception as e:
        logging.error("Error restoring shell open command: %s", str(e))

# Select and save protected folders
def select_protected_folders():
    try:
        root = tk.Tk()
        root.withdraw()
        folder = filedialog.askdirectory(title="Select Folder to Protect")
        root.destroy()
        return folder if folder else None
    except Exception as e:
        logging.error("Error selecting folder: %s", str(e))
        return None

# Load protected folders
def load_protected_folders():
    try:
        if os.path.exists(PROTECTED_FOLDERS_FILE):
            with open(PROTECTED_FOLDERS_FILE, 'r') as f:
                folders = [os.path.normpath(folder) for folder in json.load(f)]
                logging.info("Loaded protected folders: %s", folders)
                return folders
        return []
    except Exception as e:
        logging.error("Error loading protected folders: %s", str(e))
        return []

# Save protected folders
def save_protected_folders(folders):
    try:
        with open(PROTECTED_FOLDERS_FILE, 'w') as f:
            json.dump(folders, f)
        logging.info("Saved protected folders: %s", folders)
    except Exception as e:
        logging.error("Error saving protected folders: %s", str(e))

# Open folder
def open_folder(folder_path):
    try:
        folder_path = os.path.normpath(folder_path)  # Normalize path
        logging.info("Preparing to open folder: %s", folder_path)
        if os.path.exists(folder_path):
            # Retry opening folder up to 3 times
            for attempt in range(1):
                result = subprocess.run(f'explorer "{folder_path}"', shell=True, capture_output=True, text=True)
                if result.returncode == 0:
                    logging.info("Successfully opened folder (attempt %d): %s", attempt+1, folder_path)
                    refresh_explorer(folder_path)  # Refresh File Explorer
                    return
                logging.warning("Failed to open folder (attempt %d): %s, waiting 0.5 seconds", attempt+1, folder_path)
                time.sleep(0.5)
            # logging.error("Could not open folder after 3 attempts: %s", folder_path)
            # messagebox.showerror("Error", f"Cannot open folder {folder_path}. Please check permissions.")
        else:
            logging.error("Folder does not exist: %s", folder_path)
            messagebox.showerror("Error", f"Folder {folder_path} does not exist!")
    except Exception as e:
        logging.error("Error opening folder %s: %s", folder_path, str(e))
        messagebox.showerror("Error", f"Failed to open folder: {e}")

# Verify access to folder
def verify_access(folder_path):
    folder_path = os.path.normpath(folder_path)  # Normalize path
    logging.info("Authenticating for folder: %s", folder_path)
    protected_folders = load_protected_folders()
    if folder_path in protected_folders:
        if authenticate_password():
            logging.info("Password correct, starting face scan.")
            if authenticate_face():
                logging.info("Authentication successful, opening folder: %s", folder_path)
                pyttsx3.speak("Access granted!")
                unprotect_folder(folder_path)
                open_folder(folder_path)
                protect_folder(folder_path)
            else:
                logging.warning("Face does not match.")
                pyttsx3.speak("Access denied.")
        else:
            logging.warning("Incorrect password.")
            pyttsx3.speak("Incorrect password. Access denied.")
    else:
        logging.info("Folder not protected, opening normally: %s", folder_path)
        open_folder(folder_path)

# Windows Service
class FolderProtectorService(win32serviceutil.ServiceFramework):
    _svc_name_ = "FolderProtectorService"
    _svc_display_name_ = "Folder Protector Service"
    _svc_description_ = "Protects folders with password and face recognition."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.running = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        while self.running:
            protected_folders = load_protected_folders()
            for folder in protected_folders:
                if os.path.exists(folder):
                    protect_folder(folder)
            time.sleep(10)  # Check every 10 seconds

# Show application info
def show_about():
    about_window = Toplevel()
    about_window.title("About Folder Protector")
    about_window.geometry("400x300")
    
    Label(about_window, text="Folder Protector", font=("Arial", 16, "bold")).pack(pady=10)
    Label(about_window, text="Version 1.0").pack()
    Label(about_window, text="\nA secure folder protection application using\npassword and face recognition.").pack(pady=10)
    
    Label(about_window, text="\nDeveloped with ❤️").pack()
    
    Button(about_window, text="Close", command=about_window.destroy).pack(pady=20)

# Main UI
def run_protection():
    if not is_setup_complete():
        logging.info("First run detected, starting setup.")
        setup_application()
        if not is_setup_complete():
            logging.error("Setup not completed.")
            messagebox.showerror("Error", "Setup not completed. Please run again.")
            return
    else:
        logging.info("Setup already completed, skipping setup.")
    
    root = tk.Tk()
    root.title("Folder Protector")
    root.geometry("500x450")  # Increased height for new buttons

    protected_folders = load_protected_folders()

    # Folder list
    folder_listbox = Listbox(root, width=50, height=10)
    folder_listbox.pack(pady=10)

    def update_folder_list():
        folder_listbox.delete(0, tk.END)
        for folder in protected_folders:
            folder_listbox.insert(tk.END, folder)

    def add_folder():
        folder = select_protected_folders()
        if folder and folder not in protected_folders:
            folder = os.path.normpath(folder)
            protected_folders.append(folder)
            save_protected_folders(protected_folders)
            protect_folder(folder)
            update_folder_list()
            messagebox.showinfo("Info", f"Added folder: {folder}")
        elif folder in protected_folders:
            messagebox.showwarning("Warning", "Folder already protected!")

    def remove_folder():
        selected = folder_listbox.curselection()
        if selected:
            folder = folder_listbox.get(selected[0])
            unprotect_folder(folder)
            protected_folders.remove(folder)
            save_protected_folders(protected_folders)
            update_folder_list()
            messagebox.showinfo("Info", f"Removed folder: {folder}")
        else:
            messagebox.showwarning("Warning", "Please select a folder to remove!")

    def open_selected_folder():
        selected = folder_listbox.curselection()
        if selected:
            folder = folder_listbox.get(selected[0])
            verify_access(folder)
        else:
            messagebox.showwarning("Warning", "Please select a folder to open!")

    # Folder management frame
    folder_frame = Frame(root)
    folder_frame.pack(pady=5)
    
    Button(folder_frame, text="Add Folder", command=add_folder).pack(side=tk.LEFT, padx=5)
    Button(folder_frame, text="Remove Folder", command=remove_folder).pack(side=tk.LEFT, padx=5)
    Button(folder_frame, text="Open Folder", command=open_selected_folder).pack(side=tk.LEFT, padx=5)
    
    # Settings frame
    settings_frame = Frame(root)
    settings_frame.pack(pady=10)
    
    Label(settings_frame, text="Security Settings", font=("Arial", 10, "bold")).pack(pady=5)
    
    Button(settings_frame, text="Change Password", command=change_password).pack(pady=5)
    Button(settings_frame, text="Update Face Image", command=update_face_image).pack(pady=5)
    
    # Bottom frame
    bottom_frame = Frame(root)
    bottom_frame.pack(pady=10)
    
    Button(bottom_frame, text="About", command=show_about).pack(side=tk.LEFT, padx=10)
    Button(bottom_frame, text="Exit", command=root.quit).pack(side=tk.LEFT, padx=10)

    # Update initial folder list
    update_folder_list()

    # Add shell open command
    add_shell_open_command()

    root.mainloop()

# Run program
if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            if sys.argv[1] == "service":
                win32serviceutil.HandleCommandLine(FolderProtectorService)
            elif sys.argv[1] == "restore":
                restore_shell_open_command()
            else:
                folder_path = sys.argv[1]
                verify_access(folder_path)
        else:
            run_protection()
    except Exception as e:
        logging.error("Main program error: %s", str(e))
        messagebox.showerror("Error", f"Program failed: {e}")