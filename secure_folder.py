import cv2
import pyttsx3
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox, Listbox
import json
import os
import winreg
import sys
from deepface import DeepFace
import time

# Khởi tạo camera và âm thanh
video_capture = cv2.VideoCapture(0)
if not video_capture.isOpened():
    print("Lỗi: Không thể mở camera.")
    exit()
engine = pyttsx3.init()

# File lưu danh sách thư mục được bảo vệ
CONFIG_FILE = "protected_folders.json"

# Hàm xác thực mật khẩu
def authenticate_password():
    root = tk.Tk()
    root.withdraw()
    password = simpledialog.askstring("Password", "Nhập mật khẩu:", show='*')
    return password == "123456"  # Thay đổi mật khẩu nếu cần

# Hàm xác thực khuôn mặt
def authenticate_face():
    start_time = time.time()
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Xác thực khuôn mặt", "Chuẩn bị quét khuôn mặt trong 5 giây...")
    
    while True:
        ret, frame = video_capture.read()
        if not ret:
            print("Không thể chụp ảnh từ camera.")
            messagebox.showerror("Lỗi", "Không thể truy cập camera.")
            break
        cv2.imshow('Video', frame)

        if time.time() - start_time > 5:
            try:
                result = DeepFace.verify(frame, "image.jpg", model_name="VGG-Face")
                print("Kết quả xác thực:", result)
                if result["verified"]:
                    return True
                else:
                    messagebox.showerror("Lỗi", "Khuôn mặt không khớp!")
            except Exception as e:
                print("Lỗi nhận diện:", e)
                messagebox.showerror("Lỗi", f"Lỗi nhận diện khuôn mặt: {e}")

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    video_capture.release()
    cv2.destroyAllWindows()
    return False

# Hàm khóa và mở khóa thư mục
def protect_folder(folder_path):
    try:
        os.system(f'icacls "{folder_path}" /inheritance:d')
        os.system(f'icacls "{folder_path}" /deny "Everyone:(OI)(CI)(RX)"')
        print(f"Đã khóa thư mục: {folder_path}")
    except Exception as e:
        print(f"Lỗi khóa thư mục: {e}")

def unprotect_folder(folder_path):
    try:
        os.system(f'icacls "{folder_path}" /reset')
        print(f"Đã mở khóa thư mục: {folder_path}")
    except Exception as e:
        print(f"Lỗi mở khóa thư mục: {e}")

# Hàm thêm context menu
def add_context_menu():
    try:
        script_path = os.path.abspath(__file__)
        command = f'python "{script_path}" "%1"'
        key_path = r"Directory\shell\OpenProtectedFolder"
        key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Mở thư mục bảo vệ")
        winreg.CreateKey(key, "command")
        command_key = winreg.OpenKey(key, "command", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(command_key, "", 0, winreg.REG_SZ, command)
        winreg.CloseKey(command_key)
        winreg.CloseKey(key)
        print("Đã thêm context menu.")
    except Exception as e:
        print(f"Lỗi thêm context menu: {e}")

# Hàm chọn và lưu thư mục được bảo vệ
def select_protected_folders():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Chọn thư mục để bảo vệ")
    return folder if folder else None

# Hàm tải danh sách thư mục được bảo vệ
def load_protected_folders():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return []

# Hàm lưu danh sách thư mục
def save_protected_folders(folders):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(folders, f)

# Hàm mở thư mục
def open_folder(folder_path):
    if os.path.exists(folder_path):
        os.system(f'explorer "{folder_path}"')
    else:
        messagebox.showerror("Lỗi", f"Thư mục {folder_path} không tồn tại!")

# Hàm xác thực để mở thư mục
def verify_access(folder_path):
    protected_folders = load_protected_folders()
    if folder_path in protected_folders:
        if authenticate_password():
            print("Mật khẩu đúng! Đang quét khuôn mặt...")
            if authenticate_face():
                print(f"Nhận diện khuôn mặt thành công. Mở thư mục {folder_path}!")
                pyttsx3.speak("Truy cập được cấp!")
                unprotect_folder(folder_path)
                open_folder(folder_path)
                protect_folder(folder_path)
            else:
                print("Khuôn mặt không được nhận diện. Truy cập bị từ chối.")
                pyttsx3.speak("Truy cập bị từ chối.")
        else:
            print("Mật khẩu sai. Truy cập bị từ chối.")
            pyttsx3.speak("Mật khẩu sai. Truy cập bị từ chối.")
    else:
        open_folder(folder_path)

# Giao diện chính
def run_protection():
    root = tk.Tk()
    root.title("Bảo vệ thư mục")
    root.geometry("400x300")

    protected_folders = load_protected_folders()

    # Danh sách thư mục
    folder_listbox = Listbox(root, width=50, height=10)
    folder_listbox.pack(pady=10)

    def update_folder_list():
        folder_listbox.delete(0, tk.END)
        for folder in protected_folders:
            folder_listbox.insert(tk.END, folder)

    def add_folder():
        folder = select_protected_folders()
        if folder and folder not in protected_folders:
            protected_folders.append(folder)
            save_protected_folders(protected_folders)
            protect_folder(folder)
            update_folder_list()
            messagebox.showinfo("Thông báo", f"Đã thêm thư mục: {folder}")
        elif folder in protected_folders:
            messagebox.showwarning("Cảnh báo", "Thư mục đã được bảo vệ!")

    def remove_folder():
        selected = folder_listbox.curselection()
        if selected:
            folder = folder_listbox.get(selected[0])
            unprotect_folder(folder)
            protected_folders.remove(folder)
            save_protected_folders(protected_folders)
            update_folder_list()
            messagebox.showinfo("Thông báo", f"Đã xóa thư mục: {folder}")
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục để xóa!")

    def open_selected_folder():
        selected = folder_listbox.curselection()
        if selected:
            folder = folder_listbox.get(selected[0])
            verify_access(folder)
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục để mở!")

    # Nút điều khiển
    tk.Button(root, text="Thêm thư mục", command=add_folder).pack(pady=5)
    tk.Button(root, text="Xóa thư mục", command=remove_folder).pack(pady=5)
    tk.Button(root, text="Mở thư mục", command=open_selected_folder).pack(pady=5)
    tk.Button(root, text="Thoát", command=root.quit).pack(pady=5)

    # Cập nhật danh sách ban đầu
    update_folder_list()

    # Thêm context menu
    add_context_menu()

    root.mainloop()

# Chạy chương trình
if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
        verify_access(folder_path)
    else:
        run_protection()