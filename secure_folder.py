from deepface import DeepFace
import cv2
import pyttsx3
import tkinter as tk
from tkinter import simpledialog
import time

video_capture = cv2.VideoCapture(0)
engine = pyttsx3.init()

def authenticate_password():
    root = tk.Tk()
    root.withdraw()
    password = simpledialog.askstring("Password", "Enter password:", show='*')
    return password == "Sanghustk63"

def authenticate_face():
    start_time = time.time()
    while True:
        ret, frame = video_capture.read()
        if not ret:
            print("Không thể chụp ảnh từ camera.")
            break
        cv2.imshow('Video', frame)

        if time.time() - start_time > 5:
            try:
                result = DeepFace.verify(frame, "image.jpg", model_name="VGG-Face")
                print("Verification result:", result)
                if result["verified"]:
                    return True
            except Exception as e:
                print("Lỗi nhận diện:", e)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    video_capture.release()
    cv2.destroyAllWindows()
    return False

def open_folder():
    import os
    os.system('explorer "C:\\Users\\User\\Documents\\Secret"')

def run_authentication():
    print("Vui lòng nhập mật khẩu trước.")
    if authenticate_password():
        print("Mật khẩu đúng! Đang quét khuôn mặt...")
        if authenticate_face():
            print("Nhận diện khuôn mặt thành công. Truy cập được cấp!")
            pyttsx3.speak("Truy cập được cấp!")
            open_folder()
        else:
            print("Khuôn mặt không được nhận diện. Truy cập bị từ chối.")
            pyttsx3.speak("Truy cập bị từ chối.")
    else:
        print("Mật khẩu sai. Truy cập bị từ chối.")
        pyttsx3.speak("Mật khẩu sai. Truy cập bị từ chối.")

run_authentication()