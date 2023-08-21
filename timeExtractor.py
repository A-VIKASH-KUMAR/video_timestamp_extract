import tkinter as tk
from tkinter import ttk, filedialog
import cv2
from datetime import datetime, timedelta
from openpyxl import Workbook
import threading
from PIL import Image, ImageTk

class VideoPlayerApp:
    def __init__(self, root, video_path="timestamp_test.mp4"):
        self.root = root
        self.video_path = video_path
        self.cap = cv2.VideoCapture(self.video_path)
        self.desired_width = 640
        self.desired_height = 480
        self.video_frame = ttk.Frame(root)
        self.video_frame.pack()
        self.video_canvas = tk.Canvas(self.video_frame)
        self.video_canvas.pack()
        
        self.open_button = ttk.Button(root, text="Open Video File", command=self.open_video_file)
        self.open_button.pack()

        self.play_button = ttk.Button(root, text="Play", command=self.play_video)
        self.play_button.pack()

        self.pause_timestamp_button = ttk.Button(root, text="Pause and Extract Timestamp", command=self.toggle_pause_timestamp)
        self.pause_timestamp_button.pack()

        self.playback_speed = 1.0
        self.paused = False
        self.prev_time = 0
        self.update_interval = 33
        self.video_loop_active = False
        self.pause_flag = False
        self.current_frame = None

        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.excel_filename = 'timestamps.xlsx'

    def play_video(self):
        if not self.video_loop_active:
            self.video_loop_active = True
            self.prev_time = cv2.getTickCount()
            self.video_loop()

    def video_loop(self):
        if self.video_loop_active:
            ret, frame = self.cap.read()
            if ret:
                current_time = cv2.getTickCount()
                elapsed_time = (current_time - self.prev_time) / cv2.getTickFrequency()

                if elapsed_time > 1.0 / self.cap.get(cv2.CAP_PROP_FPS) and not self.paused:
                    resized_frame = cv2.resize(frame, (self.desired_width, self.desired_height))
                    self.current_frame = resized_frame
                    self.display_frame(self.current_frame)
                    self.prev_time = current_time
            
            key = cv2.waitKey(1)
            if key == ord('q'):
                self.stop_video()
                return
            delay = int(self.update_interval / self.playback_speed)
            if not self.pause_flag:
                self.root.after(delay, self.video_loop)

    def display_frame(self, frame):
        img = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img_pil = Image.fromarray(img)
        self.photo = ImageTk.PhotoImage(image=img_pil)
        self.video_canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

    def extract_timestamp(self):
        if self.cap.isOpened():
            timestamp_seconds = self.cap.get(cv2.CAP_PROP_POS_MSEC) / 1000
            formatted_timestamp = self.format_timestamp(timestamp_seconds)
            self.insert_to_excel(formatted_timestamp)
    
    def insert_to_excel(self, timestamp):
        self.sheet.append([timestamp])
        self.workbook.save(self.excel_filename)

    def format_timestamp(self, timestamp_seconds):
        time_delta = timedelta(seconds=timestamp_seconds)
        formatted_time = datetime(1, 1, 1) + time_delta
        return formatted_time.strftime("%H:%M:%S")

    def stop_video(self):
        self.cap.release()
        cv2.destroyAllWindows()
        self.video_loop_active = False

    def toggle_pause_timestamp(self):
        if self.paused:
            self.paused = False
            self.pause_flag = False
            self.video_loop()
            self.extract_timestamp()
        else:
            self.paused = True
            self.pause_flag = True

    def open_video_file(self):
        self.video_path = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4 *.avi")])
        if not self.video_path:
            self.video_path = "timestamp_test.mp4"
        if self.video_path:
            self.cap = cv2.VideoCapture(self.video_path)
            self.play_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = VideoPlayerApp(root, video_path="timestamp_test.mp4")
    root.mainloop()
