import tkinter as tk
from tkinter import ttk, filedialog
import cv2
from datetime import datetime, timedelta
from openpyxl import Workbook
import threading
from PIL import Image, ImageTk

workbook = Workbook()
sheet = workbook.active

class VideoPlayerApp:
    def __init__(self, root, video_path = "timestamp_test.mp4"):
        self.root = root
        if video_path == "timestamp_test.mp4":
            self.video_path = video_path  # Initialize video path
            self.cap = cv2.VideoCapture(self.video_path)
        else:
            self.video_path = None  # Initialize video path
            self.cap = None
        self.desired_width = 640  # Specify the desired width
        self.desired_height = 480  # Specify the desired height
        
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

        self.playback_speed = 1.0  # Initial playback speed
        self.fast_forward_button = ttk.Button(root, text="Fast Forward", command=self.fast_forward)
        self.fast_forward_button.pack()

        self.slow_down_button = ttk.Button(root, text="Slow Down", command=self.slow_down)
        self.slow_down_button.pack()

        


        self.paused = False
        self.prev_time = 0
        self.update_interval = 33  # milliseconds (30 frames per second)
        self.video_loop_active = False
        self.pause_flag = False  # New flag to handle pausing and resuming
        self.current_frame = None

    # ... (rest of the class)

    def play_video(self):
        if not self.video_loop_active:
            self.video_loop_active = True
            self.prev_time = cv2.getTickCount()
            self.video_loop()

    def fast_forward(self):
        self.playback_speed *= 2.0  # Double the playback speed
        print("Playback speed:", self.playback_speed)

    def slow_down(self):
        self.playback_speed *= 0.5  # Halve the playback speed
        print("Playback speed:", self.playback_speed)

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
            # If paused, set pause_flag and do not re-call the video_loop
            if not self.pause_flag:
                self.root.after(delay, self.video_loop)  # Call video_loop again after a delay

    def display_frame(self, frame):
        img = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img_pil = Image.fromarray(img)
        self.photo = ImageTk.PhotoImage(image=img_pil)
        self.video_canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

    def convert_to_tk_image(self, frame):
        height, width, channels = frame.shape
        img_bgr = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
        img_bytes = img_bgr.tobytes()
        return tk.PhotoImage(width=width, height=height, data=img_bytes)

    def stop_video(self):
        self.cap.release()
        cv2.destroyAllWindows()
        self.video_loop_active = False

    def extract_timestamp(self):
        if self.cap.isOpened():
            timestamp_seconds = self.cap.get(cv2.CAP_PROP_POS_MSEC) / 1000  # Convert milliseconds to seconds
            formatted_timestamp = self.format_timestamp(timestamp_seconds)
            self.insert_to_excel(formatted_timestamp)

    def format_timestamp(self, timestamp_seconds):
        time_delta = timedelta(seconds=timestamp_seconds)
        formatted_time = datetime(1, 1, 1) + time_delta
        return formatted_time.strftime("%H:%M:%S")
    
    def insert_to_excel(self, timestamp):
        # Insert the timestamp into the Excel sheet
        sheet.append([timestamp])
        workbook.save('timestamps.xlsx')

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
            # Provide a default video file path if none selected
            self.video_path = "timestamp_test.mp4"
        if self.video_path:
            self.cap = cv2.VideoCapture(self.video_path)
            self.play_button.config(state=tk.NORMAL) 

if __name__ == "__main__":
    root = tk.Tk()
    app = VideoPlayerApp(root,video_path="timestamp_test.mp4")
    root.mainloop()
