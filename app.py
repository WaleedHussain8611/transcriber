import customtkinter as ctk
from tkinter import filedialog
from threading import Thread
import pandas as pd
from faster_whisper import WhisperModel
import os
import math

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class TranscriberApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI Video Transcriber (5GB Support)")
        self.geometry("600x450")

        # UI Layout
        self.label = ctk.CTkLabel(self, text="Step 1: Select your video file", font=("Arial", 16))
        self.label.pack(pady=30)

        self.btn_select = ctk.CTkButton(self, text="Select Video", command=self.select_file)
        self.btn_select.pack(pady=10)

        self.progress_label = ctk.CTkLabel(self, text="Progress: 0%")
        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.set(0)
        
        self.status_log = ctk.CTkTextbox(self, width=500, height=150)
        self.status_log.pack(pady=20)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4 *.mkv *.mov *.avi")])
        if file_path:
            self.btn_select.configure(state="disabled")
            self.progress_bar.pack(pady=10)
            self.progress_label.pack(pady=5)
            # Start transcription in a background thread
            Thread(target=self.process_video, args=(file_path,), daemon=True).start()

    def log(self, message):
        self.status_log.insert("end", message + "\n")
        self.status_log.see("end")

    def process_video(self, video_path):
        try:
            self.log(f"Loading AI Model...")
            # 'int8' is best for local CPU; use 'float16' if you have a GPU
            model = WhisperModel("base", device="cpu", compute_type="int8")
            
            self.log(f"Processing: {os.path.basename(video_path)}")
            segments, info = model.transcribe(video_path, beam_size=5)
            
            total_duration = info.duration
            video_name = os.path.basename(video_path)
            
            # Transcription results storage
            transcript_data = []
            
            # Iterate through segments and update progress
            for segment in segments:
                transcript_data.append(segment)
                # Update progress bar based on current segment timestamp
                progress = min(segment.end / total_duration, 1.0)
                self.progress_bar.set(progress)
                self.progress_label.configure(text=f"Progress: {int(progress * 100)}%")

            self.log("Aligning text to 5-second intervals...")
            final_rows = self.map_to_intervals(video_name, transcript_data, total_duration)
            
            # Export to Excel
            df = pd.DataFrame(final_rows, columns=['video name', 'start time (hh:mm:ss)', 'end time (hh:mm:ss)', 'transcription'])
            output_file = video_path.rsplit(".", 1)[0] + "_transcription.xlsx"
            df.to_excel(output_file, index=False)
            
            self.log(f"Success! Saved to: {output_file}")
            self.label.configure(text="Transcription Complete!")
            
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.btn_select.configure(state="normal")

    def map_to_intervals(self, name, segments, total_time):
        rows = []
        fmt = lambda s: f"{int(s//3600):02}:{int((s%3600)//60):02}:{int(s%60):02}"
        
        for start in range(0, int(total_time), 5):
            end = min(start + 5, int(total_time))
            # Find all text that overlaps with this 5s window
            text_bits = [s.text for s in segments if s.start >= start and s.start < end]
            combined_text = " ".join(text_bits).strip()
            rows.append([name, fmt(start), fmt(end), combined_text])
        return rows

if __name__ == "__main__":
    app = TranscriberApp()
    app.mainloop()