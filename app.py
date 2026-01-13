import customtkinter as ctk
from tkinter import filedialog
from threading import Thread
import pandas as pd
from faster_whisper import WhisperModel
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

class TranscriberApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI Video Transcriber Pro")
        self.geometry("600x550")

        self.is_running = False
        self.cancel_requested = False

        self.label = ctk.CTkLabel(self, text="Select a video to begin", font=("Arial", 16))
        self.label.pack(pady=20)

        self.btn_select = ctk.CTkButton(self, text="Select Video", command=self.select_file)
        self.btn_select.pack(pady=10)

        self.btn_cancel = ctk.CTkButton(self, text="Cancel Task", fg_color="#d9534f", 
                                        hover_color="#c9302c", command=self.request_cancel)
        
        self.progress_label = ctk.CTkLabel(self, text="Progress: 0%")
        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.set(0)

        self.status_log = ctk.CTkTextbox(self, width=500, height=180)
        self.status_log.pack(pady=20)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4 *.mkv *.mov *.avi")])
        if file_path:
            self.status_log.delete("1.0", "end") 
            self.progress_bar.set(0)
            self.progress_label.configure(text="Progress: 0%")
            self.btn_select.configure(state="disabled")
            self.btn_cancel.pack(pady=5)
            self.progress_bar.pack(pady=10)
            self.progress_label.pack(pady=5)
            self.is_running = True
            self.cancel_requested = False
            Thread(target=self.process_video, args=(file_path,), daemon=True).start()

    def request_cancel(self):
        if self.is_running:
            self.cancel_requested = True
            self.log("!!! Cancellation requested. Stopping...")
            self.btn_cancel.configure(state="disabled", text="Stopping...")

    def log(self, message):
        self.status_log.insert("end", message + "\n")
        self.status_log.see("end")

    def process_video(self, video_path):
        try:
            self.log("Loading AI Model...")
            model = WhisperModel("base", device="cpu", compute_type="int8")
            self.log(f"Transcribing: {os.path.basename(video_path)}")
            segments, info = model.transcribe(video_path, beam_size=5)
            
            total_duration = info.duration
            transcript_data = []
            for segment in segments:
                if self.cancel_requested:
                    self.log("Task cancelled.")
                    return
                transcript_data.append(segment)
                progress = min(segment.end / total_duration, 0.99)
                self.progress_bar.set(progress)
                self.progress_label.configure(text=f"Progress: {int(progress * 100)}%")

            self.progress_bar.set(1.0)
            self.progress_label.configure(text="Progress: 100%")
            
            self.log("Writing to your exact Template structure...")
            output_file = video_path.rsplit(".", 1)[0] + "_script.xlsx"
            self.save_to_template(video_path, transcript_data, total_duration, output_file)
            self.log(f"Success! Saved to: {output_file}")
            
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            self.is_running = False
            self.btn_select.configure(state="normal")
            self.btn_cancel.pack_forget()
            self.btn_cancel.configure(state="normal", text="Cancel Task")

    def save_to_template(self, video_path, segments, total_time, output_path):
        video_name = os.path.basename(video_path)
        fmt = lambda s: f"{int(s//3600):02}:{int((s%3600)//60):02}:{int(s%60):02}"
        
        # Define the structure based on your provided file
        headers = [
            "Video", "Background music", "Theme", "Guest", "Start time", 
            "End time", "Time length", "audio", "start time", "end time", 
            "Phrase", "Angle", "Comment", "Additional comment"
        ]
        
        rows = []
        for start in range(0, int(total_time), 5):
            end = min(start + 5, int(total_time))
            text = " ".join([s.text for s in segments if s.start >= start and s.start < end]).strip()
            
            # Creating a row that matches the column index of the template
            row = [
                video_name, # A: Video
                "",         # B: Background music
                "",         # C: Theme
                "",         # D: Guest
                fmt(start), # E: Start time
                fmt(end),   # F: End time
                "00:00:05", # G: Time length
                "",         # H: audio
                "",         # I: start time
                "",         # J: end time
                text,       # K: Phrase (Transcription)
                "",         # L: Angle
                "",         # M: Comment
                ""          # N: Additional comment
            ]
            rows.append(row)

        # Write to Excel
        df = pd.DataFrame(rows, columns=headers)
        df.to_excel(output_path, index=False, startrow=3) # Data starts after instructions

        # Formatting
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Re-insert the template instruction headers if needed manually or just style:
        column_widths = {'A': 25, 'E': 12, 'F': 12, 'K': 60}
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        for row in ws.iter_rows(min_row=4):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        
        wb.save(output_path)

if __name__ == "__main__":
    app = TranscriberApp()
    app.mainloop()