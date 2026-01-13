import customtkinter as ctk
from tkinter import filedialog
from threading import Thread
import pandas as pd
from faster_whisper import WhisperModel
import os
import subprocess
from openpyxl import load_workbook
from openpyxl.styles import Alignment

class TranscriberApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI Video Transcriber Pro")
        self.geometry("600x550")

        # Control Variables
        self.is_running = False
        self.cancel_requested = False

        # UI Layout
        self.label = ctk.CTkLabel(self, text="Step 1: Select your video file", font=("Arial", 16))
        self.label.pack(pady=20)

        self.btn_select = ctk.CTkButton(self, text="Select Video", command=self.select_file)
        self.btn_select.pack(pady=10)

        # Cancel Button (Hidden by default)
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
            # UI Refresh
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
            self.log("!!! Cancellation requested. Cleaning up...")
            self.btn_cancel.configure(state="disabled", text="Stopping...")

    def log(self, message):
        self.status_log.insert("end", message + "\n")
        self.status_log.see("end")

    def process_video(self, video_path):
        audio_temp = "temp_audio.mp3"
        try:
            # 1. STRIP AUDIO (Speed Improvement)
            self.log("Step 1: Stripping audio (this is 100x faster than reading 5GB video)...")
            subprocess.run([
                'ffmpeg', '-i', video_path, 
                '-vn', '-ar', '16000', '-ac', '1', 
                '-ab', '128k', '-f', 'mp3', '-y', audio_temp
            ], check=True, capture_output=True)

            if self.cancel_requested: return

            # 2. LOAD MODEL
            self.log("Step 2: Loading AI Model...")
            model = WhisperModel("base", device="cpu", compute_type="int8")
            
            # 3. TRANSCRIBE
            self.log("Step 3: Transcribing audio...")
            segments, info = model.transcribe(audio_temp, beam_size=5)
            
            total_duration = info.duration
            transcript_data = []
            
            for segment in segments:
                if self.cancel_requested:
                    self.log("Task stopped by user.")
                    return

                transcript_data.append(segment)
                progress = min(segment.end / total_duration, 0.99)
                self.progress_bar.set(progress)
                self.progress_label.configure(text=f"Progress: {int(progress * 100)}%")

            self.progress_bar.set(1.0)
            self.progress_label.configure(text="Progress: 100%")
            
            # 4. SAVE TO TEMPLATE
            self.log("Step 4: Mapping to script template structure...")
            output_file = video_path.rsplit(".", 1)[0] + "_script.xlsx"
            self.save_to_template(video_path, transcript_data, total_duration, output_file)
            
            self.log(f"Success! Script saved at: {output_file}")
            
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            if os.path.exists(audio_temp):
                os.remove(audio_temp)
            self.is_running = False
            self.btn_select.configure(state="normal")
            self.btn_cancel.pack_forget()
            self.btn_cancel.configure(state="normal", text="Cancel Task")

    def save_to_template(self, video_path, segments, total_time, output_path):
        video_name = os.path.basename(video_path)
        fmt = lambda s: f"{int(s//3600):02}:{int((s%3600)//60):02}:{int(s%60):02}"
        
        # Structure matching your CSV/Excel sample:
        # A=Video, E=Start, F=End, K=Phrase
        rows = []
        for start in range(0, int(total_time), 5):
            end = min(start + 5, int(total_time))
            text = " ".join([s.text for s in segments if s.start >= start and s.start < end]).strip()
            
            row = [
                video_name, # Col A
                "", "", "", # Col B, C, D
                fmt(start), # Col E
                fmt(end),   # Col F
                "00:00:05", # Col G
                "", "", "", # Col H, I, J
                text,       # Col K (Phrase)
                "", "", ""  # Col L, M, N
            ]
            rows.append(row)

        # Write data starting from Row 4 to leave room for instruction headers
        df = pd.DataFrame(rows)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False, startrow=3)

        # Professional Alignment & Column Widths
        wb = load_workbook(output_path)
        ws = wb.active
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['K'].width = 80 # Expand for text readability

        for row in ws.iter_rows(min_row=4):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        
        wb.save(output_path)

if __name__ == "__main__":
    app = TranscriberApp()
    app.mainloop()