import customtkinter as ctk
from tkinter import filedialog
from threading import Thread
import pandas as pd
from faster_whisper import WhisperModel
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Updated the UI to make it more modern and visually appealing
class TranscriberApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Wazza Studios - AI Video Transcriber Pro")
        self.geometry("700x600")
        self.configure(fg_color="#2b2b2b")  # Dark background

        # Control Variables
        self.is_running = False
        self.cancel_requested = False

        # Header
        self.header = ctk.CTkLabel(self, text="Wazza Studios - AI Video Transcriber Pro", font=("Arial", 24, "bold"), text_color="#ffffff")
        self.header.pack(pady=20)

        # Select Video Button
        self.btn_select = ctk.CTkButton(self, text="Select a Video", command=self.select_file, fg_color="#1f6aa5", hover_color="#144e78", text_color="#ffffff")
        self.btn_select.pack(pady=10)

        # Cancel Button (Hidden by default)
        self.btn_cancel = ctk.CTkButton(self, text="Cancel Task", fg_color="#d9534f", hover_color="#c9302c", text_color="#ffffff", command=self.request_cancel)

        # Progress Bar and Label
        self.progress_label = ctk.CTkLabel(self, text="Progress: 0%", font=("Arial", 14), text_color="#d1d1d1")
        self.progress_bar = ctk.CTkProgressBar(self, width=500, progress_color="#1f6aa5", fg_color="#3a3a3a")
        self.progress_bar.set(0)

        # Status Log
        self.status_log = ctk.CTkTextbox(self, width=600, height=200, fg_color="#3a3a3a", text_color="#ffffff", font=("Consolas", 12))
        self.status_log.pack(pady=20)

        # Footer
        self.footer = ctk.CTkLabel(self, text="Powered by Wazza Studios", font=("Arial", 12, "italic"), text_color="#d1d1d1")
        self.footer.pack(side="bottom", pady=10)

    def log(self, message):
        self.status_log.insert("end", message + "\n")
        self.status_log.see("end")

    def request_cancel(self):
        if self.is_running:
            self.cancel_requested = True
            self.log("!!! Cancellation requested. Stopping...")
            self.btn_cancel.configure(state="disabled", text="Stopping...")

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Video files", "*.mp4 *.mkv *.mov *.avi")])
        if file_path:
            # UI Preparation
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

    def process_video(self, video_path):
        try:
            self.log("Loading AI Model...")
            # Optimized for speed on CPU
            model = WhisperModel("base", device="cpu", compute_type="int8", cpu_threads=4)
            
            self.log(f"Transcribing: {os.path.basename(video_path)}")
            segments, info = model.transcribe(video_path, beam_size=1)
            
            total_duration = info.duration
            transcript_data = []
            
            for segment in segments:
                # CHECK FOR CANCEL SIGNAL
                if self.cancel_requested:
                    self.log("Task cancelled.")
                    # RESET PROGRESS BAR ON CANCEL
                    self.after(0, lambda: self.progress_bar.set(0))
                    self.after(0, lambda: self.progress_label.configure(text="Progress: 0%"))
                    return

                transcript_data.append(segment)
                progress = min(segment.end / total_duration, 0.99)
                self.progress_bar.set(progress)
                self.progress_label.configure(text=f"Progress: {int(progress * 100)}%")

            # Finalize success
            self.progress_bar.set(1.0)
            self.progress_label.configure(text="Progress: 100%")
            
            self.log("Writing to script template...")
            output_file = video_path.rsplit(".", 1)[0] + "_script.xlsx"
            self.save_to_template(video_path, transcript_data, total_duration, output_file)
            self.log(f"Success! Saved to: {output_file}")
            
        except Exception as e:
            self.log(f"Error: {str(e)}")
        finally:
            # Cleanup UI state
            self.is_running = False
            self.btn_select.configure(state="normal")
            self.btn_cancel.pack_forget()
            self.btn_cancel.configure(state="normal", text="Cancel Task")

    def save_to_template(self, video_path, segments, total_time, output_path):
        video_name = os.path.basename(video_path)
        fmt = lambda s: f"{int(s//3600):02}:{int((s%3600)//60):02}:{int(s%60):02}"
        
        headers = [
            "Video", "Background music", "Theme", "Guest", "Start time", 
            "End time", "Time length", "audio", "start time", "end time", 
            "Phrase", "Angle", "Comment", "Additional comment"
        ]
        
        rows = []
        for start in range(0, int(total_time), 5):
            end = min(start + 5, int(total_time))
            text = " ".join([s.text for s in segments if s.start >= start and s.start < end]).strip()
            
            row = [video_name, "", "", "", fmt(start), fmt(end), "00:00:05", "", "", "", text, "", "", ""]
            rows.append(row)

        df = pd.DataFrame(rows, columns=headers)
        df.to_excel(output_path, index=False, startrow=3)

        wb = load_workbook(output_path)
        ws = wb.active
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