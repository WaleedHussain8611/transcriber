import streamlit as st
import whisper
import pandas as pd
from moviepy import VideoFileClip # New way (v2.0+)
import os

st.title("Local Video-to-Excel Transcriber")

# 1. File Uploader
uploaded_file = st.file_uploader("Upload Video (up to 5GB)", type=["mp4", "mkv", "mov"])

if uploaded_file:
    # Save the uploaded file locally to avoid keeping 5GB in RAM
    with open("temp_video.mp4", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    st.success("Video uploaded! Extracting audio...")

    # 2. Extract Audio
    video = VideoFileClip("temp_video.mp4")
    video.audio.write_audiofile("temp_audio.mp3")
    duration = int(video.duration)
    video.close()

    # 3. Transcribe
    st.info("Transcribing... this may take a while for large files.")
    model = whisper.load_model("base") # Use 'large' for better quality if you have a GPU
    result = model.transcribe("temp_audio.mp3")

    # 4. Process to 5-second intervalsx
    rows = []
    for start in range(0, duration, 5):
        end = min(start + 5, duration)
        
        # Format time
        fmt = lambda s: f"{int(s//3600):02}:{int((s%3600)//60):02}:{int(s%60):02}"
        
        # Get text for this window
        text = " ".join([s['text'] for s in result['segments'] if s['start'] >= start and s['start'] < end])
        rows.append([uploaded_file.name, fmt(start), fmt(end), text.strip()])

    # 5. Export
    df = pd.DataFrame(rows, columns=['video name', 'start time', 'end time', 'transcription'])
    df.to_excel("output.xlsx", index=False)

    with open("output.xlsx", "rb") as file:
        st.download_button("Download Excel Template", data=file, file_name="transcription.xlsx")
    
    # Cleanup local files
    os.remove("temp_video.mp4")
    os.remove("temp_audio.mp3")