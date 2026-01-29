# AI Video Transcriber Pro

**Powered by Wazza Studios**

AI Video Transcriber Pro is a modern, user-friendly desktop application that leverages OpenAI's Whisper model (via `faster-whisper`) to generate accurate transcriptions from video files. It automatically exports the transcription to a formatted Excel file, perfect for content creators, editors, and archivists.

## Features

-   **High-Accuracy Transcription**: Uses the `faster-whisper` implementation for efficient and accurate AI speech-to-text.
-   **Excel Export**: Automatically generates an Excel file (`.xlsx`) with timestamps, separating speech into 5-second chunks for easy editing.
-   **Modern UI**: Built with `customtkinter` for a sleek, dark-themed interface.
-   **Local Processing**: Runs entirely on your machine—no data is uploaded to the cloud.
-   **FFmpeg Integration**: Handles audio extraction seamlessly.

## Prerequisites

-   **Python 3.8+**
-   **FFmpeg**: Must be installed and accessible in your system's PATH (or bundled with the app).

## Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd transcriber
    ```

2.  **Create and activate a virtual environment (optional but recommended):**
    ```bash
    python -m venv venv
    # Windows:
    .\venv\Scripts\activate
    # Mac/Linux:
    source venv/bin/activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

    *Note: If `requirements.txt` is missing, you'll likely need:*
    ```bash
    pip install customtkinter faster-whisper pandas openpyxl
    ```

## Usage

1.  **Run the application:**
    ```bash
    python app.py
    ```

2.  **Select a Video:**
    -   Click the "Select a Video" button.
    -   Choose an `.mp4`, `.mkv`, `.mov`, or `.avi` file.

3.  **Wait for Processing:**
    -   The app will extract audio, load the AI model, and transcribe the content.
    -   Progress is displayed in the progress bar and status log.

4.  **View Results:**
    -   Once complete, an Excel file (e.g., `video_script.xlsx`) is created in the same directory as the source video.
    -   The transcription is located in **Column K**, formatted and left-aligned for readability.

## Building Executable

To build a standalone executable using PyInstaller:

```bash
pyinstaller --noconfirm --onedir --windowed --add-data "ffmpeg.exe;." --icon="icon.ico" app.py
```
*(Adjust the command based on your specific assets and OS).*

## License

This project is proprietary software of **Wazza Studios**.
