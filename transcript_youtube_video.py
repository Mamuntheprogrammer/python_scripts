import os
import whisper
import yt_dlp

def download_audio_ytdlp(url, output_path="audio", filename="audio.mp3"):
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    
    output_file = os.path.join(output_path, filename)

    # Skip download if file already exists
    if os.path.exists(output_file):
        print(f"✅ Audio already exists: {output_file}")
        return output_file

    ydl_opts = {
        'format': 'bestaudio/best',
        'outtmpl': output_file,
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
        'quiet': False,
    }

    print("⬇️ Downloading audio...")
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        ydl.download([url])

    # Handle double extension issue
    if os.path.exists(output_file + ".mp3"):
        os.rename(output_file + ".mp3", output_file)

    print(f"✅ Audio downloaded to: {output_file}")
    return output_file

def transcribe_audio_whisper(audio_path, output_txt="transcription.txt"):
    print("🧠 Loading Whisper model...")
    model = whisper.load_model("small")  # Change to "small" or "base" for faster performance

    print("✍️ Transcribing audio...")
    result = model.transcribe(audio_path, language=None)

    print(f"✅ Transcription done. Language detected: {result['language']}")
    transcript = result["text"]

    with open(output_txt, "w", encoding="utf-8") as f:
        f.write(transcript)

    print(f"📝 Transcript saved to: {output_txt}")
    return transcript

if __name__ == "__main__":
    output_path = "audio"
    filename = "audio.mp3"
    audio_path = os.path.join(output_path, filename)

    # Skip YouTube prompt if file already exists
    if os.path.exists(audio_path):
        print(f"📁 Using existing audio file: {audio_path}")
    else:
        youtube_link = input("🔗 Enter YouTube video URL: ").strip()
        try:
            audio_path = download_audio_ytdlp(youtube_link, output_path, filename)
        except Exception as e:
            print(f"❌ Error downloading audio: {e}")
            exit(1)

    try:
        transcribe_audio_whisper(audio_path)
    except Exception as e:
        print(f"❌ Error transcribing audio: {e}")
