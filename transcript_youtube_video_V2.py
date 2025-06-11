import os
import yt_dlp
from pydub import AudioSegment
from faster_whisper import WhisperModel

def download_audio_ytdlp(url, output_path="audio", filename="audio"):
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    output_file = os.path.join(output_path, filename)

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

    print("⏬ Downloading audio...")
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        ydl.download([url])

    print(f"✅ Audio downloaded to: {output_file}")
    return output_file

def split_audio(audio_path, chunk_length_ms=30000):  # 30 seconds
    print("🔪 Splitting audio into chunks...")
    audio = AudioSegment.from_file(audio_path)
    chunks = []
    for i in range(0, len(audio), chunk_length_ms):
        chunk = audio[i:i + chunk_length_ms]
        chunk_path = f"{audio_path}_chunk_{i//chunk_length_ms}.mp3"
        chunk.export(chunk_path, format="mp3")
        chunks.append(chunk_path)
    print(f"✅ Created {len(chunks)} chunks.")
    return chunks

def transcribe_chunks(chunk_paths, output_txt="transcription.txt", model_size="small"):
    print("🚀 Loading faster-whisper model...")
    model = WhisperModel(model_size, compute_type="int8")  # Best for CPU

    transcript = ""

    for idx, chunk_path in enumerate(chunk_paths):
        print(f"📝 Transcribing chunk {idx+1}/{len(chunk_paths)}...")
        segments, info = model.transcribe(chunk_path)

        for segment in segments:
            transcript += segment.text.strip() + " "

    with open(output_txt, "w", encoding="utf-8") as f:
        f.write(transcript.strip())

    print(f"🌐 Language: {info.language}")
    print(f"📄 Transcript saved to: {output_txt}")
    return transcript.strip()

if __name__ == "__main__":
    output_audio_path = "audio/audio.mp3"

    try:
        if os.path.exists(output_audio_path):
            print(f"📂 Using existing file: {output_audio_path}")
            audio_file = output_audio_path
        else:
            youtube_link = input("🎥 Enter YouTube video URL: ").strip()
            audio_file = download_audio_ytdlp(youtube_link)

        chunk_paths = split_audio(audio_file)
        transcribe_chunks(chunk_paths)

    except Exception as e:
        print(f"❌ Error: {e}")
