# MyAI — Personal NVIDIA AI companion

A redesigned Streamlit chat app with a premium dark interface, NVIDIA Nemotron streaming, file/image attachments, microphone recording with optional speech transcription, browser text-to-speech, and adaptive conversational behavior.

## Deploy on Streamlit Community Cloud

1. Upload `app.py`, `requirements.txt`, and `README.md` to your GitHub repository.
2. Create/open the app at https://share.streamlit.io/.
3. In app settings → **Secrets**, add:

```toml
NVIDIA_API_KEY = "nvapi-your-real-key"
NVIDIA_MODEL = "nvidia/nemotron-3.5-super-vl-120b-a12b"
```

4. Set `app.py` as the main file and deploy/reboot.

You can alternatively enter the key in the sidebar under **AI connection**. That key remains in the active Streamlit session; using Streamlit Secrets is safer for deployment. Never put a real key in GitHub or share it in screenshots.

## Features

- ChatGPT-inspired responsive dark interface and conversation history during the active session.
- Streaming responses from NVIDIA's OpenAI-compatible chat completions endpoint.
- File attachments within the chat workspace: PDF, DOCX, XLSX, CSV, text, code, and images.
- Microphone recording through Streamlit's audio input widget.
- Optional speech-to-text via the SpeechRecognition Google Web Speech recognizer.
- Browser-native speech playback using SpeechSynthesis; voice availability depends on the user's browser/device.
- Adaptive personal companion instructions: naturally warm and affectionate when the user leads that way, without forcing named mood modes.
- Export current conversation as TXT.

## Important notes

- NVIDIA API key stays server-side; use Streamlit Secrets. Uploaded text and images are sent to NVIDIA for processing.
- Microphone transcription uses Google's web speech recognition service through SpeechRecognition; it requires internet and may not be available in every environment. Do not use it for sensitive audio.
- Voice playback uses the browser's built-in speech synthesis voices; it is not a dedicated real-time two-way voice engine.
- Free NVIDIA API endpoints have trial quotas, rate limits, availability and model lifecycle changes. Free does not mean unlimited.
- Streamlit Community Cloud free apps may sleep after inactivity and do not guarantee 24/7 uptime.
- Chat history is session-based, not a permanent database. A page refresh or a new session may clear it.
- The app does not execute generated code or control your computer.
