# Educational Slide Automation Application

A Python desktop application that automatically creates educational slides from video content.

## Features

- **Hotkey Trigger**: Instant screen capture with Ctrl+V
- **AI-Powered Analysis**: Image analysis with Google Gemini Pro Vision
- **Automatic Slide Creation**: PowerPoint presentations generated automatically
- **Background Operation**: Application runs silently in the background
- **Async Processing**: API calls don't freeze the application
- **Session-Based**: All captures during a session go to one presentation file
- **Concise Summaries**: AI generates brief, focused summaries
- **Multiple Choice Questions**: AI creates test questions with 4 options

## Requirements

- Python 3.10 or higher
- Windows operating system
- Google Gemini API key

## Installation

### 1. Clone the Repository

```bash
git clone <repo-url>
cd SnapLearn
```

### 2. Create Virtual Environment (Recommended)

```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Set Up API Key

Create a `.env` file in the project root directory:

```env
GEMINI_API_KEY=your_api_key_here
```

> Get your Gemini API key from [Google AI Studio](https://makersuite.google.com/app/apikey)

## Usage

### Start the Application

```bash
python main.py
```

### Controls

| Hotkey | Action |
|--------|--------|
| `Ctrl+V` | Capture with AI analysis (summary + multiple choice question) |
| `Ctrl+B` | Direct capture (screenshot only, no AI - faster) |
| `Ctrl+C` | Close the application |

### Workflow

1. **Prepare**: Open a video or educational content on your screen
2. **Capture**: Press `Ctrl+V`
3. **Wait**: The application will automatically:
   - Capture the screen
   - Analyze with Gemini AI
   - Add a slide to the presentation
4. **Repeat**: Add as many slides as you want
5. **Result**: Find your PowerPoint file in the `output/` folder

### Session Behavior

- **Same Session**: All screenshots are added to the same PowerPoint file
- **New Session**: Restarting the app creates a new presentation file

## Project Structure

```
SnapLearn/
├── main.py              # Main application
├── config.py            # Configuration
├── screen_capture.py    # Screen capture module
├── gemini_service.py    # Gemini API integration
├── slide_generator.py   # PowerPoint generation
├── requirements.txt     # Dependencies
├── README.md            # This file
├── screenshots/         # Screenshots (auto-created)
└── output/              # PowerPoint files (auto-created)
```

## Configuration

You can modify these settings in `config.py`:

| Setting | Default | Description |
|---------|---------|-------------|
| `HOTKEY` | `ctrl+v` | AI analysis hotkey |
| `HOTKEY_DIRECT` | `ctrl+b` | Direct capture hotkey |
| `GEMINI_MODEL` | `gemini-2.5-flash` | AI model to use |
| `SCREENSHOT_FORMAT` | `png` | Image format |
| `SCREENSHOT_QUALITY` | `95` | JPEG quality (1-100) |

## Troubleshooting

### "GEMINI_API_KEY is not set" error

Make sure the `.env` file exists in the project root directory with the correct format:

```env
GEMINI_API_KEY=AIza...
```

### Screenshot capture fails

- Try running the application as administrator
- Make sure antivirus software isn't blocking the application

### API call fails

- Check your internet connection
- Verify your API key is valid
- Check your API quota at [Google AI Studio](https://makersuite.google.com/)

## Output Example

The application creates a new PowerPoint and PDF file for each session:

```
output/
├── presentation_2024-01-15_14-30-45.pptx
├── presentation_2024-01-15_14-30-45.pdf
├── presentation_2024-01-16_09-15-22.pptx
├── presentation_2024-01-16_09-15-22.pdf
└── ...
```

> Note: PDF export requires Microsoft PowerPoint to be installed.

## Slide Formats

### AI Analysis (Ctrl+V) - Creates 2 slides:
1. **Slide 1**: Full-screen screenshot
2. **Slide 2**: Summary + Multiple choice question

### Direct Capture (Ctrl+B) - Creates 1 slide:
- Full-screen screenshot only (no AI)

### On Exit (Ctrl+C):
- Automatically exports presentation to PDF



## Acknowledgments

- [Google Gemini](https://deepmind.google/technologies/gemini/) - AI image analysis
- [python-pptx](https://python-pptx.readthedocs.io/) - PowerPoint creation
- [keyboard](https://github.com/boppreh/keyboard) - Global hotkey capture
- [Pillow](https://pillow.readthedocs.io/) - Image processing
