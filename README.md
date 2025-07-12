# Live Slide Enhancer

A Python application that uses AI to automatically generate and update PowerPoint slides in real-time based on the presenter's speech or given topic

## Features

* **Real-time Transcription:** Uses Whisper to transcribe the presenter's speech.
* **Deviation Detection:** Determines when the presenter has moved to a new topic.
* **Automatic Slide Generation:** Generates new slides with relevant content and visuals.
* **Theme Analysis:** Analyzes the existing presentation to match the new slides' style.
* **Data-Driven Visuals:** Creates charts and finds icons/images for the generated slides.

## How It Works

The application listens to the presenter's speech and analyzes it against the content of the current slide. If it detects a significant deviation in the topic, it uses a generative AI model to create new slide content, including a title, bullet points, and visuals. It then automatically inserts the new slide into the live presentation.

## Getting Started

### Prerequisites

* Python 3.13.2+
* Microsoft PowerPoint

### Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/cocoderji/live-slide-enhancer.git](https://github.com/cocoderji/live-slide-enhancer.git)
    cd live-slide-enhancer
    ```
2.  **Install the required libraries:**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Configure your API keys:**
    * Rename `config.ini.example` to `config.ini`.
    * Add your API keys for Gemini, Pexels, and The Noun Project to `config.ini`.
2.  **Run the application:**
    ```bash
    python main.py
    ```
3.  **Select your PowerPoint file** and the application will start the slideshow and begin listening.

## Configuration

The `config.ini` file is used to store your API keys. You will need to get keys from:

* **Google AI for Developers (Gemini)**
* **Pexels (for images)**
* **The Noun Project (for icons)**

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
