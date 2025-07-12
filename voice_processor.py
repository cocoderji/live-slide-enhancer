# voice_processor.py
# -*- coding: utf-8 -*-

"""
Part 2: Real-time Voice Processor (Whisper Medium Model)

This version uses the high-accuracy 'medium.en' Whisper model for the best
possible transcription quality.
"""

import speech_recognition as sr
import threading
import time

class VoiceProcessor(threading.Thread):
    def __init__(self, text_callback):
        """
        Initializes the VoiceProcessor for local speech recognition using Whisper.
        """
        super().__init__()
        self.daemon = True
        self.recognizer = sr.Recognizer()
        self.recognizer.dynamic_energy_threshold = True
        self.microphone = sr.Microphone()
        self.text_callback = text_callback
        self.running = True

    def run(self):
        """The main loop for the voice processing thread."""
        # --- CHANGE: Upgraded to the "medium.en" model ---
        model_name = "small.en"
        print(f"VoiceProcessor thread started. Using Whisper model: '{model_name}'...")
        print("NOTE: The first run will download the model (approx. 1.5 GB), which can take several minutes.")
        
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
        
        # Use listen_in_background for non-blocking audio capture
        stop_listening = self.recognizer.listen_in_background(self.microphone, self._audio_callback, phrase_time_limit=5)

        while self.running:
            time.sleep(0.5)

        print("Stopping voice listener...")
        stop_listening(wait_for_stop=False)
        print("VoiceProcessor thread stopped.")

    def _audio_callback(self, recognizer, audio):
        """
        Callback to transcribe audio using the Whisper 'medium.en' model and
        send the resulting text to the main application thread.
        """
        try:
            # --- CHANGE: Using the more accurate "medium.en" model ---
            text = recognizer.recognize_whisper(audio, language="english", model="medium.en")
            print(f"Transcribed: '{text}'")
            if text.strip():
                self.text_callback(text)
        except sr.UnknownValueError:
            # This is normal if there's silence or unrecognizable speech.
            pass
        except sr.RequestError as e:
            # This error is unlikely with local Whisper but good practice to keep.
            print(f"Could not request results from Whisper engine; {e}")

    def stop(self):
        """Signals the thread to stop running."""
        self.running = False
