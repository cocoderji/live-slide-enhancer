# main.py
# -*- coding: utf-8 -*-

"""
Part 5: Main Application Logic (No Approval Mode)

This version removes the Approval Mode UI and logic for a fully
automatic workflow.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import configparser
import threading
import queue
import os
from collections import deque

# Import custom modules
from theme_analyzer import ThemeAnalyzer
from voice_processor import VoiceProcessor
from content_generator import ContentGenerator
from slide_updater import SlideUpdater
from visual_generator import VisualGenerator

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Savi.ai - Live Slide Enhancer")
        self.geometry("450x450") # Adjusted size
        self.configure(bg="#FFFFFF")

        # --- UI Styling ---
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure("TButton", background="#4285F4", foreground="white", bordercolor="#FFFFFF", lightcolor="#4285F4", darkcolor="#4285F4", padding=10, font=('Segoe UI', 10, 'bold'))
        self.style.map("TButton", background=[('active', '#3367D6')], foreground=[('disabled', '#9E9E9E')])
        self.style.configure("TFrame", background="#FFFFFF")
        self.style.configure("File.TLabel", background="#F5F5F5", foreground="#212121", padding=10, anchor='center', font=('Segoe UI', 9), borderwidth=1, relief="solid", bordercolor="#E0E0E0")
        self.style.configure("Status.TLabel", background="#FFFFFF", foreground="#424242", font=('Segoe UI', 10))
        self.style.configure("Header.TLabel", background="#FFFFFF", foreground="#212121", font=('Segoe UI', 11, 'bold'))

        # --- Application State ---
        self.pptx_path = None
        self.style_guide = None
        self.is_running = False
        self.is_checking = False
        self.voice_thread = None
        self.speech_queue = queue.Queue()
        self.speech_buffer = deque(maxlen=10)
        self.periodic_check_id = None
        self.pulse_id = None
        self.generated_slide_indices = set()

        # --- Load APIs & Instantiate Core Components ---
        self.api_key, self.pexels_key, self.noun_key, self.noun_secret = self._load_api_keys()
        if not self.api_key:
            self.destroy(); return
        self.content_generator = ContentGenerator(api_key=self.api_key)
        self.slide_updater = None
        self.visual_generator = VisualGenerator(self.pexels_key, self.noun_key, self.noun_secret)
        
        # --- Create UI ---
        self._create_widgets()

        # --- Final Setup ---
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.after(200, self.process_speech_queue)

    def _create_widgets(self):
        """Creates and grids all the UI elements."""
        self.grid_columnconfigure(0, weight=1)
        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        # --- Top Controls ---
        self.file_label = ttk.Label(main_frame, text="No presentation selected.", style="File.TLabel", wraplength=400)
        self.file_label.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.select_button = ttk.Button(main_frame, text="Select PowerPoint & Start Show", command=self.select_file_and_start_show)
        self.select_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")
        
        self.start_button = ttk.Button(main_frame, text="Start Auto-Enhancer", command=self.start_processing, state=tk.DISABLED)
        self.start_button.grid(row=2, column=0, pady=(10, 5), sticky="ew")
        self.stop_button = ttk.Button(main_frame, text="Stop Auto-Enhancer", command=self.stop_processing, state=tk.DISABLED)
        self.stop_button.grid(row=2, column=1, pady=(10, 5), sticky="ew", padx=(10, 0))

        # --- Manual Input Section ---
        ttk.Separator(main_frame, orient='horizontal').grid(row=3, column=0, columnspan=2, pady=15, sticky="ew")
        ttk.Label(main_frame, text="Manual Slide Generation", style="Header.TLabel").grid(row=4, column=0, columnspan=2, sticky="w")
        self.manual_topic_entry = ttk.Entry(main_frame, font=('Segoe UI', 10))
        self.manual_topic_entry.grid(row=5, column=0, columnspan=2, pady=5, sticky="ew")
        self.manual_generate_button = ttk.Button(main_frame, text="Generate Slide Manually", command=self.handle_manual_generate, state=tk.DISABLED)
        self.manual_generate_button.grid(row=6, column=0, columnspan=2, pady=5, sticky="ew")

        # --- Status Indicator ---
        status_frame = ttk.Frame(main_frame, style="TFrame")
        status_frame.grid(row=7, column=0, columnspan=2, pady=(20, 0), sticky="ew")
        status_frame.grid_columnconfigure(1, weight=1)
        self.status_icon = tk.Canvas(status_frame, width=12, height=12, bg="#FFFFFF", highlightthickness=0)
        self.status_dot = self.status_icon.create_oval(2, 2, 12, 12, fill="#BDBDBD", outline="")
        self.status_icon.grid(row=0, column=0, padx=(0, 10))
        self.status_label = ttk.Label(status_frame, text="Status: Idle", style="Status.TLabel")
        self.status_label.grid(row=0, column=1, sticky="w")

    def _set_status(self, text, color, pulse=False):
        self.after(0, self._update_status_widgets, text, color, pulse)

    def _update_status_widgets(self, text, color, pulse):
        self.status_label.config(text=f"Status: {text}")
        self.status_icon.itemconfig(self.status_dot, fill=color)
        if self.pulse_id: self.after_cancel(self.pulse_id); self.pulse_id = None
        if pulse: self._pulse_animation(True)

    def _pulse_animation(self, to_bright):
        color = "#4285F4" if to_bright else "#8AB4F8"
        self.status_icon.itemconfig(self.status_dot, fill=color)
        self.pulse_id = self.after(700, self._pulse_animation, not to_bright)

    def _load_api_keys(self):
        config = configparser.ConfigParser()
        if not os.path.exists('config.ini'):
            messagebox.showerror("Error", "config.ini file not found!"); return None, None, None, None
        config.read('config.ini')
        try:
            gemini_key = config['API_KEYS']['GEMINI_API_KEY']
            pexels_key = config['API_KEYS'].get('PEXELS_API_KEY')
            noun_key = config['API_KEYS'].get('NOUN_PROJECT_API_KEY')
            noun_secret = config['API_KEYS'].get('NOUN_PROJECT_SECRET_KEY')
            if not gemini_key or 'YOUR_GEMINI_API_KEY' in gemini_key:
                messagebox.showerror("Error", "Please set your Gemini API key in config.ini"); return None, None, None, None
            return gemini_key, pexels_key, noun_key, noun_secret
        except KeyError:
            messagebox.showerror("Error", "API_KEYS section not found in config.ini"); return None, None, None, None

    def select_file_and_start_show(self):
        path = filedialog.askopenfilename(title="Select a PowerPoint Presentation", filetypes=(("PowerPoint files", "*.pptx"), ("All files", "*.*")))
        if path:
            self.pptx_path = path
            self.file_label.config(text=f"Selected: {os.path.basename(self.pptx_path)}")
            self.generated_slide_indices.clear()
            try:
                analyzer = ThemeAnalyzer(self.pptx_path)
                self.style_guide = analyzer.get_style()
                if not self.style_guide:
                    messagebox.showerror("Error", "Could not analyze the presentation theme."); return
                self.slide_updater = SlideUpdater(self.pptx_path)
                if self.slide_updater.start_presentation_show():
                    self.start_button.config(state=tk.NORMAL)
                    self.manual_generate_button.config(state=tk.NORMAL)
                    self._set_status("Slideshow running. Ready.", "#34A853")
                else:
                    messagebox.showerror("Error", "Failed to start the PowerPoint slideshow.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open or analyze presentation: {e}")

    def start_processing(self):
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self._set_status("Listening for speech...", "#4285F4", pulse=True)
        self.periodic_check()
        self.voice_thread = VoiceProcessor(text_callback=lambda text: self.speech_queue.put(text))
        self.voice_thread.start()

    def stop_processing(self):
        if self.periodic_check_id: self.after_cancel(self.periodic_check_id); self.periodic_check_id = None
        if self.voice_thread and self.voice_thread.is_alive(): self.voice_thread.stop(); self.voice_thread.join()
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self._set_status("Stopped. Ready to start again.", "#34A853")

    def process_speech_queue(self):
        try:
            while not self.speech_queue.empty(): self.speech_buffer.append(self.speech_queue.get_nowait())
        except queue.Empty: pass
        finally: self.after(200, self.process_speech_queue)

    def periodic_check(self):
        if self.is_running and not self.is_checking:
            self.is_checking = True
            check_thread = threading.Thread(target=self.run_deviation_check)
            check_thread.start()
        if self.is_running:
            self.periodic_check_id = self.after(8000, self.periodic_check)

    def run_deviation_check(self):
        try:
            if not self.slide_updater or len(self.speech_buffer) < 3: return
            current_index = self.slide_updater.get_current_slide_index()
            if not current_index: return
            slide_text = self.slide_updater.get_text_from_slide(current_index)
            if not slide_text.strip(): slide_text = "An empty slide."
            recent_speech = " ".join(self.speech_buffer)
            self._set_status("Analyzing speech vs. slide...", "#FBBC05")
            new_topic = self.content_generator.check_for_deviation(slide_text, recent_speech)
            if new_topic:
                self.speech_buffer.clear()
                is_update_action = current_index in self.generated_slide_indices
                # Directly handle the update instead of previewing
                self.handle_update(new_topic, current_index, is_update_action)
            else:
                self._set_status("Listening for speech...", "#4285F4", pulse=True)
        finally:
            self.is_checking = False

    def handle_manual_generate(self):
        """Handles the manual generate button click."""
        topic = self.manual_topic_entry.get()
        if not topic.strip():
            messagebox.showwarning("Input Error", "Please enter a topic."); return
        if not self.slide_updater:
            messagebox.showwarning("Error", "Please select a presentation first."); return
        current_index = self.slide_updater.get_current_slide_index()
        if not current_index:
            messagebox.showwarning("Error", "Could not find active slideshow."); return
        # Manually generated slides are always inserted
        self.handle_update(topic, current_index, is_update=False)

    def handle_update(self, topic, slide_index, is_update):
        """Generates content and immediately applies it to the presentation."""
        action_text = "Updating" if is_update else "Inserting"
        self._set_status(f"New topic '{topic}'. Generating content...", "#FBBC05")
        
        # Run content generation in a separate thread to keep UI responsive
        threading.Thread(target=self._generate_and_apply, args=(topic, slide_index, is_update, action_text)).start()

    def _generate_and_apply(self, topic, slide_index, is_update, action_text):
        """Helper function to run the blocking tasks in a background thread."""
        content = self.content_generator.generate_slide_content(topic, self.style_guide)
        if content:
            self._set_status(f"{action_text} slide...", "#FBBC05")
            if is_update:
                self.slide_updater.update_existing_slide(slide_index, content, self.style_guide, self.visual_generator)
            else:
                new_slide_index = slide_index + 1
                self.slide_updater.insert_new_slide_after_current(slide_index, content, self.style_guide, self.visual_generator)
                self.generated_slide_indices.add(new_slide_index)
                print(f"INFO: New slide at index {new_slide_index} is now being tracked.")
            self._set_status("Listening for speech...", "#4285F4", pulse=True)
        else:
            self._set_status("Failed to generate content. Listening...", "#EA4335")

    def on_closing(self):
        if self.is_running: self.stop_processing()
        self.destroy()

if __name__ == "__main__":
    app = Application()
    app.mainloop()
