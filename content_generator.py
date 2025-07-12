# content_generator.py
# -*- coding: utf-8 -*-

import google.generativeai as genai
import json

class ContentGenerator:
    def __init__(self, api_key):
        self.api_key = api_key
        genai.configure(api_key=self.api_key)
        self.model = genai.GenerativeModel('gemini-1.5-flash-latest')

    def _extract_json(self, text):
        """A robust function to extract a JSON object from a string."""
        try:
            start_index = text.index('{')
            end_index = text.rindex('}') + 1
            json_str = text[start_index:end_index]
            return json.loads(json_str)
        except (ValueError, json.JSONDecodeError) as e:
            print(f"ERROR: Could not extract or parse JSON from response. Details: {e}")
            return None

    def check_for_deviation(self, slide_text, speech_text):
        """Checks if the user's speech has deviated from the slide's content."""
        prompt = f"""
        Analyze the following. Has the presenter's speech moved to a new topic not covered by the slide?
        Current Slide Content: "{slide_text}"
        Presenter's Speech: "{speech_text}"
        Task: If the speech has clearly shifted to a new topic, respond ONLY with the new topic in up to 10 words. Otherwise, respond with the single word: None.
        Response:
        """
        try:
            response = self.model.generate_content(prompt)
            new_topic = response.text.strip().splitlines()[-1].strip()
            if new_topic.lower() != "none" and len(new_topic) > 0:
                print(f"--- Deviation Detected. New Topic: {new_topic} ---")
                return new_topic
            return None
        except Exception:
            return None

    def generate_slide_content(self, topic, style_guide):
        """Generates engaging, data-driven slide content."""
        prompt = f"""
        You are a factual research assistant creating a presentation slide. Your goal is to be informative and data-driven.
        The presentation topic is: "{topic}".

        **CRITICAL INSTRUCTIONS:**
        1.  **Title:** Create a clear, factual title directly related to the topic.
        2.  **Bullet Points:** Write 4-5 bullet points. You MUST find and include real, quantifiable data (e.g., financial figures, percentages, statistics) relevant to the topic. For a topic like "Virat Kohli's wealth," you must include his estimated net worth and earnings. For "Google's quarterly earnings," you must include revenue and net income figures.
        3.  **Layout Suggestion:**
            - If you generate `chart_data`, the layout MUST be 'text_left_visual_right'.
            - If you do not generate `chart_data`, the layout can be 'text_only' or 'text_left_visual_right'.
        4.  **Chart Data (Conditional):**
            - If the topic is about finances, statistics, market share, or any quantifiable data, you MUST generate a simple JSON object for a chart summarizing the key data points from your bullet points. Use a 'pie' chart for breakdowns (like earnings sources) or a 'bar' chart for comparisons. For a bar chart, use a simple "values" array, not a complex "datasets" structure.
            - If the topic is purely qualitative (e.g., "The History of Origami"), this value MUST be null.
        5.  **Image Suggestion (Conditional):**
            - If you generated `chart_data`, this value MUST be null.
            - If `chart_data` is null, suggest a simple, direct search query for a relevant icon or image.

        Format the entire response as a single JSON object with the keys: "title", "points", "image_suggestion", "layout", and "chart_data".
        """
        print(f"Generating data-driven content for topic: {topic}...")
        try:
            response = self.model.generate_content(prompt)
            print(f"DEBUG (generate_slide_content): Raw response from Gemini: '{response.text}'")
            
            content_json = self._extract_json(response.text)
            
            if content_json:
                print("Engaging content generated successfully.")
                return content_json
            else:
                return None

        except Exception as e:
            print(f"Error generating slide content with Gemini API: {e}")
            return None
