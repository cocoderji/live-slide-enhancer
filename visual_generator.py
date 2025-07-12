# visual_generator.py
# -*- coding: utf-8 -*-

"""
Visual Generator (Advanced Charting)

This module contains an upgraded chart generation engine capable of
handling both simple and multi-dataset bar charts.
"""

import os
import requests
from thenounproject.api import Api
import matplotlib.pyplot as plt

class VisualGenerator:
    def __init__(self, pexels_key, noun_key, noun_secret):
        self.pexels_key = pexels_key
        self.noun_project = Api(noun_key, noun_secret)
        self.temp_dir = "temp_visuals"
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)

    def _save_temp_file(self, data, extension):
        filename = f"temp_{int(plt.fignum_exists(1))}.{extension}"
        path = os.path.join(self.temp_dir, filename)
        with open(path, 'wb') as f:
            f.write(data)
        return path

    def get_image(self, query):
        if not self.pexels_key or "YOUR_PEXELS_API_KEY" in self.pexels_key: return None
        headers = {"Authorization": self.pexels_key}
        url = f"https://api.pexels.com/v1/search?query={query}&per_page=1"
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()
            if data['photos']:
                image_url = data['photos'][0]['src']['large']
                image_data = requests.get(image_url).content
                return self._save_temp_file(image_data, 'jpg')
        except Exception as e:
            print(f"ERROR: Failed to fetch image from Pexels: {e}")
        return None

    def get_icon(self, query):
        if not self.noun_project.api_key or "YOUR_NOUN_PROJECT" in self.noun_project.api_key: return None
        try:
            icons = self.noun_project.icon.list(query, limit=1)
            if icons:
                icon_url = icons[0].preview_url
                icon_data = requests.get(icon_url).content
                return self._save_temp_file(icon_data, 'png')
        except Exception as e:
            print(f"ERROR: Failed to fetch icon from Noun Project: {e}")
        return None

    def create_chart(self, chart_data):
        """Creates a bar or pie chart image from potentially complex data."""
        print(f"INFO: Generating chart with data: {chart_data}")
        try:
            labels = chart_data.get('labels', [])
            chart_type = chart_data.get('type', 'bar')

            plt.style.use('seaborn-v0_8-whitegrid')
            fig, ax = plt.subplots(figsize=(6, 4))
            
            if chart_type == 'pie':
                values = chart_data.get('values', [])
                ax.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
                ax.axis('equal')
            else: # Handle bar charts (simple or multi-dataset)
                datasets = chart_data.get('datasets')
                if datasets: # Multi-series bar chart
                    for dataset in datasets:
                        ax.bar(labels, dataset.get('data', []), label=dataset.get('label'))
                    ax.legend()
                else: # Simple bar chart
                    values = chart_data.get('values', [])
                    ax.bar(labels, values, color='#4285F4')

            ax.set_title(chart_data.get('title', 'Chart'), fontsize=14, weight='bold')
            ax.tick_params(axis='x', rotation=45)
            plt.tight_layout()
            
            chart_path = os.path.join(self.temp_dir, "temp_chart.png")
            plt.savefig(chart_path)
            plt.close(fig)
            print(f"SUCCESS: Saved chart to {chart_path}")
            return chart_path
        except Exception as e:
            print(f"ERROR: Failed to create chart: {e}")
        return None

    def cleanup_temp_files(self):
        """Removes all files from the temporary visuals directory."""
        print("INFO: Cleaning up temporary visual files...")
        for filename in os.listdir(self.temp_dir):
            file_path = os.path.join(self.temp_dir, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"ERROR: Failed to delete temp file {file_path}: {e}")
