import pandas as pd
from fastai.vision.all import *
from duckduckgo_search import DDGS
from pathlib import Path
import requests
from PIL import Image, ImageTk
from io import BytesIO
import os
import time
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
from datetime import datetime, timedelta
import asyncio
import aiohttp
import queue
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import logging
import traceback
import shutil

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('image_downloader.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# Set console to use UTF-8
sys.stdout.reconfigure(encoding='utf-8')

class ImageGalleryWindow:
    def __init__(self, parent):
        self.top = ctk.CTkToplevel(parent)
        self.top.title("Image Gallery")
        self.top.geometry("1200x800")
        
        # Create temp directory if it doesn't exist
        self.temp_dir = os.path.join("downloaded_images", "temp")
        os.makedirs(self.temp_dir, exist_ok=True)
        
        # Store search results for each SKU
        self.search_results = {}  # {sku: {'results': [], 'current_index': 0}}
        
        # Create main container with white background for better visibility
        self.main_container = ctk.CTkScrollableFrame(self.top, fg_color="white")
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Grid for images
        self.grid_frame = ctk.CTkFrame(self.main_container, fg_color="white")
        self.grid_frame.pack(fill="both", expand=True)
        
        self.image_frames = {}  # Store frames for each SKU
        self.photos = {}        # Store PhotoImage references
        self.current_replacements = {}  # Store temporary replacement images
        
        # Configure grid columns
        self.columns = 3
        for i in range(self.columns):
            self.grid_frame.grid_columnconfigure(i, weight=1)
            
    def add_image(self, sku, description, image_path):
        if sku in self.image_frames:
            return
            
        # Calculate grid position
        position = len(self.image_frames)
        row = position // self.columns
        col = position % self.columns
        
        # Create frame for this image with light gray background
        frame = ctk.CTkFrame(self.grid_frame, fg_color="#f0f0f0")
        frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
        
        # SKU and description label
        info_label = ctk.CTkLabel(frame, text=f"SKU: {sku}\n{description}", 
                                wraplength=300, text_color="black")
        info_label.pack(pady=5)
        
        # Create a frame for the image with white background
        image_container = ctk.CTkFrame(frame, fg_color="white")
        image_container.pack(pady=5, padx=5, fill="both", expand=True)
        
        # Image label with white background
        image_label = ctk.CTkLabel(image_container, text="")
        image_label.pack(pady=5, expand=True)
        
        # Buttons
        button_frame = ctk.CTkFrame(frame, fg_color="#f0f0f0")
        button_frame.pack(pady=5)
        
        replace_button = ctk.CTkButton(
            button_frame, 
            text="Replace Image", 
            command=lambda s=sku, d=description: self.get_replacement(s, d)
        )
        replace_button.pack(side="left", padx=2)
        
        approve_button = ctk.CTkButton(
            button_frame, 
            text="Approve New", 
            command=lambda s=sku: self.approve_replacement(s)
        )
        approve_button.pack(side="left", padx=2)
        approve_button.configure(state="disabled")  # Initially disabled
        
        # Store references
        self.image_frames[sku] = {
            'frame': frame,
            'image_label': image_label,
            'approve_button': approve_button,
            'description': description
        }
        
        # Load and display image
        self.load_image(sku, image_path)
        
    def load_image(self, sku, image_path):
        if os.path.exists(image_path):
            try:
                # Open and convert image
                image = Image.open(image_path)
                if image.mode in ('RGBA', 'P'):
                    image = image.convert('RGB')
                
                # Calculate size while maintaining aspect ratio
                display_size = (300, 300)
                original_size = image.size
                ratio = min(display_size[0]/original_size[0], display_size[1]/original_size[1])
                new_size = tuple(int(dim * ratio) for dim in original_size)
                
                # Resize image
                image = image.resize(new_size, Image.Resampling.LANCZOS)
                
                # Convert to PhotoImage
                photo = ImageTk.PhotoImage(image)
                
                # Store reference and update label
                self.photos[sku] = photo
                self.image_frames[sku]['image_label'].configure(image=photo)
            except Exception as e:
                logging.error(f"Error loading image for SKU {sku}: {str(e)}")
                self.image_frames[sku]['image_label'].configure(text="Error loading image")
        else:
            self.image_frames[sku]['image_label'].configure(text="Image not found")
            
    def get_replacement(self, sku, description):
        try:
            logging.info(f"Getting replacement image for SKU {sku} with description: {description}")
            
            # Initialize or get existing search results
            if sku not in self.search_results:
                ddgs = DDGS()
                results = list(ddgs.images(description, max_results=10))  # Get more results
                if not results:
                    messagebox.showinfo("Info", "No images found for this description")
                    return
                self.search_results[sku] = {
                    'results': results,
                    'current_index': 0,
                    'used_urls': set()  # Keep track of used URLs
                }
            
            search_data = self.search_results[sku]
            results = search_data['results']
            
            # Try to find a new image we haven't used yet
            found_new = False
            start_index = search_data['current_index']
            
            for i in range(len(results)):
                index = (start_index + i) % len(results)
                url = results[index]['image']
                
                if url not in search_data['used_urls']:
                    search_data['current_index'] = (index + 1) % len(results)
                    search_data['used_urls'].add(url)
                    found_new = True
                    
                    try:
                        logging.info(f"Trying new image {index + 1} of {len(results)} for SKU {sku}")
                        response = requests.get(url, timeout=10)
                        if response.status_code == 200:
                            temp_path = os.path.join(self.temp_dir, f"temp_{sku}.jpg")
                            img = Image.open(BytesIO(response.content))
                            
                            if img.mode in ('RGBA', 'P'):
                                img = img.convert('RGB')
                            
                            img.save(temp_path, "JPEG", quality=85)
                            self.current_replacements[sku] = {
                                'path': temp_path,
                                'url': url
                            }
                            self.load_image(sku, temp_path)
                            self.image_frames[sku]['approve_button'].configure(state="normal")
                            logging.info(f"Successfully loaded replacement image for SKU {sku}")
                            break
                    except Exception as e:
                        logging.error(f"Error downloading image from {url}: {str(e)}")
                        continue
            
            # If we've used all images, get a new batch
            if not found_new:
                ddgs = DDGS()
                new_results = list(ddgs.images(description, max_results=10))
                if new_results:
                    self.search_results[sku] = {
                        'results': new_results,
                        'current_index': 0,
                        'used_urls': set([self.current_replacements[sku]['url']] if sku in self.current_replacements else set())
                    }
                    # Recursively try again with new results
                    self.get_replacement(sku, description)
                else:
                    messagebox.showinfo("Info", "No more new images found, try a different search term")
            
        except Exception as e:
            logging.error(f"Error getting replacement for SKU {sku}: {str(e)}")
            logging.error(traceback.format_exc())
            messagebox.showerror("Error", f"Failed to get replacement image: {str(e)}")
            
    def approve_replacement(self, sku):
        if sku in self.current_replacements:
            temp_info = self.current_replacements[sku]
            temp_path = temp_info['path']
            final_path = os.path.join("downloaded_images", f"{sku}.jpg")
            
            try:
                if os.path.exists(temp_path):
                    shutil.copy2(temp_path, final_path)
                    os.remove(temp_path)
                    del self.current_replacements[sku]
                    self.image_frames[sku]['approve_button'].configure(state="disabled")
                    messagebox.showinfo("Success", f"Image for SKU {sku} has been replaced")
                    logging.info(f"Successfully replaced image for SKU {sku}")
            except Exception as e:
                logging.error(f"Error approving replacement for SKU {sku}: {str(e)}")
                logging.error(traceback.format_exc())
                messagebox.showerror("Error", f"Failed to approve replacement: {str(e)}")
            
    def __del__(self):
        """Cleanup temporary files when the window is closed"""
        try:
            if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
                for file in os.listdir(self.temp_dir):
                    try:
                        os.remove(os.path.join(self.temp_dir, file))
                    except:
                        pass
                os.rmdir(self.temp_dir)
        except:
            pass

class ImageDownloaderApp:
    def __init__(self):
        logging.info("Initializing ImageDownloaderApp")
        self.window = ctk.CTk()
        self.window.title("Image Downloader")
        self.window.geometry("600x400")
        
        # Create main frame
        self.main_frame = ctk.CTkFrame(self.window)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # File selection
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(pady=10, padx=10, fill="x")
        
        self.file_label = ctk.CTkLabel(self.file_frame, text="Excel File:")
        self.file_label.pack(side="left", padx=5)
        
        self.file_path = ctk.CTkEntry(self.file_frame, width=300)
        self.file_path.pack(side="left", padx=5)
        
        self.browse_button = ctk.CTkButton(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.pack(side="left", padx=5)
        
        # Image size options
        self.size_frame = ctk.CTkFrame(self.main_frame)
        self.size_frame.pack(pady=10, padx=10, fill="x")
        
        self.size_label = ctk.CTkLabel(self.size_frame, text="Max Image Size:")
        self.size_label.pack(side="left", padx=5)
        
        self.size_var = ctk.StringVar(value="800")
        self.size_entry = ctk.CTkEntry(self.size_frame, width=100, textvariable=self.size_var)
        self.size_entry.pack(side="left", padx=5)
        
        self.size_px_label = ctk.CTkLabel(self.size_frame, text="pixels")
        self.size_px_label.pack(side="left", padx=5)
        
        # Concurrent downloads option
        self.concurrent_frame = ctk.CTkFrame(self.main_frame)
        self.concurrent_frame.pack(pady=10, padx=10, fill="x")
        
        self.concurrent_label = ctk.CTkLabel(self.concurrent_frame, text="Concurrent Downloads:")
        self.concurrent_label.pack(side="left", padx=5)
        
        self.concurrent_var = ctk.StringVar(value="3")
        self.concurrent_entry = ctk.CTkEntry(self.concurrent_frame, width=50, textvariable=self.concurrent_var)
        self.concurrent_entry.pack(side="left", padx=5)
        
        # Skip existing files option
        self.skip_var = ctk.BooleanVar(value=True)
        self.skip_checkbox = ctk.CTkCheckBox(self.main_frame, text="Skip existing files", variable=self.skip_var)
        self.skip_checkbox.pack(pady=5)
        
        # Progress frame
        self.progress_frame = ctk.CTkFrame(self.main_frame)
        self.progress_frame.pack(pady=10, padx=10, fill="x")
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        self.progress_bar.set(0)
        
        # Configure progress bar appearance
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.configure(height=20)
        
        self.status_label = ctk.CTkLabel(self.progress_frame, text="")
        self.status_label.pack(pady=5)
        
        self.stats_label = ctk.CTkLabel(self.progress_frame, text="")
        self.stats_label.pack(pady=5)
        
        # Create buttons frame
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(pady=10, padx=10, fill="x")
        
        # Control buttons frame
        self.control_buttons = ctk.CTkFrame(self.button_frame)
        self.control_buttons.pack(side="left", fill="x", expand=True)
        
        # Start button
        self.start_button = ctk.CTkButton(
            self.control_buttons,
            text="Start Download",
            command=self.start_download
        )
        self.start_button.pack(side="left", padx=5)
        
        # Stop button
        self.stop_button = ctk.CTkButton(
            self.control_buttons,
            text="Stop",
            command=self.stop_download
        )
        self.stop_button.pack(side="left", padx=5)
        self.stop_button.configure(state="disabled")
        
        # Gallery button
        self.gallery_button = ctk.CTkButton(
            self.control_buttons,
            text="Open Gallery",
            command=self.show_gallery
        )
        self.gallery_button.pack(side="left", padx=5)
        
        # Log frame
        self.log_frame = ctk.CTkFrame(self.main_frame)
        self.log_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.log_text = ctk.CTkTextbox(self.log_frame, height=150)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Initialize counters and state
        self.completed_downloads = 0
        self.skipped_downloads = 0
        self.failed_downloads = 0
        self.successful_downloads = 0
        self.total_downloads = 0
        self.is_running = False
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.gallery_window = None
        logging.info("ImageDownloaderApp initialized")
        
    def browse_file(self):
        logging.info("Browse file dialog opened")
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            logging.info(f"Selected file: {filename}")
            self.file_path.delete(0, "end")
            self.file_path.insert(0, filename)
    
    def log_message(self, message):
        logging.info(message)
        try:
            def update_log():
                self.log_text.insert("end", f"{message}\n")
                self.log_text.see("end")
            self.window.after(0, update_log)
        except Exception as e:
            logging.error(f"Error updating log: {str(e)}")
    
    def update_progress(self):
        try:
            if self.total_downloads > 0:
                progress = (self.successful_downloads + self.skipped_downloads + self.failed_downloads) / self.total_downloads
                logging.debug(f"Updating progress: {self.successful_downloads + self.skipped_downloads + self.failed_downloads}/{self.total_downloads}")
                def update():
                    try:
                        self.progress_bar.set(progress)
                        self.status_label.configure(text=f"Progress: {self.successful_downloads + self.skipped_downloads + self.failed_downloads}/{self.total_downloads}")
                        stats_text = f"Completed: {self.successful_downloads} | "
                        stats_text += f"Skipped: {self.skipped_downloads} | "
                        stats_text += f"Failed: {self.failed_downloads}"
                        self.stats_label.configure(text=stats_text)
                        logging.debug(f"Stats - Success: {self.successful_downloads}, "
                                    f"Skipped: {self.skipped_downloads}, "
                                    f"Failed: {self.failed_downloads}, "
                                    f"Total Progress: {self.successful_downloads + self.skipped_downloads + self.failed_downloads}/{self.total_downloads}")
                    except Exception as e:
                        logging.error(f"Error in progress update: {str(e)}")
                self.window.after(0, update)
        except Exception as e:
            logging.error(f"Error in update_progress: {str(e)}")

    def increment_counter(self, counter_name):
        def _increment():
            if counter_name == 'successful':
                self.successful_downloads += 1
            elif counter_name == 'skipped':
                self.skipped_downloads += 1
            elif counter_name == 'failed':
                self.failed_downloads += 1
            self.update_progress()
        self.window.after(0, _increment)

    def download_image(self, url, sku, output_dir, max_size, description):
        logging.info(f"Starting download for SKU: {sku}")
        output_path = os.path.join(output_dir, f"{sku}.jpg")
        
        # Skip if file exists and skip option is enabled
        if self.skip_var.get() and os.path.exists(output_path):
            logging.info(f"Skipping existing image for SKU: {sku}")
            self.log_message(f"Skipping existing image for מק\"ט: {sku}")
            self.increment_counter('skipped')
            return True
            
        try:
            logging.debug(f"Downloading image from URL: {url}")
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                logging.debug(f"Download successful for SKU: {sku}")
                img = Image.open(BytesIO(response.content))
                
                logging.debug(f"Original image mode: {img.mode}")
                # Convert to RGB if necessary
                if img.mode in ('RGBA', 'P'):
                    logging.debug(f"Converting {img.mode} to RGB")
                    img = img.convert('RGB')
                elif img.mode != 'RGB':
                    logging.debug(f"Converting {img.mode} to RGB")
                    img = img.convert('RGB')
                
                # Resize image while maintaining aspect ratio
                if max(img.size) > max_size:
                    ratio = max_size / max(img.size)
                    new_size = tuple(int(dim * ratio) for dim in img.size)
                    logging.debug(f"Resizing image from {img.size} to {new_size}")
                    img = img.resize(new_size, Image.Resampling.LANCZOS)
                
                logging.debug(f"Saving image to: {output_path}")
                img.save(output_path, "JPEG", quality=85, optimize=True)
                self.log_message(f"Successfully downloaded image for מק\"ט: {sku}")
                self.increment_counter('successful')
                if self.gallery_window:
                    self.window.after(0, lambda: self.gallery_window.add_image(sku, description, output_path))
                return True
        except Exception as e:
            logging.error(f"Error downloading image for SKU {sku}: {str(e)}")
            logging.error(traceback.format_exc())
            self.log_message(f"Error downloading image for מק\"ט {sku}: {str(e)}")
            self.increment_counter('failed')
        return False
    
    def process_item(self, row, output_dir, max_size):
        if not self.is_running:
            logging.info("Process stopped by user")
            return
            
        try:
            sku = str(row['מק"ט'])
            description = str(row['תאור'])
            
            logging.info(f"Processing SKU: {sku}")
            logging.info(f"Description: {description}")
            
            self.log_message(f"Processing מק\"ט: {sku}")
            self.log_message(f"תאור: {description}")
            
            try:
                ddgs = DDGS()
                logging.debug(f"Searching for images with description: {description}")
                results = list(ddgs.images(description, max_results=1))
                if results:
                    logging.debug(f"Found image for SKU {sku}")
                    self.download_image(results[0]['image'], sku, output_dir, max_size, description)
                else:
                    logging.warning(f"No images found for SKU: {sku}")
                    self.log_message(f"No images found for מק\"ט: {sku}")
                    self.increment_counter('failed')
            except Exception as e:
                logging.error(f"Error searching for SKU {sku}: {str(e)}")
                logging.error(traceback.format_exc())
                self.log_message(f"Error searching for מק\"ט {sku}: {str(e)}")
                self.increment_counter('failed')
            
            time.sleep(0.5)  # Small delay between requests
            
        except Exception as e:
            logging.error(f"Error in process_item: {str(e)}")
            logging.error(traceback.format_exc())
            self.increment_counter('failed')
    
    def download_process(self, excel_path, max_size, concurrent_limit):
        try:
            logging.info("Starting download process")
            output_dir = "downloaded_images"
            os.makedirs(output_dir, exist_ok=True)
            
            logging.info(f"Reading Excel file: {excel_path}")
            df = pd.read_excel(excel_path)
            self.total_downloads = len(df)
            self.completed_downloads = 0
            self.skipped_downloads = 0
            self.failed_downloads = 0
            self.successful_downloads = 0
            logging.info(f"Total items to process: {self.total_downloads}")
            self.update_progress()
            
            with ThreadPoolExecutor(max_workers=concurrent_limit) as executor:
                futures = []
                for index, row in df.iterrows():
                    if not self.is_running:
                        logging.info("Download process stopped by user")
                        break
                    logging.debug(f"Submitting task for row {index + 1}/{self.total_downloads}")
                    future = executor.submit(self.process_item, row, output_dir, max_size)
                    futures.append(future)
                
                logging.info("Waiting for all tasks to complete")
                concurrent.futures.wait(futures)
            
            if self.is_running:
                logging.info("Download process completed successfully")
                self.log_message("Download process completed!")
                self.window.after(0, self.status_label.configure, {"text": "Download completed!"})
                
        except Exception as e:
            logging.error(f"Error in download process: {str(e)}")
            logging.error(traceback.format_exc())
            self.log_message(f"Error: {str(e)}")
        finally:
            logging.info("Download process finished")
            self.is_running = False
            self.window.after(0, self.start_button.configure, {"state": "normal"})
    
    def start_download(self):
        try:
            excel_path = self.file_path.get()
            if not excel_path:
                logging.warning("No Excel file selected")
                self.log_message("Please select an Excel file first!")
                return
            
            logging.info(f"Starting download with Excel file: {excel_path}")
            max_size = int(self.size_var.get())
            concurrent_limit = int(self.concurrent_var.get())
            logging.info(f"Parameters - Max size: {max_size}, Concurrent limit: {concurrent_limit}")
            
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.is_running = True
            
            # Start download in a separate thread
            self.executor.submit(self.download_process, excel_path, max_size, concurrent_limit)
            
        except Exception as e:
            logging.error(f"Error starting download: {str(e)}")
            logging.error(traceback.format_exc())
            self.log_message(f"Error starting download: {str(e)}")
    
    def stop_download(self):
        """Stop the current download process"""
        self.is_running = False
        self.log_message("Download process stopped by user")
        logging.info("Download process stopped by user")
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
    
    def show_gallery(self):
        if not self.gallery_window:
            self.gallery_window = ImageGalleryWindow(self.window)
            self.gallery_window.top.protocol("WM_DELETE_WINDOW", self.close_gallery)
            
        # Load all existing images
        output_dir = "downloaded_images"
        if os.path.exists(output_dir):
            for filename in os.listdir(output_dir):
                if filename.endswith('.jpg'):
                    sku = filename[:-4]  # Remove .jpg extension
                    image_path = os.path.join(output_dir, filename)
                    # Get description from excel data if available
                    description = next((str(row['תאור']) for _, row in pd.read_excel(self.file_path.get()).iterrows() 
                                     if str(row['מק"ט']) == sku), "No description")
                    self.gallery_window.add_image(sku, description, image_path)
    
    def close_gallery(self):
        if self.gallery_window:
            self.gallery_window.top.destroy()
            self.gallery_window = None
    
    def run(self):
        logging.info("Starting application")
        self.window.mainloop()

if __name__ == "__main__":
    try:
        app = ImageDownloaderApp()
        app.run()
    except Exception as e:
        logging.error(f"Application error: {str(e)}")
        logging.error(traceback.format_exc())
