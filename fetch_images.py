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
import logging
import traceback
import shutil
from tkinter import messagebox, filedialog
from concurrent.futures import ThreadPoolExecutor
import threading
from datetime import datetime, timedelta
import asyncio
import aiohttp
import queue
import concurrent.futures
import logging
import traceback
import shutil

def setup_logging():
    # Remove all existing handlers
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
        if hasattr(handler, 'close'):
            handler.close()
            
    # Create logs directory if it doesn't exist
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    
    # Create a new log file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"image_downloader_{timestamp}.log")
    
    # Configure logging
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()  # Also log to console
        ]
    )
    logging.info("Logging initialized")
    return log_file

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
        try:
            # Create frame for this image if it doesn't exist
            if sku not in self.image_frames:
                row = len(self.image_frames) // self.columns
                col = len(self.image_frames) % self.columns
                
                frame = ctk.CTkFrame(self.grid_frame)
                frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                
                # Load and resize image
                img = Image.open(image_path)
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                
                # Calculate size while maintaining aspect ratio
                display_size = (200, 200)
                original_size = img.size
                ratio = min(display_size[0]/original_size[0], display_size[1]/original_size[1])
                new_size = tuple(int(dim * ratio) for dim in original_size)
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                
                # Convert to CTkImage
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=new_size)
                
                # Create and pack widgets
                image_label = ctk.CTkLabel(frame, image=ctk_img, text="")
                image_label.pack(pady=5)
                
                sku_label = ctk.CTkLabel(frame, text=f"SKU: {sku}")
                sku_label.pack(pady=2)
                
                desc_label = ctk.CTkLabel(frame, text=description, wraplength=180)
                desc_label.pack(pady=2)
                
                replace_button = ctk.CTkButton(
                    frame,
                    text="Replace Image",
                    command=lambda s=sku, d=description: self.get_replacement(s, d)
                )
                replace_button.pack(pady=5)
                
                approve_button = ctk.CTkButton(
                    frame,
                    text="Approve New",
                    command=lambda s=sku: self.approve_replacement(s),
                    state="disabled"
                )
                approve_button.pack(pady=5)
                
                # Store references
                self.image_frames[sku] = {
                    'frame': frame,
                    'image_label': image_label,
                    'ctk_image': ctk_img,
                    'approve_button': approve_button
                }
                
        except Exception as e:
            logging.error(f"Error adding image for SKU {sku}: {str(e)}")
            logging.error(traceback.format_exc())

    def load_image(self, sku, image_path):
        try:
            if sku in self.image_frames:
                img = Image.open(image_path)
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                
                # Calculate size while maintaining aspect ratio
                display_size = (200, 200)
                original_size = img.size
                ratio = min(display_size[0]/original_size[0], display_size[1]/original_size[1])
                new_size = tuple(int(dim * ratio) for dim in original_size)
                
                # Resize image
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                img.save(image_path, "JPEG", quality=85)
                
                # Update CTkImage
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=new_size)
                self.image_frames[sku]['ctk_image'] = ctk_img
                self.image_frames[sku]['image_label'].configure(image=ctk_img)
                
        except Exception as e:
            logging.error(f"Error loading image for SKU {sku}: {str(e)}")
            logging.error(traceback.format_exc())
            
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

class SingleImageWindow:
    def __init__(self, parent):
        self.top = ctk.CTkToplevel(parent)
        self.top.title("Single Image Download")
        self.top.geometry("600x800")
        
        # Create temp directory if it doesn't exist
        self.temp_dir = os.path.join("downloaded_images", "temp")
        os.makedirs(self.temp_dir, exist_ok=True)
        
        # Store search results
        self.search_results = []
        self.current_index = 0
        self.used_urls = set()
        self.current_image_path = None
        self.photo = None  # Keep reference to prevent garbage collection
        
        # Create main container
        self.main_frame = ctk.CTkFrame(self.top)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Description entry
        self.desc_frame = ctk.CTkFrame(self.main_frame)
        self.desc_frame.pack(fill="x", pady=5)
        self.desc_label = ctk.CTkLabel(self.desc_frame, text="Description:")
        self.desc_label.pack(side="left", padx=5)
        self.desc_entry = ctk.CTkEntry(self.desc_frame, width=300)
        self.desc_entry.pack(side="left", padx=5)
        
        # Filename entry
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(fill="x", pady=5)
        self.file_label = ctk.CTkLabel(self.file_frame, text="Save as:")
        self.file_label.pack(side="left", padx=5)
        self.file_entry = ctk.CTkEntry(self.file_frame, width=300)
        self.file_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Image preview frame with white background
        self.preview_frame = ctk.CTkFrame(self.main_frame, fg_color="white")
        self.preview_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Image label
        self.image_label = ctk.CTkLabel(self.preview_frame, text="No image loaded")
        self.image_label.pack(pady=10, expand=True)
        
        # Buttons frame
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(fill="x", pady=5)
        
        # Search button
        self.search_button = ctk.CTkButton(
            self.button_frame,
            text="Search Image",
            command=self.search_image
        )
        self.search_button.pack(side="left", padx=5)
        
        # Next Image button
        self.next_button = ctk.CTkButton(
            self.button_frame,
            text="Next Image",
            command=self.next_image,
            state="disabled"
        )
        self.next_button.pack(side="left", padx=5)
        
        # Save button
        self.save_button = ctk.CTkButton(
            self.button_frame,
            text="Save Image",
            command=self.save_image,
            state="disabled"
        )
        self.save_button.pack(side="left", padx=5)
        
    def search_image(self):
        description = self.desc_entry.get().strip()
        if not description:
            messagebox.showwarning("Warning", "Please enter a description")
            return
            
        try:
            ddgs = DDGS()
            self.search_results = list(ddgs.images(description, max_results=10))
            self.current_index = 0
            self.used_urls = set()
            
            if self.search_results:
                self.show_current_image()
                self.next_button.configure(state="normal")
            else:
                messagebox.showinfo("Info", "No images found for this description")
                
        except Exception as e:
            logging.error(f"Error searching for image: {str(e)}")
            messagebox.showerror("Error", f"Failed to search for image: {str(e)}")
    
    def show_current_image(self):
        if not self.search_results:
            return
            
        try:
            url = self.search_results[self.current_index]['image']
            self.used_urls.add(url)
            
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                temp_path = os.path.join(self.temp_dir, "temp_preview.jpg")
                img = Image.open(BytesIO(response.content))
                
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                
                # Calculate size while maintaining aspect ratio
                display_size = (400, 400)
                original_size = img.size
                ratio = min(display_size[0]/original_size[0], display_size[1]/original_size[1])
                new_size = tuple(int(dim * ratio) for dim in original_size)
                
                # Resize image
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                img.save(temp_path, "JPEG", quality=85)
                
                # Update preview
                self.photo = ImageTk.PhotoImage(img)
                self.image_label.configure(image=self.photo)
                self.current_image_path = temp_path
                self.save_button.configure(state="normal")
                
        except Exception as e:
            logging.error(f"Error showing image: {str(e)}")
            self.next_image()  # Try next image if current fails
    
    def next_image(self):
        if not self.search_results:
            return
            
        self.current_index = (self.current_index + 1) % len(self.search_results)
        # If we've seen all images, get new ones
        if len(self.used_urls) >= len(self.search_results):
            self.search_image()
        else:
            self.show_current_image()
    
    def save_image(self):
        if not self.current_image_path:
            return
            
        filename = self.file_entry.get().strip()
        if not filename:
            messagebox.showwarning("Warning", "Please enter a filename")
            return
            
        if not filename.endswith('.jpg'):
            filename += '.jpg'
            
        try:
            output_path = os.path.join("downloaded_images", filename)
            shutil.copy2(self.current_image_path, output_path)
            messagebox.showinfo("Success", f"Image saved as {filename}")
            self.top.destroy()
        except Exception as e:
            logging.error(f"Error saving image: {str(e)}")
            messagebox.showerror("Error", f"Failed to save image: {str(e)}")

class ImageDownloaderApp:
    def __init__(self):
        logging.info("Initializing ImageDownloaderApp")
        self.window = ctk.CTk()
        self.window.title("Image Downloader")
        self.window.geometry("800x600")  # Increased window size
        
        # Initialize variables
        self.skip_var = ctk.BooleanVar(value=True)
        self.max_size_var = ctk.StringVar(value="800")
        self.concurrent_var = ctk.StringVar(value="3")
        self.file_path = ctk.StringVar()
        self.completed_downloads = 0
        self.failed_downloads = 0
        self.skipped_downloads = 0
        self.total_downloads = 0
        self.is_running = False
        self.gallery_window = None
        
        # Create main frame with padding
        self.main_frame = ctk.CTkFrame(self.window)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Add Single Image Download button at the top
        self.single_image_button = ctk.CTkButton(
            self.main_frame,
            text="Single Image Download",
            command=self.open_single_image_window
        )
        self.single_image_button.pack(pady=(0, 10))
        
        # Add separator
        self.separator = ctk.CTkFrame(self.main_frame, height=2, fg_color="gray75")
        self.separator.pack(fill="x", pady=10)
        
        # File selection frame
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.pack(fill="x", pady=(0, 10))
        
        self.file_label = ctk.CTkLabel(self.file_frame, text="Excel File:")
        self.file_label.pack(side="left", padx=5)
        
        self.file_entry = ctk.CTkEntry(self.file_frame, textvariable=self.file_path, width=400)
        self.file_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        self.browse_button = ctk.CTkButton(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.pack(side="left", padx=5)
        
        # Settings frame
        self.settings_frame = ctk.CTkFrame(self.main_frame)
        self.settings_frame.pack(fill="x", pady=(0, 10))
        
        # Image size setting
        self.size_label = ctk.CTkLabel(self.settings_frame, text="Max Image Size:")
        self.size_label.pack(side="left", padx=5)
        
        self.size_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.max_size_var, width=100)
        self.size_entry.pack(side="left", padx=5)
        
        # Skip existing files checkbox
        self.skip_checkbox = ctk.CTkCheckBox(self.settings_frame, text="Skip existing files", variable=self.skip_var)
        self.skip_checkbox.pack(side="left", padx=(20, 5))
        
        # Concurrent downloads setting
        self.concurrent_label = ctk.CTkLabel(self.settings_frame, text="Concurrent Downloads:")
        self.concurrent_label.pack(side="left", padx=(20, 5))
        
        self.concurrent_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.concurrent_var, width=100)
        self.concurrent_entry.pack(side="left", padx=5)
        
        # Control buttons frame
        self.control_frame = ctk.CTkFrame(self.main_frame)
        self.control_frame.pack(fill="x", pady=(0, 10))
        
        self.start_button = ctk.CTkButton(self.control_frame, text="Start Download", command=self.start_download)
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ctk.CTkButton(self.control_frame, text="Stop", command=self.stop_download, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        
        self.gallery_button = ctk.CTkButton(self.control_frame, text="Open Gallery", command=self.show_gallery)
        self.gallery_button.pack(side="left", padx=5)
        
        # Progress frame
        self.progress_frame = ctk.CTkFrame(self.main_frame)
        self.progress_frame.pack(fill="x", pady=(0, 10))
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(self.progress_frame, text="Ready")
        self.status_label.pack(pady=5)
        
        # Statistics frame
        self.stats_frame = ctk.CTkFrame(self.main_frame)
        self.stats_frame.pack(fill="x", pady=(0, 10))
        
        self.stats_label = ctk.CTkLabel(self.stats_frame, text="Statistics: ")
        self.stats_label.pack(side="left", padx=5)
        
        # Log frame with scrollable text
        self.log_frame = ctk.CTkFrame(self.main_frame)
        self.log_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.log_label = ctk.CTkLabel(self.log_frame, text="Log:")
        self.log_label.pack(anchor="w", padx=5)
        
        self.log_text = ctk.CTkTextbox(self.log_frame, height=150)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        
        logging.info("ImageDownloaderApp initialized")

    def browse_file(self):
        logging.info("Browse file dialog opened")
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            logging.info(f"Selected file: {filename}")
            self.file_path.set(filename)  # Use set() for StringVar
            self.log_message(f"Selected file: {filename}")
    
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
                progress = (self.completed_downloads / self.total_downloads) * 100
                self.progress_bar.set(progress)
                
                # Update progress text
                progress_text = f"Progress: {self.completed_downloads}/{self.total_downloads}"
                self.status_label.configure(text=progress_text)
                
                # Update statistics
                success_count = self.completed_downloads - self.failed_downloads - self.skipped_downloads
                stats_text = f"Completed: {success_count} | "
                stats_text += f"Skipped: {self.skipped_downloads} | "
                stats_text += f"Failed: {self.failed_downloads}"
                self.stats_label.configure(text=stats_text)
                logging.debug(f"Stats - Success: {success_count}, Skipped: {self.skipped_downloads}, Failed: {self.failed_downloads}, Total Progress: {self.completed_downloads}/{self.total_downloads}")
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

    def search_with_variations(self, description, max_results=5):
        """Try different search variations to find images"""
        logging.info(f"Searching with variations for: {description}")
        
        # Remove common words that might confuse the search
        words_to_remove = ['גרם', 'יח', 'ק.א', 'שקית', 'מארז']
        clean_desc = description
        for word in words_to_remove:
            clean_desc = clean_desc.replace(word, '')
        clean_desc = ' '.join(clean_desc.split())  # Remove extra spaces
        
        variations = [
            description,  # Original description
            clean_desc,   # Cleaned description
            f"מוצר {clean_desc}",  # Add "product" prefix
            f"{clean_desc} אריזה",  # Add "package" suffix
        ]
        
        all_results = []
        ddg = DDGS()
        
        for variation in variations:
            try:
                logging.info(f"Trying search variation: {variation}")
                results = list(ddg.images(variation, max_results=max_results))
                all_results.extend(results)
                if results:
                    logging.info(f"Found {len(results)} results for variation: {variation}")
                    break  # Stop if we found results
            except Exception as e:
                logging.error(f"Error searching with variation '{variation}': {str(e)}")
                continue
        
        # Remove duplicates while preserving order
        seen = set()
        unique_results = []
        for result in all_results:
            if result['image'] not in seen:
                seen.add(result['image'])
                unique_results.append(result)
        
        return unique_results[:max_results]  # Return at most max_results unique results

    def download_image(self, url, sku, output_dir, max_size, description):
        logging.info(f"Starting download for SKU: {sku}")
        output_path = os.path.join(output_dir, f"{sku}.jpg")
        
        # Skip if file exists and skip option is enabled
        if self.skip_var.get() and os.path.exists(output_path):
            logging.info(f"Skipping existing image for SKU {sku}")
            self.skipped_downloads += 1
            self.completed_downloads += 1
            self.update_progress()
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
                self.log_message(f"Successfully downloaded image for SKU {sku}")
                self.increment_counter('successful')
                if self.gallery_window:
                    self.window.after(0, lambda: self.gallery_window.add_image(sku, description, output_path))
                return True
        except Exception as e:
            logging.error(f"Error downloading image for SKU {sku}: {str(e)}")
            logging.error(traceback.format_exc())
            self.log_message(f"Error downloading image for SKU {sku}: {str(e)}")
            self.increment_counter('failed')
        return False
    
    def process_item(self, row, output_dir, max_size):
        try:
            sku = str(row['מק"ט'])
            description = str(row['תאור'])
            filename = f"{sku}.jpg"
            output_path = os.path.join(output_dir, filename)
            
            logging.info(f"Processing SKU: {sku}, Description: {description}")
            
            # Skip if file exists and skip option is enabled
            if os.path.exists(output_path) and self.skip_var.get():
                logging.info(f"Skipping existing file for SKU {sku}")
                self.skipped_downloads += 1
                self.completed_downloads += 1
                self.update_progress()
                return True
            
            # Search for images with variations
            try:
                results = self.search_with_variations(description)
                
                if not results:
                    logging.warning(f"No images found for SKU {sku} after trying variations")
                    self.failed_downloads += 1
                    self.completed_downloads += 1
                    self.update_progress()
                    return False
                
                # Try each result until we get a valid image
                for idx, result in enumerate(results, 1):
                    try:
                        image_url = result['image']
                        logging.info(f"Attempting download {idx}/{len(results)} from URL: {image_url}")
                        
                        # Download image with retries
                        max_retries = 3
                        retry_count = 0
                        while retry_count < max_retries:
                            try:
                                response = requests.get(image_url, timeout=10)
                                response.raise_for_status()
                                break
                            except requests.exceptions.RequestException as e:
                                retry_count += 1
                                if retry_count == max_retries:
                                    raise
                                logging.warning(f"Retry {retry_count}/{max_retries} for URL: {image_url}")
                                time.sleep(1)  # Wait before retry
                        
                        content_type = response.headers.get('content-type', '')
                        if not content_type.startswith('image/'):
                            logging.warning(f"Invalid content type: {content_type} for URL: {image_url}")
                            continue
                        
                        content_length = len(response.content)
                        if content_length < 1000:  # Skip tiny images
                            logging.warning(f"Image too small ({content_length} bytes) from URL: {image_url}")
                            continue
                            
                        logging.info(f"Downloaded {content_length} bytes with content type: {content_type}")
                        
                        # Save image temporarily
                        with open(output_path, 'wb') as f:
                            f.write(response.content)
                        
                        # Verify and process image
                        with Image.open(output_path) as img:
                            logging.info(f"Image opened successfully. Format: {img.format}, Size: {img.size}, Mode: {img.mode}")
                            
                            # Convert to RGB if necessary
                            if img.mode in ('RGBA', 'P'):
                                logging.info(f"Converting image from {img.mode} to RGB")
                                img = img.convert('RGB')
                            
                            # Resize if needed
                            if max(img.size) > max_size:
                                ratio = max_size / max(img.size)
                                new_size = tuple(int(dim * ratio) for dim in img.size)
                                logging.info(f"Resizing image from {img.size} to {new_size}")
                                img = img.resize(new_size, Image.Resampling.LANCZOS)
                            
                            # Save processed image
                            logging.info(f"Saving processed image to: {output_path}")
                            img.save(output_path, "JPEG", quality=85)
                        
                        self.successful_downloads += 1
                        logging.info(f"Successfully downloaded and processed image for SKU {sku}")
                        
                        # Update gallery if open
                        if self.gallery_window:
                            self.window.after(0, lambda s=sku, d=description, p=output_path: 
                                           self.gallery_window.add_image(s, d, p))
                        
                        break
                        
                    except requests.exceptions.RequestException as e:
                        logging.error(f"Network error downloading from {image_url}: {str(e)}")
                        if os.path.exists(output_path):
                            os.remove(output_path)
                        continue
                    except PIL.UnidentifiedImageError as e:
                        logging.error(f"Invalid image format from {image_url}: {str(e)}")
                        if os.path.exists(output_path):
                            os.remove(output_path)
                        continue
                    except Exception as e:
                        logging.error(f"Error processing image from {image_url}: {str(e)}")
                        logging.error(traceback.format_exc())
                        if os.path.exists(output_path):
                            os.remove(output_path)
                        continue
                
                else:  # No successful download from any URL
                    self.failed_downloads += 1
                    logging.warning(f"Failed to download any valid image for SKU {sku} after trying {len(results)} URLs")
            
            except Exception as e:
                logging.error(f"Error in DuckDuckGo search for SKU {sku}: {str(e)}")
                logging.error(traceback.format_exc())
                self.failed_downloads += 1
            
            self.completed_downloads += 1
            self.update_progress()
            
        except Exception as e:
            logging.error(f"Error processing item for SKU {sku}: {str(e)}")
            logging.error(traceback.format_exc())
            self.failed_downloads += 1
            self.completed_downloads += 1
            self.update_progress()
        return False
    
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
                        logging.info("Download process stopped")
                        break
                    logging.debug(f"Submitting task for row {index + 1}/{self.total_downloads}")
                    future = executor.submit(self.process_item, row, output_dir, max_size)
                    futures.append(future)
                
                # Wait for tasks to complete, but don't block the UI
                def check_futures():
                    if not self.is_running:
                        # Cancel any pending futures
                        for future in futures:
                            if not future.done():
                                future.cancel()
                        executor.shutdown(wait=False)
                        return
                        
                    if all(future.done() for future in futures):
                        if self.is_running:
                            logging.info("Download process completed successfully")
                            self.log_message("Download process completed!")
                            self.window.after(0, self.status_label.configure, {"text": "Download completed!"})
                            self.is_running = False
                            self.window.after(0, self.start_button.configure, {"state": "normal"})
                            self.window.after(0, self.stop_button.configure, {"state": "disabled"})
                    else:
                        self.window.after(100, check_futures)  # Check again in 100ms
                
                self.window.after(100, check_futures)
            
        except Exception as e:
            logging.error(f"Error in download process: {str(e)}")
            logging.error(traceback.format_exc())
            self.log_message(f"Error: {str(e)}")
            self.is_running = False
            self.window.after(0, self.start_button.configure, {"state": "normal"})
            self.window.after(0, self.stop_button.configure, {"state": "disabled"})
    
    def start_download(self):
        try:
            excel_path = self.file_path.get()
            if not excel_path:
                logging.warning("No Excel file selected")
                self.log_message("Please select an Excel file first!")
                return
            
            logging.info(f"Starting download with Excel file: {excel_path}")
            max_size = int(self.max_size_var.get())
            concurrent_limit = int(self.concurrent_var.get())
            logging.info(f"Parameters - Max size: {max_size}, Concurrent limit: {concurrent_limit}")
            
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.is_running = True
            
            # Start download process in a new thread
            thread = threading.Thread(target=self.download_process, args=(excel_path, max_size, concurrent_limit))
            thread.daemon = True  # Make thread daemon so it doesn't block program exit
            thread.start()
            
        except Exception as e:
            logging.error(f"Error starting download: {str(e)}")
            logging.error(traceback.format_exc())
            self.log_message(f"Error starting download: {str(e)}")
    
    def stop_download(self):
        if self.is_running:
            logging.info("Stopping download process")
            self.is_running = False
            self.status_label.configure(text="Download stopped")
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.log_message("Download process stopped by user")
    
    def show_gallery(self):
        if not self.gallery_window:
            self.gallery_window = ImageGalleryWindow(self.window)
            self.gallery_window.top.protocol("WM_DELETE_WINDOW", self.close_gallery)
            
        # Load existing images in batches
        output_dir = "downloaded_images"
        if os.path.exists(output_dir):
            images = [(f[:-4], os.path.join(output_dir, f)) for f in os.listdir(output_dir) if f.endswith('.jpg')]
            
            def load_batch(start_idx):
                batch_size = 5  # Load 5 images at a time
                end_idx = min(start_idx + batch_size, len(images))
                
                for i in range(start_idx, end_idx):
                    sku, image_path = images[i]
                    # Get description from excel data if available
                    description = next((str(row['תאור']) for _, row in pd.read_excel(self.file_path.get()).iterrows() 
                                     if str(row['מק"ט']) == sku), "No description")
                    self.gallery_window.add_image(sku, description, image_path)
                
                if end_idx < len(images):
                    # Schedule next batch
                    self.window.after(100, lambda: load_batch(end_idx))
            
            # Start loading first batch
            if images:
                load_batch(0)
    
    def close_gallery(self):
        if self.gallery_window:
            self.gallery_window.top.destroy()
            self.gallery_window = None
    
    def open_single_image_window(self):
        SingleImageWindow(self.window)
    
    def run(self):
        logging.info("Starting application")
        self.window.mainloop()

if __name__ == "__main__":
    log_file = setup_logging()
    app = ImageDownloaderApp()
    app.run()
