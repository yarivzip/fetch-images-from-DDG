# Import only what we need initially
import os
import sys
import logging
import traceback
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox, filedialog
import threading
import config
import shutil
import time
from io import BytesIO

# Lazy imports - only import when needed
class LazyLoader:
    _pandas = None
    _ddgs = None
    _pillow = None
    _requests = None
    _concurrent_futures = None
    
    @classmethod
    def pandas(cls):
        if cls._pandas is None:
            import pandas as pd
            cls._pandas = pd
        return cls._pandas
    
    @classmethod
    def ddgs(cls):
        if cls._ddgs is None:
            from duckduckgo_search import DDGS
            cls._ddgs = DDGS
        return cls._ddgs
    
    @classmethod
    def pillow(cls):
        if cls._pillow is None:
            from PIL import Image, ImageTk
            cls._pillow = (Image, ImageTk)
        return cls._pillow
    
    @classmethod
    def requests(cls):
        if cls._requests is None:
            import requests
            cls._requests = requests
        return cls._requests
    
    @classmethod
    def concurrent_futures(cls):
        if cls._concurrent_futures is None:
            import concurrent.futures
            cls._concurrent_futures = concurrent.futures
        return cls._concurrent_futures

class ImageGalleryWindow:
    def __init__(self, parent):
        self.parent = parent
        self.top = ctk.CTkToplevel()
        self.top.title("Image Gallery")
        self.top.geometry("1200x800")
        
        # Create main frame with padding
        self.main_frame = ctk.CTkFrame(self.top)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create canvas with white background for better visibility
        self.canvas = ctk.CTkCanvas(self.main_frame, bg="white", highlightthickness=0)
        self.scrollbar = ctk.CTkScrollbar(self.main_frame, orientation="vertical", command=self.canvas.yview)
        
        # Create frame inside canvas for images
        self.scrollable_frame = ctk.CTkFrame(self.canvas, fg_color="white")
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window inside canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Configure canvas to expand with window
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Grid layout for canvas and scrollbar
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Bind canvas resizing
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        
        # Initialize image references and data
        self.image_references = {}
        self.current_row = 0
        self.images_per_row = 3
        self.image_frames = {}
        self.current_replacements = {}
        
    def _on_canvas_configure(self, event):
        # Update the scrollable region when the canvas is resized
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        
    def add_image(self, filename, description, image_path):
        try:
            # Create frame for this image
            image_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="gray90")
            image_frame.grid(row=self.current_row // self.images_per_row,
                           column=self.current_row % self.images_per_row,
                           padx=10, pady=10, sticky="nsew")
            
            # Load and resize image
            Image, ImageTk = LazyLoader.pillow()
            img = Image.open(image_path)
            
            # Convert to RGB if necessary
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
            
            # Resize image
            img.thumbnail((200, 200))
            photo = ImageTk.PhotoImage(img)
            
            # Keep reference to prevent garbage collection
            self.image_references[filename] = {
                'photo': photo,
                'description': description
            }
            
            # Create and pack image label
            img_label = ctk.CTkLabel(image_frame, image=photo, text="")
            img_label.pack(padx=5, pady=5)
            
            # Create and pack filename label - remove .jpg extension for display
            display_filename = filename[:-4] if filename.lower().endswith('.jpg') else filename
            name_label = ctk.CTkLabel(image_frame, text=display_filename, 
                                    font=("Helvetica", 12, "bold"))
            name_label.pack(padx=5)
            
            # Create and pack description label with larger wraplength and height
            desc_label = ctk.CTkLabel(image_frame, text=description, 
                                    wraplength=250,  
                                    font=("Helvetica", 12),  
                                    height=120)  
            desc_label.pack(padx=10, pady=(5, 10), fill="both", expand=True)  
            
            # Add Replace button
            replace_button = ctk.CTkButton(
                image_frame,
                text="Replace Image",
                command=lambda f=filename: self.get_replacement(f)
            )
            replace_button.pack(pady=5)
            
            # Add Approve button (initially disabled)
            approve_button = ctk.CTkButton(
                image_frame,
                text="Approve New",
                command=lambda f=filename: self.approve_replacement(f),
                state="disabled"
            )
            approve_button.pack(pady=5)
            
            # Store frame references
            self.image_frames[filename] = {
                'frame': image_frame,
                'image_label': img_label,
                'desc_label': desc_label,
                'approve_button': approve_button,
                'photo': photo
            }
            
            self.current_row += 1
            
            # Configure column weights
            self.scrollable_frame.grid_columnconfigure(self.current_row % self.images_per_row, weight=1)
            
        except Exception as e:
            logging.error(f"Error adding image to gallery: {str(e)}")
            logging.error(traceback.format_exc())
            
    def get_replacement(self, filename):
        try:
            if filename not in self.image_frames:
                return
                
            # Get original description from Excel
            pd = LazyLoader.pandas()
            if os.path.exists(self.parent.file_path.get()):
                df = pd.read_excel(self.parent.file_path.get())
                filename_col = self.parent.filename_column_var.get()
                desc_col = self.parent.description_column_var.get()
                
                # Find the matching row in Excel
                base_filename = filename
                if base_filename.lower().endswith('.jpg'):
                    base_filename = base_filename[:-4]
                    
                matching_row = df[df[filename_col].astype(str) == base_filename]
                if not matching_row.empty:
                    description = str(matching_row.iloc[0][desc_col])
                    logging.info(f"Found description for {filename}: {description}")
                else:
                    description = self.image_references[filename]['description']
                    logging.warning(f"No Excel description found for {filename}, using stored description")
            else:
                description = self.image_references[filename]['description']
                logging.warning("Excel file not found, using stored description")
            
            # Improve search by adding product context if available
            search_terms = []
            if "product" in description.lower():
                search_terms.append(description)
            else:
                search_terms.append(f"product {description}")
            search_terms.append(description)
            
            # Try each search term until we find results
            results = []
            for term in search_terms:
                logging.info(f"Searching for: {term}")
                results = self.parent.search_images(term)
                if results:
                    break
            
            if not results:
                messagebox.showinfo("Info", "No replacement images found")
                return
                
            # Find first unused image
            for result in results:
                image_url = result['image']
                
                # Skip if this URL was already tried
                if filename in self.current_replacements and image_url == self.current_replacements[filename].get('url'):
                    continue
                    
                try:
                    # Download and process image
                    requests = LazyLoader.requests()
                    response = requests.get(image_url, timeout=10)
                    
                    if response.status_code == 200:
                        Image, ImageTk = LazyLoader.pillow()
                        image_data = BytesIO(response.content)
                        img = Image.open(image_data)
                        
                        if img.mode in ('RGBA', 'P'):
                            img = img.convert('RGB')
                            
                        # Resize for preview
                        img.thumbnail((200, 200))
                        
                        # Update preview
                        photo = ImageTk.PhotoImage(img)
                        self.image_frames[filename]['image_label'].configure(image=photo)
                        self.image_frames[filename]['photo'] = photo
                        self.image_frames[filename]['approve_button'].configure(state="normal")
                        
                        # Store replacement data
                        self.current_replacements[filename] = {
                            'url': image_url,
                            'data': response.content,
                            'description': description  
                        }
                        
                        return
                        
                except Exception as e:
                    logging.error(f"Error downloading replacement image: {str(e)}")
                    continue
                    
            messagebox.showinfo("Info", "No more replacement images available")
            
        except Exception as e:
            logging.error(f"Error getting replacement: {str(e)}")
            messagebox.showerror("Error", f"Failed to get replacement: {str(e)}")
            
    def approve_replacement(self, filename):
        try:
            if filename not in self.current_replacements:
                return
                
            replacement_data = self.current_replacements[filename]
            output_dir = self.parent.download_dir_var.get()
            
            # Ensure filename has only one .jpg extension
            base_filename = filename[:-4] if filename.lower().endswith('.jpg') else filename
            target_filename = f"{base_filename}.jpg"
            
            target_path = os.path.join(output_dir, target_filename)
            
            # Save the new image
            with open(target_path, 'wb') as f:
                f.write(replacement_data['data'])
            
            # Update the description
            if 'description' in replacement_data:
                self.image_references[filename]['description'] = replacement_data['description']
            
            # Disable approve button
            if filename in self.image_frames:
                self.image_frames[filename]['approve_button'].configure(state="disabled")
            
            # Clear replacement data
            del self.current_replacements[filename]
            
            messagebox.showinfo("Success", "Image replaced successfully")
            
        except Exception as e:
            logging.error(f"Error approving replacement: {str(e)}")
            messagebox.showerror("Error", f"Failed to approve replacement: {str(e)}")

class SingleImageWindow:
    def __init__(self, parent):
        self.parent = parent
        self.top = ctk.CTkToplevel()
        self.top.title("Single Image Download")
        self.top.geometry("800x800")
        
        # Create main frame
        self.main_frame = ctk.CTkFrame(self.top)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Description entry
        self.desc_label = ctk.CTkLabel(self.main_frame, text="Description:")
        self.desc_label.pack(pady=(0, 5))
        
        self.desc_var = ctk.StringVar()
        self.desc_entry = ctk.CTkEntry(self.main_frame, textvariable=self.desc_var, width=400)
        self.desc_entry.pack(pady=(0, 20))
        
        # Filename entry
        self.filename_label = ctk.CTkLabel(self.main_frame, text="Filename:")
        self.filename_label.pack(pady=(0, 5))
        
        self.filename_var = ctk.StringVar()
        self.filename_entry = ctk.CTkEntry(self.main_frame, textvariable=self.filename_var, width=400)
        self.filename_entry.pack(pady=(0, 20))
        
        # Preview frame with white background
        self.preview_frame = ctk.CTkFrame(self.main_frame, fg_color="white")
        self.preview_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Image preview label
        self.image_label = ctk.CTkLabel(self.preview_frame, text="No image loaded")
        self.image_label.pack(pady=10, expand=True)
        
        # Buttons frame
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.pack(fill="x", pady=10)
        
        # Search button
        self.search_button = ctk.CTkButton(
            self.button_frame,
            text="Search Images",
            command=self.search_images
        )
        self.search_button.pack(side="left", padx=5)
        
        # Next button
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
            command=self.save_current_image,
            state="disabled"
        )
        self.save_button.pack(side="left", padx=5)
        
        # Status label
        self.status_var = ctk.StringVar(value="Enter description and filename to search")
        self.status_label = ctk.CTkLabel(self.main_frame, textvariable=self.status_var)
        self.status_label.pack(pady=10)
        
        # Initialize variables
        self.current_results = []
        self.current_index = 0
        self.current_image = None
        self.photo_reference = None  
        
    def search_images(self):
        description = self.desc_var.get().strip()
        if not description:
            self.status_var.set("Please enter a description")
            return
            
        self.status_var.set("Searching for images...")
        self.search_button.configure(state="disabled")
        
        try:
            # Search for images using parent's search function
            self.current_results = self.parent.search_images(description)
            
            if self.current_results:
                self.current_index = 0
                self.show_current_image()
                self.next_button.configure(state="normal")
            else:
                self.status_var.set("No images found for this description")
                
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            logging.error(f"Error searching images: {str(e)}")
            
        finally:
            self.search_button.configure(state="normal")
            
    def show_current_image(self):
        if not self.current_results:
            return
            
        try:
            result = self.current_results[self.current_index]
            image_url = result["image"]
            
            # Download and show preview
            requests = LazyLoader.requests()
            response = requests.get(image_url, timeout=10)
            
            if response.status_code == 200:
                Image, ImageTk = LazyLoader.pillow()
                image_data = BytesIO(response.content)
                img = Image.open(image_data)
                
                # Convert to RGB if necessary
                if img.mode in ('RGBA', 'P'):
                    img = img.convert('RGB')
                
                # Resize for preview while maintaining aspect ratio
                max_size = (400, 400)
                img.thumbnail(max_size, Image.Resampling.LANCZOS)
                
                # Store current image
                self.current_image = image_data.getvalue()
                
                # Update preview
                self.photo_reference = ImageTk.PhotoImage(img)
                self.image_label.configure(image=self.photo_reference, text="")
                self.save_button.configure(state="normal")
                self.status_var.set(f"Image {self.current_index + 1} of {len(self.current_results)}")
                
        except Exception as e:
            self.status_var.set(f"Error loading image: {str(e)}")
            logging.error(f"Error showing image: {str(e)}")
            self.next_image()  
            
    def next_image(self):
        if not self.current_results:
            return
            
        self.current_index = (self.current_index + 1) % len(self.current_results)
        self.show_current_image()
        
    def save_current_image(self):
        if not self.current_image:
            self.status_var.set("No image to save")
            return
            
        filename = self.filename_var.get().strip()
        if not filename:
            self.status_var.set("Please enter a filename")
            return
            
        try:
            # Ensure .jpg extension
            if not filename.lower().endswith('.jpg'):
                filename = f"{filename}.jpg"
                
            # Get output path
            output_path = os.path.join(self.parent.download_dir_var.get(), filename)
            
            # Save image
            Image, _ = LazyLoader.pillow()
            img = Image.open(BytesIO(self.current_image))
            
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
                
            # Resize if needed
            max_size = int(self.parent.max_size_var.get())
            if max(img.size) > max_size:
                ratio = max_size / max(img.size)
                new_size = tuple(int(dim * ratio) for dim in img.size)
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                
            img.save(output_path, "JPEG", quality=85)
            self.status_var.set("Image saved successfully!")
            self.top.destroy()
            
        except Exception as e:
            self.status_var.set(f"Error saving image: {str(e)}")
            logging.error(f"Error saving image: {str(e)}")

class ImageDownloaderApp:
    def __init__(self):
        logging.info("Initializing ImageDownloaderApp")
        self.window = ctk.CTk()
        self.window.title("Image Downloader")
        self.window.geometry("1000x800")  
        
        # Load user preferences
        self.config = config.load_config()
        
        # Initialize variables
        self.skip_var = ctk.BooleanVar(value=self.config["skip_existing"])
        self.max_size_var = ctk.StringVar(value=self.config["max_size"])
        self.concurrent_var = ctk.StringVar(value=self.config["concurrent_downloads"])
        self.description_column_var = ctk.StringVar(value=self.config["description_column"])
        self.filename_column_var = ctk.StringVar(value=self.config["filename_column"])
        self.download_dir_var = ctk.StringVar(value=self.config["download_directory"])
        self.file_path = ctk.StringVar()
        self.completed_downloads = 0
        self.failed_downloads = 0
        self.skipped_downloads = 0
        self.successful_downloads = 0
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
        
        self.file_label = ctk.CTkLabel(self.file_frame, text="Excel File:", width=100)
        self.file_label.pack(side="left", padx=5)
        
        self.file_entry = ctk.CTkEntry(self.file_frame, textvariable=self.file_path, width=600)
        self.file_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        self.browse_button = ctk.CTkButton(self.file_frame, text="Browse", command=self.browse_file, width=100)
        self.browse_button.pack(side="left", padx=5)

        # Download directory frame
        self.download_dir_frame = ctk.CTkFrame(self.main_frame)
        self.download_dir_frame.pack(fill="x", pady=(0, 10))
        
        self.download_dir_label = ctk.CTkLabel(self.download_dir_frame, text="Download Directory:", width=120)
        self.download_dir_label.pack(side="left", padx=5)
        
        self.download_dir_entry = ctk.CTkEntry(self.download_dir_frame, textvariable=self.download_dir_var, width=600)
        self.download_dir_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        self.download_dir_button = ctk.CTkButton(
            self.download_dir_frame, 
            text="Browse", 
            command=self.browse_download_dir,
            width=100
        )
        self.download_dir_button.pack(side="left", padx=5)
        
        # Column settings frame
        self.columns_frame = ctk.CTkFrame(self.main_frame)
        self.columns_frame.pack(fill="x", pady=(0, 10))
        
        # Description column setting
        self.desc_col_label = ctk.CTkLabel(self.columns_frame, text="Description Column:", width=120)
        self.desc_col_label.pack(side="left", padx=5)
        self.desc_col_entry = ctk.CTkEntry(self.columns_frame, textvariable=self.description_column_var, width=150)
        self.desc_col_entry.pack(side="left", padx=5)
        
        # Filename column setting
        self.filename_col_label = ctk.CTkLabel(self.columns_frame, text="Filename Column:", width=120)
        self.filename_col_label.pack(side="left", padx=(20, 5))
        self.filename_col_entry = ctk.CTkEntry(self.columns_frame, textvariable=self.filename_column_var, width=150)
        self.filename_col_entry.pack(side="left", padx=5)
        
        # Settings frame
        self.settings_frame = ctk.CTkFrame(self.main_frame)
        self.settings_frame.pack(fill="x", pady=(0, 10))
        
        # Image size setting
        self.size_label = ctk.CTkLabel(self.settings_frame, text="Max Image Size:", width=100)
        self.size_label.pack(side="left", padx=5)
        self.size_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.max_size_var, width=80)
        self.size_entry.pack(side="left", padx=5)
        
        # Skip existing files checkbox
        self.skip_checkbox = ctk.CTkCheckBox(self.settings_frame, text="Skip existing files", variable=self.skip_var)
        self.skip_checkbox.pack(side="left", padx=(20, 5))
        
        # Concurrent downloads setting
        self.concurrent_label = ctk.CTkLabel(self.settings_frame, text="Concurrent Downloads:", width=140)
        self.concurrent_label.pack(side="left", padx=(20, 5))
        self.concurrent_entry = ctk.CTkEntry(self.settings_frame, textvariable=self.concurrent_var, width=80)
        self.concurrent_entry.pack(side="left", padx=5)
        
        # Control buttons frame
        self.control_frame = ctk.CTkFrame(self.main_frame)
        self.control_frame.pack(fill="x", pady=(0, 10))
        
        self.start_button = ctk.CTkButton(self.control_frame, text="Start Download", command=self.start_download, width=150)
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ctk.CTkButton(self.control_frame, text="Stop", command=self.stop_download, state="disabled", width=100)
        self.stop_button.pack(side="left", padx=5)
        
        self.gallery_button = ctk.CTkButton(self.control_frame, text="Open Gallery", command=self.show_gallery, width=120)
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
        
        self.log_text = ctk.CTkTextbox(self.log_frame, height=200)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        
        logging.info("ImageDownloaderApp initialized")

    def browse_file(self):
        logging.info("Browse file dialog opened")
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            logging.info(f"Selected file: {file_path}")
            self.file_path.set(file_path)  
            self.log_message(f"Selected file: {file_path}")
    
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

    def search_images(self, query, max_results=5):
        """Search for images using DuckDuckGo"""
        try:
            ddgs = LazyLoader.ddgs()
            with ddgs() as ddg:
                results = list(ddg.images(
                    keywords=query,
                    max_results=max_results,
                    safesearch="off"
                ))
                logging.info(f"Found {len(results)} images for query: {query}")
                return results
        except Exception as e:
            logging.error(f"Error searching for images: {str(e)}")
            return []

    def process_item(self, row, output_dir, max_size):
        """Process a single item from the Excel file"""
        try:
            filename = str(row[self.filename_column_var.get()])
            description = str(row[self.description_column_var.get()])
            
            # Ensure filename ends with .jpg and is valid
            if not filename.lower().endswith('.jpg'):
                filename = f"{filename}.jpg"
            
            # Remove any invalid characters from filename
            filename = "".join(c for c in filename if c.isalnum() or c in (' ', '-', '_', '.'))
            
            output_path = os.path.join(self.download_dir_var.get(), filename)
            
            logging.info(f"Processing file: {filename}, Description: {description}")
            
            # Skip if file exists and skip option is enabled
            if os.path.exists(output_path) and self.skip_var.get():
                logging.info(f"Skipping existing file: {filename}")
                self.skipped_downloads += 1
                self.completed_downloads += 1
                self.update_progress()
                return
            
            # Create variations of the search query
            variations = [
                description,
                " ".join(description.split()[:4]),  
                f"product {description}",
                f"{description} package"
            ]
            
            logging.info(f"Searching with variations for: {description}")
            image_found = False
            
            for variation in variations:
                if not self.is_running:
                    return
                
                try:
                    logging.info(f"Trying search variation: {variation}")
                    results = self.search_images(variation)
                    
                    if results:
                        for result in results:
                            if not self.is_running:
                                return
                                
                            try:
                                image_url = result["image"]
                                logging.info(f"Attempting to download image from: {image_url}")
                                
                                # Download and process image
                                if self.download_and_save_image(image_url, output_path, max_size):
                                    image_found = True
                                    self.successful_downloads += 1
                                    break
                            except Exception as e:
                                logging.error(f"Error processing image result: {str(e)}")
                                continue
                    
                    if image_found:
                        break
                        
                except Exception as e:
                    logging.error(f"Error searching with variation '{variation}': {str(e)}")
                    continue
            
            if not image_found:
                logging.warning(f"No images found for filename {filename} after trying variations")
                self.failed_downloads += 1
            
            self.completed_downloads += 1
            self.update_progress()
            
        except Exception as e:
            logging.error(f"Error processing item: {str(e)}")
            logging.error(traceback.format_exc())
            self.failed_downloads += 1
            self.completed_downloads += 1
            self.update_progress()
        return False
    
    def download_and_save_image(self, url, output_path, max_size):
        try:
            logging.debug(f"Downloading image from URL: {url}")
            requests = LazyLoader.requests()
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                logging.debug(f"Download successful for URL: {url}")
                Image, _ = LazyLoader.pillow()
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
                return True
        except Exception as e:
            logging.error(f"Error downloading and saving image from {url}: {str(e)}")
            return False
    
    def download_process(self, excel_path, max_size, concurrent_limit):
        try:
            logging.info("Starting download process")
            output_dir = self.download_dir_var.get()
            os.makedirs(output_dir, exist_ok=True)
            
            logging.info(f"Reading Excel file: {excel_path}")
            pd = LazyLoader.pandas()
            df = pd.read_excel(excel_path)
            self.total_downloads = len(df)
            self.completed_downloads = 0
            self.skipped_downloads = 0
            self.failed_downloads = 0
            self.successful_downloads = 0
            logging.info(f"Total items to process: {self.total_downloads}")
            self.update_progress()
            
            concurrent = LazyLoader.concurrent_futures()
            with concurrent.ThreadPoolExecutor(max_workers=concurrent_limit) as executor:
                futures = []
                for index, row in df.iterrows():
                    if not self.is_running:
                        break
                    futures.append(executor.submit(self.process_item, row, output_dir, max_size))
                
                # Wait for all futures to complete
                for future in concurrent.as_completed(futures):
                    try:
                        future.result()
                    except Exception as e:
                        logging.error(f"Error in future: {str(e)}")
                        logging.error(traceback.format_exc())
            
            logging.info("Download process completed")
            messagebox.showinfo("Complete", "Download process completed!")
            
        except Exception as e:
            logging.error(f"Error in download process: {str(e)}")
            logging.error(traceback.format_exc())
            messagebox.showerror("Error", f"Error in download process: {str(e)}")
        finally:
            self.is_running = False
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
    
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
            
            # Save current preferences
            self.save_preferences()
            
            # Start download process in a new thread
            thread = threading.Thread(target=self.download_process, args=(excel_path, max_size, concurrent_limit))
            thread.daemon = True  
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
        """Show the image gallery window"""
        try:
            if self.gallery_window is None or not self.gallery_window.top.winfo_exists():
                self.gallery_window = ImageGalleryWindow(self)
                
                # Load existing images
                output_dir = self.download_dir_var.get()
                logging.info(f"Searching for images in directory: {output_dir}")
                
                # Debug current settings
                logging.info(f"Excel file path: {self.file_path.get()}")
                logging.info(f"Filename column: {self.filename_column_var.get()}")
                logging.info(f"Description column: {self.description_column_var.get()}")
                
                if not os.path.exists(output_dir):
                    logging.error(f"Output directory does not exist: {output_dir}")
                    messagebox.showerror("Error", f"Output directory not found: {output_dir}")
                    return
                    
                # Get Excel data if available
                descriptions = {}
                if os.path.exists(self.file_path.get()):
                    try:
                        pd = LazyLoader.pandas()
                        df = pd.read_excel(self.file_path.get())
                        filename_col = self.filename_column_var.get()
                        desc_col = self.description_column_var.get()
                        
                        logging.info(f"Excel columns found: {list(df.columns)}")
                        logging.info(f"Sample data: {df[[filename_col, desc_col]].head()}")
                        
                        for _, row in df.iterrows():
                            base_filename = str(row[filename_col]).strip()
                            description = str(row[desc_col]).strip()
                            
                            # Store both with and without .jpg
                            filename_with_jpg = f"{base_filename}.jpg"
                            descriptions[base_filename] = description
                            descriptions[filename_with_jpg] = description
                            
                            logging.info(f"Mapped description for {base_filename}: {description}")
                            
                        logging.info(f"Found {len(descriptions)} descriptions in Excel")
                        logging.info(f"Description mappings: {descriptions}")
                    except Exception as e:
                        logging.error(f"Error reading Excel file: {str(e)}")
                        logging.error(traceback.format_exc())
                        messagebox.showwarning("Warning", f"Error reading Excel file: {str(e)}")
                else:
                    logging.warning(f"Excel file not found: {self.file_path.get()}")
                
                # Add all images to gallery
                image_count = 0
                for filename in os.listdir(output_dir):
                    if filename.lower().endswith('.jpg'):
                        image_path = os.path.join(output_dir, filename)
                        base_filename = filename[:-4]  
                        
                        # Try to find description
                        description = descriptions.get(base_filename, descriptions.get(filename, "No description available"))
                        logging.info(f"Image {filename} -> base: {base_filename} -> description: {description}")
                        
                        self.gallery_window.add_image(filename, description, image_path)
                        image_count += 1
                
                logging.info(f"Added {image_count} images to gallery")
                
                if image_count == 0:
                    messagebox.showinfo("Info", f"No images found in {output_dir}")
                    return
                
                # Configure grid columns
                for i in range(self.gallery_window.images_per_row):
                    self.gallery_window.scrollable_frame.grid_columnconfigure(i, weight=1, minsize=250)
                
                # Update canvas scroll region
                self.gallery_window.scrollable_frame.update_idletasks()
                self.gallery_window.canvas.configure(scrollregion=self.gallery_window.canvas.bbox("all"))
                
        except Exception as e:
            logging.error(f"Error showing gallery: {str(e)}")
            logging.error(traceback.format_exc())
            messagebox.showerror("Error", f"Error showing gallery: {str(e)}")
    
    def open_single_image_window(self):
        """Open the single image download window"""
        try:
            SingleImageWindow(self)
        except Exception as e:
            logging.error(f"Error opening single image window: {str(e)}")
            messagebox.showerror("Error", f"Error opening single image window: {str(e)}")
    
    def save_preferences(self):
        """Save current settings to config file"""
        current_config = {
            "description_column": self.description_column_var.get(),
            "filename_column": self.filename_column_var.get(),
            "max_size": self.max_size_var.get(),
            "concurrent_downloads": self.concurrent_var.get(),
            "skip_existing": self.skip_var.get(),
            "download_directory": self.download_dir_var.get()
        }
        config.save_config(current_config)
        logging.info("Preferences saved")

    def browse_download_dir(self):
        """Browse for download directory"""
        dir_path = filedialog.askdirectory(
            initialdir=self.download_dir_var.get(),
            title="Select Download Directory"
        )
        if dir_path:
            self.download_dir_var.set(dir_path)
            logging.info(f"Download directory set to: {dir_path}")
            
    def run(self):
        logging.info("Starting application")
        self.window.mainloop()

if __name__ == "__main__":
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
                logging.StreamHandler()  
            ]
        )
        logging.info("Logging initialized")
        return log_file

    log_file = setup_logging()
    app = ImageDownloaderApp()
    app.run()
