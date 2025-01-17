# Image Downloader for DuckDuckGo

A Python-based GUI application for downloading product images from DuckDuckGo search engine, with support for batch processing from Excel files.

## Features

### Batch Processing
- Read product data from Excel files (SKUs and descriptions)
- Concurrent downloads for improved performance
- Skip existing files option
- Progress tracking with detailed statistics
- Image gallery with preview functionality

### Single Image Mode
- Search and download individual images
- Preview before downloading
- Multiple search results navigation
- Custom filename support

### Image Processing
- Automatic image resizing
- Format validation and conversion
- Temporary storage for preview
- Cleanup of temporary files

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Required Python Packages
```bash
pip install -r requirements.txt
```

Required packages include:
- customtkinter: For modern GUI elements
- pandas: For Excel file processing
- Pillow: For image processing
- duckduckgo_search: For image searching
- requests: For downloading images
- fastai: For additional image processing capabilities

## Usage

### Starting the Application
```bash
python fetch_images.py
```

### Batch Download Mode
1. Click "Browse" to select your Excel file
2. Configure settings:
   - Max Image Size (default: 800px)
   - Concurrent Downloads (default: 3)
   - Skip Existing Files (enabled by default)
3. Click "Start Download" to begin
4. Monitor progress in the log window
5. Use "Open Gallery" to view downloaded images

### Single Image Mode
1. Click "Single Image Download"
2. Enter search description
3. Enter desired filename
4. Click "Search Image"
5. Use "Next Image" to browse results
6. Click "Save Image" when satisfied

### Excel File Format
Your Excel file must contain these columns:
- `מק"ט`: SKU/Product ID (string)
- `תאור`: Product Description (string)

Example:
| מק"ט | תאור |
|------|-------|
| 10001 | מוצר א |
| 10002 | מוצר ב |

### Output Structure
```
/downloaded_images/
    ├── [SKU].jpg         # Final images
    └── /temp/            # Temporary files
/logs/
    └── image_downloader_[TIMESTAMP].log
```

## Error Handling
- Network errors are automatically retried
- Invalid images are skipped
- Detailed error logging
- Progress is saved even if process is stopped

## Logging
- Logs are stored in `/logs` directory
- Each session creates a new timestamped log file
- Includes DEBUG level information for troubleshooting
- Console output for immediate feedback

## Performance
- Concurrent downloads (configurable)
- Image caching for gallery view
- Efficient memory management
- Progress updates are thread-safe

## Known Limitations
- Maximum concurrent downloads: 10
- Supported image formats: JPG, PNG
- Maximum image size: 5000px
- Hebrew text support required for descriptions

## Troubleshooting
1. If downloads fail:
   - Check internet connection
   - Verify Excel file format
   - Check log files for specific errors
   
2. If images are skipped:
   - Verify file permissions
   - Check disk space
   - Ensure valid product descriptions

## Contributing
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License
MIT License - See LICENSE file for details

## Version History
- 1.0.0 (2025-01-17)
  - Initial release
  - Basic GUI functionality
  - Batch and single image processing
  - Image gallery support

## Acknowledgments
- DuckDuckGo for search API
- CustomTkinter for modern GUI elements
- Contributors and testers

## Contact
For issues and feature requests, please use the GitHub issue tracker.

## Last Updated
2025-01-17
