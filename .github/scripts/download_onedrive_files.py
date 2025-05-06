#!/usr/bin/env python3
# .github/scripts/download_onedrive_files.py

import os
import sys
import requests
from onedrivedownloader import download

def download_from_onedrive(share_url, output_path):
    """
    Download a file from OneDrive using a shared link.
    
    Args:
        share_url: The OneDrive share URL
        output_path: Path where the file should be saved
    """
    print(f"Downloading from: {share_url}")
    print(f"To: {output_path}")
    
    try:
        # Use onedrivedownloader library to handle the download
        download(share_url, filename=output_path, force_download=True)
        
        # Verify file size is not 0
        file_size = os.path.getsize(output_path)
        print(f"Download complete. File size: {file_size} bytes")
        
        if file_size == 0:
            print(f"WARNING: Downloaded file is empty: {output_path}")
            return False
        
        return True
    except Exception as e:
        print(f"Error downloading file: {str(e)}")
        return False

def main():
    # Create data directory if it doesn't exist
    os.makedirs("data", exist_ok=True)
    
    # Get OneDrive share URLs from environment variables
    solar_lab_tests_url = os.environ.get("SOLAR_LAB_TESTS_URL")
    line_trials_url = os.environ.get("LINE_TRIALS_URL")
    certifications_url = os.environ.get("CERTIFICATIONS_URL")
    
    success = True
    
    # Download Solar Lab Tests Excel file
    if solar_lab_tests_url:
        success = download_from_onedrive(solar_lab_tests_url, "data/Solar_Lab_Tests.xlsx") and success
    else:
        print("WARNING: SOLAR_LAB_TESTS_URL environment variable not set")
        success = False
    
    # Download Line Trials Excel file
    if line_trials_url:
        success = download_from_onedrive(line_trials_url, "data/Line_Trials.xlsx") and success
    else:
        print("WARNING: LINE_TRIALS_URL environment variable not set")
        success = False
    
    # Download Certifications Excel file
    if certifications_url:
        success = download_from_onedrive(certifications_url, "data/Certifications.xlsx") and success
    else:
        print("WARNING: CERTIFICATIONS_URL environment variable not set")
        success = False
    
    # Exit with error code if any download failed
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()
