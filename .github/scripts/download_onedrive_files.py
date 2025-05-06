#!/usr/bin/env python3
# .github/scripts/download_onedrive_files.py

import os
import sys
import time
import requests
import urllib.parse
import re

def get_direct_download_url(share_url):
    """
    Convert a OneDrive share URL to a direct download URL.
    
    Works for both OneDrive personal and Business accounts.
    """
    # Add timestamp to avoid caching
    timestamp = int(time.time())
    
    # Try to determine if it's a OneDrive personal or business link
    if "1drv.ms" in share_url:
        # For personal OneDrive short links (1drv.ms)
        # First do a HEAD request to get the redirect URL
        try:
            response = requests.head(share_url, allow_redirects=True)
            redirect_url = response.url
            
            # Convert the redirected URL to a direct download URL
            if "onedrive.live.com" in redirect_url:
                # Replace 'redir' with 'download' in the URL
                direct_url = redirect_url.replace("redir", "download")
                # Add timestamp to avoid caching
                direct_url = f"{direct_url}&_t={timestamp}"
                return direct_url
        except Exception as e:
            print(f"Error resolving 1drv.ms URL: {e}")
            return None
    
    elif "sharepoint.com" in share_url or "onedrive.live.com" in share_url:
        # For OneDrive business or regular OneDrive links
        try:
            if "download=1" not in share_url:
                # Add download parameter
                if "?" in share_url:
                    direct_url = f"{share_url}&download=1&_t={timestamp}"
                else:
                    direct_url = f"{share_url}?download=1&_t={timestamp}"
                return direct_url
            else:
                # Already has download parameter, just add timestamp
                direct_url = f"{share_url}&_t={timestamp}"
                return direct_url
        except Exception as e:
            print(f"Error creating direct URL: {e}")
            return None
    
    # If we couldn't determine the type, return the original URL
    return f"{share_url}?_t={timestamp}"

def download_file(url, output_path):
    """
    Download a file from a URL to the specified path.
    """
    print(f"Downloading from: {url}")
    print(f"To: {output_path}")
    
    try:
        direct_url = get_direct_download_url(url)
        if not direct_url:
            direct_url = url
            
        print(f"Using direct URL: {direct_url}")
        
        # Make the request with a custom User-Agent
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': '*/*',
            'Cache-Control': 'no-cache',
            'Pragma': 'no-cache'
        }
        
        response = requests.get(direct_url, headers=headers, stream=True)
        response.raise_for_status()
        
        # Get the content length if available
        total_size = int(response.headers.get('content-length', 0))
        print(f"Content length: {total_size} bytes")
        
        # Save the file
        with open(output_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        # Verify file size
        file_size = os.path.getsize(output_path)
        print(f"Download complete. File size: {file_size} bytes")
        
        if file_size == 0:
            print(f"WARNING: Downloaded file is empty: {output_path}")
            return False
        
        return True
    except Exception as e:
        print(f"Error downloading file: {str(e)}")
        return False

def force_changes():
    """Create a marker file to force Git to recognize changes"""
    marker_path = "data/.files_changed"
    with open(marker_path, "w") as f:
        f.write(f"Files updated at {time.time()}")
    print(f"Created marker file at {marker_path}")

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
        result = download_file(solar_lab_tests_url, "data/Solar_Lab_Tests.xlsx")
        success = result and success
    else:
        print("WARNING: SOLAR_LAB_TESTS_URL environment variable not set")
        success = False
    
    # Download Line Trials Excel file
    if line_trials_url:
        result = download_file(line_trials_url, "data/Line_Trials.xlsx")
        success = result and success
    else:
        print("WARNING: LINE_TRIALS_URL environment variable not set")
        success = False
    
    # Download Certifications Excel file
    if certifications_url:
        result = download_file(certifications_url, "data/Certifications.xlsx")
        success = result and success
    else:
        print("WARNING: CERTIFICATIONS_URL environment variable not set")
        success = False
    
    # Always force changes to be recognized
    if success:
        force_changes()
    
    # Exit with error code if any download failed
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()
