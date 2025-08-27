#!/usr/bin/env python3
# .github/scripts/download_onedrive_files.py

import os
import sys
import time
import random
import requests
import urllib.parse
import re

def get_direct_download_url(share_url):
    """
    Convert a OneDrive share URL to a direct download URL with aggressive cache busting.
    """
    # Add timestamp and random string to avoid caching
    timestamp = int(time.time())
    random_str = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz0123456789', k=8))
    cache_buster = f"{timestamp}-{random_str}"
    
    print(f"Processing share URL: {share_url}")
    
    # Clean the URL first (remove any existing parameters)
    if '?' in share_url:
        base_url = share_url.split('?')[0]
    else:
        base_url = share_url
    
    # Try to determine if it's a OneDrive personal or business link
    if "1drv.ms" in share_url:
        # For personal OneDrive short links (1drv.ms)
        print("Detected OneDrive personal short link")
        try:
            # Set headers to mimic a browser
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Cache-Control': 'no-cache, no-store, must-revalidate',
                'Pragma': 'no-cache',
                'Expires': '0'
            }
            
            # First do a request to get the redirect URL
            response = requests.get(share_url, headers=headers, allow_redirects=True)
            redirect_url = response.url
            print(f"Redirected to: {redirect_url}")
            
            # Convert the redirected URL to a direct download URL
            if "onedrive.live.com" in redirect_url:
                # Replace view parameters with download
                if "view.aspx" in redirect_url:
                    direct_url = redirect_url.replace("view.aspx", "download.aspx")
                else:
                    direct_url = redirect_url
                
                # Add download parameter and cache buster
                if "?" in direct_url:
                    direct_url = f"{direct_url}&download=1&cb={cache_buster}"
                else:
                    direct_url = f"{direct_url}?download=1&cb={cache_buster}"
                
                print(f"Created direct URL: {direct_url}")
                return direct_url
        except Exception as e:
            print(f"Error resolving 1drv.ms URL: {e}")
            return None
    
    elif "sharepoint.com" in share_url or "onedrive.live.com" in share_url:
        # For OneDrive business or regular OneDrive links
        print("Detected OneDrive business or regular link")
        try:
            # Create direct URL with cache buster
            direct_url = f"{base_url}?download=1&cb={cache_buster}"
            print(f"Created direct URL: {direct_url}")
            return direct_url
        except Exception as e:
            print(f"Error creating direct URL: {e}")
            return None
    
    # If we couldn't determine the type, add cache buster to original URL
    print("Unknown URL type, adding cache buster to original URL")
    return f"{base_url}?cb={cache_buster}"

def download_file(url, output_path):
    """
    Download a file from a URL to the specified path with improved error handling and reporting.
    """
    print(f"\n=== Downloading from: {url}")
    print(f"To: {output_path}")
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            direct_url = get_direct_download_url(url)
            if not direct_url:
                print(f"Warning: Could not create direct URL, using original URL")
                direct_url = url
                
            print(f"Using direct URL (attempt {attempt+1}/{max_retries}): {direct_url}")
            
            # Make the request with a custom User-Agent and cache prevention
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': '*/*',
                'Cache-Control': 'no-cache, no-store, must-revalidate',
                'Pragma': 'no-cache',
                'Expires': '0'
            }
            
            # Use a session to maintain cookies and other state
            session = requests.Session()
            
            # First make a HEAD request to check content type and size
            head_response = session.head(direct_url, headers=headers, timeout=30)
            print(f"HEAD response status: {head_response.status_code}")
            print(f"HEAD response headers: {dict(head_response.headers)}")
            
            # Now make the GET request
            response = session.get(direct_url, headers=headers, stream=True, timeout=30)
            response.raise_for_status()
            
            # Debug response
            print(f"GET response status: {response.status_code}")
            print(f"GET response headers: {dict(response.headers)}")
            
            # Get the content length if available
            total_size = int(response.headers.get('content-length', 0))
            print(f"Content length: {total_size} bytes")
            
            # Get content type
            content_type = response.headers.get('content-type', '')
            print(f"Content type: {content_type}")
            
            # Verify it's Excel or a binary file
            if not (
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type or
                'application/vnd.ms-excel' in content_type or
                'application/octet-stream' in content_type or
                'application/binary' in content_type
            ):
                print(f"WARNING: Content type doesn't look like Excel: {content_type}")
                # Continue anyway - sometimes content type is misreported
            
            # Save the file
            downloaded_bytes = 0
            with open(output_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded_bytes += len(chunk)
                        # Print progress for large files
                        if total_size > 0 and downloaded_bytes % (1024*1024) == 0:
                            print(f"Downloaded {downloaded_bytes / (1024*1024):.2f} MB of {total_size / (1024*1024):.2f} MB")
            
            # Verify file size
            file_size = os.path.getsize(output_path)
            print(f"Download complete. File size: {file_size} bytes")
            
            if file_size == 0:
                print(f"WARNING: Downloaded file is empty: {output_path}")
                if attempt < max_retries - 1:
                    print(f"Retrying download...")
                    continue
                return False
            
            # Basic Excel file signature check
            with open(output_path, 'rb') as f:
                header = f.read(4)
                # XLSX files start with PK signature (zip file)
                if not (header.startswith(b'PK') or header.startswith(b'\xd0\xcf\x11\xe0')):
                    print(f"WARNING: File doesn't look like an Excel file. First bytes: {header.hex()}")
                    # Continue anyway - it might be valid in some cases
            
            print(f"Successfully downloaded {url} to {output_path}")
            return True
            
        except Exception as e:
            print(f"Error downloading file (attempt {attempt+1}/{max_retries}): {str(e)}")
            if attempt < max_retries - 1:
                print(f"Retrying in 5 seconds...")
                time.sleep(5)
            else:
                return False
    
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
    chamber_tests_url = os.environ.get("CHAMBER_TESTS_URL")
    
    print("\n=== Starting OneDrive file download process ===")
    print(f"Current directory: {os.getcwd()}")
    print(f"Data directory: {os.path.abspath('data')}")
    
    success = True
    
    # Download Solar Lab Tests Excel file
    if solar_lab_tests_url:
        print("\n=== Processing Solar Lab Tests file ===")
        result = download_file(solar_lab_tests_url, "data/Solar_Lab_Tests.xlsx")
        success = result and success
    else:
        print("WARNING: SOLAR_LAB_TESTS_URL environment variable not set")
        success = False
    
    # Download Line Trials Excel file
    if line_trials_url:
        print("\n=== Processing Line Trials file ===")
        result = download_file(line_trials_url, "data/Line_Trials.xlsx")
        success = result and success
    else:
        print("WARNING: LINE_TRIALS_URL environment variable not set")
        success = False
    
    # Download Certifications Excel file
    if certifications_url:
        print("\n=== Processing Certifications file ===")
        result = download_file(certifications_url, "data/Certifications.xlsx")
        success = result and success
    else:
        print("WARNING: CERTIFICATIONS_URL environment variable not set")
        success = False

    # ADD THIS NEW SECTION - Download Chamber Tests Excel file
    if chamber_tests_url:
        print("\n=== Processing Chamber Tests file ===")
        result = download_file(chamber_tests_url, "data/Chamber_Tests.xlsx")
        success = result and success
    else:
        print("WARNING: CHAMBER_TESTS_URL environment variable not set")
        success = False


    rnd_todos_url = os.environ.get("RND_TODOS_URL")
    if rnd_todos_url:
        print("\n=== Processing R&D Todos file ===")
        result = download_file(rnd_todos_url, "data/RND_Todos.xlsx")
        success = result and success
    else:
        print("WARNING: RND_TODOS_URL environment variable not set")
        success = False

    daily_updates_url = os.environ.get("DAILY_UPDATES_URL")
    if daily_updates_url:
        print("\n=== Processing Daily Updates file ===")
        result = download_file(daily_updates_url, "data/Daily_Updates.xlsx")
        success = result and success
    else:
        print("WARNING: DAILY_UPDATES_URL environment variable not set")
        success = False
    # Always force changes to be recognized
    force_changes()
    
    print("\n=== Download process completed ===")
    print(f"Overall success: {success}")
    
    # List files in data directory
    print("\n=== Files in data directory ===")
    for file in os.listdir("data"):
        file_path = os.path.join("data", file)
        if os.path.isfile(file_path):
            size = os.path.getsize(file_path)
            print(f"{file}: {size} bytes")
    
    # Exit with error code if any download failed
    if not success:
        print("Exiting with error code 1 due to download failures")
        sys.exit(1)
    else:
        print("All downloads successful")

if __name__ == "__main__":
    main()
