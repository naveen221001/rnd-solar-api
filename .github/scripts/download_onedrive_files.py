#!/usr/bin/env python3
# .github/scripts/download_onedrive_files.py

import os
import sys
import time
import requests
import hashlib
import urllib.parse
from onedrivedownloader import download

def download_from_onedrive(share_url, output_path):
    """
    Download a file from OneDrive using a shared link with cache bypass.
    
    Args:
        share_url: The OneDrive share URL
        output_path: Path where the file should be saved
    """
    print(f"Downloading from: {share_url}")
    print(f"To: {output_path}")
    
    try:
        # Add a cache-busting timestamp parameter to the URL
        timestamp = int(time.time())
        
        # Check if URL already has parameters
        if '?' in share_url:
            cache_busting_url = f"{share_url}&_cb={timestamp}"
        else:
            cache_busting_url = f"{share_url}?_cb={timestamp}"
            
        print(f"Using cache-busting URL: {cache_busting_url}")
        
        # Use onedrivedownloader library with force download option
        download(cache_busting_url, filename=output_path, force_download=True)
        
        # Verify file size is not 0
        file_size = os.path.getsize(output_path)
        print(f"Download complete. File size: {file_size} bytes")
        
        if file_size == 0:
            print(f"WARNING: Downloaded file is empty: {output_path}")
            return False
        
        # If we have a previous version to compare against in Git, verify the file is different
        if os.path.exists(f"{output_path}.previous"):
            with open(output_path, 'rb') as f_new, open(f"{output_path}.previous", 'rb') as f_old:
                hash_new = hashlib.md5(f_new.read()).hexdigest()
                hash_old = hashlib.md5(f_old.read()).hexdigest()
                
                if hash_new == hash_old:
                    print(f"WARNING: Downloaded file is identical to previous version")
                    # Return True anyway since we have a valid file
                    return True
        
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
    changed = False
    
    # Download Solar Lab Tests Excel file
    if solar_lab_tests_url:
        # If the file exists, make a backup for comparison
        if os.path.exists("data/Solar_Lab_Tests.xlsx"):
            os.rename("data/Solar_Lab_Tests.xlsx", "data/Solar_Lab_Tests.xlsx.previous")
        
        result = download_from_onedrive(solar_lab_tests_url, "data/Solar_Lab_Tests.xlsx")
        success = result and success
        
        # If download was successful and we have a backup, compare them
        if result and os.path.exists("data/Solar_Lab_Tests.xlsx.previous"):
            with open("data/Solar_Lab_Tests.xlsx", 'rb') as f_new, open("data/Solar_Lab_Tests.xlsx.previous", 'rb') as f_old:
                hash_new = hashlib.md5(f_new.read()).hexdigest()
                hash_old = hashlib.md5(f_old.read()).hexdigest()
                
                if hash_new != hash_old:
                    print("Solar_Lab_Tests.xlsx has changed!")
                    changed = True
                else:
                    print("Solar_Lab_Tests.xlsx is identical to previous version")
            
            # Clean up the backup file if no longer needed
            os.remove("data/Solar_Lab_Tests.xlsx.previous")
    else:
        print("WARNING: SOLAR_LAB_TESTS_URL environment variable not set")
        success = False
    
    # Download Line Trials Excel file
    if line_trials_url:
        # If the file exists, make a backup for comparison
        if os.path.exists("data/Line_Trials.xlsx"):
            os.rename("data/Line_Trials.xlsx", "data/Line_Trials.xlsx.previous")
            
        result = download_from_onedrive(line_trials_url, "data/Line_Trials.xlsx")
        success = result and success
        
        # If download was successful and we have a backup, compare them
        if result and os.path.exists("data/Line_Trials.xlsx.previous"):
            with open("data/Line_Trials.xlsx", 'rb') as f_new, open("data/Line_Trials.xlsx.previous", 'rb') as f_old:
                hash_new = hashlib.md5(f_new.read()).hexdigest()
                hash_old = hashlib.md5(f_old.read()).hexdigest()
                
                if hash_new != hash_old:
                    print("Line_Trials.xlsx has changed!")
                    changed = True
                else:
                    print("Line_Trials.xlsx is identical to previous version")
            
            # Clean up the backup file if no longer needed
            os.remove("data/Line_Trials.xlsx.previous")
    else:
        print("WARNING: LINE_TRIALS_URL environment variable not set")
        success = False
    
    # Download Certifications Excel file
    if certifications_url:
        # If the file exists, make a backup for comparison
        if os.path.exists("data/Certifications.xlsx"):
            os.rename("data/Certifications.xlsx", "data/Certifications.xlsx.previous")
            
        result = download_from_onedrive(certifications_url, "data/Certifications.xlsx")
        success = result and success
        
        # If download was successful and we have a backup, compare them
        if result and os.path.exists("data/Certifications.xlsx.previous"):
            with open("data/Certifications.xlsx", 'rb') as f_new, open("data/Certifications.xlsx.previous", 'rb') as f_old:
                hash_new = hashlib.md5(f_new.read()).hexdigest()
                hash_old = hashlib.md5(f_old.read()).hexdigest()
                
                if hash_new != hash_old:
                    print("Certifications.xlsx has changed!")
                    changed = True
                else:
                    print("Certifications.xlsx is identical to previous version")
            
            # Clean up the backup file if no longer needed
            os.remove("data/Certifications.xlsx.previous")
    else:
        print("WARNING: CERTIFICATIONS_URL environment variable not set")
        success = False
    
    # Create a marker file to indicate if files have changed
    if changed:
        with open("data/.files_changed", "w") as f:
            f.write("true")
        print("Files have changed! Created marker file.")
    else:
        print("No files have changed.")
    
    # Exit with error code if any download failed
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()
