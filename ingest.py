#!/usr/bin/env python3
"""
Folder Ingestion Script
Monitors a source directory for new subfolders, validates naming conventions,
moves folders to a destination, and logs file details to Excel.
"""

import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Tuple, List, Dict
import yaml
import pandas as pd
from openpyxl import load_workbook, Workbook


def load_config(config_path: str = "config.yaml") -> dict:
    """Load configuration from YAML file."""
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)


def normalize_folder_name(name: str, pattern: str) -> Tuple[str, str]:
    """
    Validate and normalize folder name according to naming convention.
    
    Args:
        name: Original folder name
        pattern: Regex pattern for validation
        
    Returns:
        Tuple of (normalized_name, flag) where flag is 'OK' or 'X'
    """
    # Check if name already matches the pattern
    if re.match(pattern, name):
        return name, 'OK'
    
    # Try to fix common issues: replace underscores/hyphens with spaces
    # Look for pattern: digits followed by underscore/hyphen
    fixed_name = re.sub(r'(\d{4}(?:\.\d{2})?)[-_]', r'\1 ', name)
    
    # Check if the fix worked
    if re.match(pattern, fixed_name):
        return fixed_name, 'OK'
    
    # Try another approach: add space after project number if missing
    fixed_name = re.sub(r'^(\d{4}(?:\.\d{2})?)([^\s])', r'\1 \2', name)
    
    if re.match(pattern, fixed_name):
        return fixed_name, 'OK'
    
    # If all fixes fail, return original name with flag
    return name, 'X'


def get_folder_files(folder_path: Path) -> List[Dict[str, any]]:
    """
    Get all files in a folder with their creation dates.
    
    Args:
        folder_path: Path to the folder
        
    Returns:
        List of dictionaries with file info
    """
    files_info = []
    
    for item in folder_path.rglob('*'):
        if item.is_file():
            created_time = datetime.fromtimestamp(item.stat().st_ctime)
            files_info.append({
                'file_name': item.name,
                'file_path': str(item.relative_to(folder_path)),
                'created_date': created_time
            })
    
    return files_info


def move_folder(source: Path, destination_root: Path) -> Path:
    """
    Move folder to destination directory.
    
    Args:
        source: Source folder path
        destination_root: Destination root directory
        
    Returns:
        Path to the moved folder
    """
    dest_path = destination_root / source.name
    
    # Check if destination already exists
    if dest_path.exists():
        # Add timestamp to avoid conflicts
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest_path = destination_root / f"{source.name}_{timestamp}"
    
    shutil.move(str(source), str(dest_path))
    return dest_path


def log_to_excel(data: List[Dict[str, any]], excel_path: str, retry_attempts: int, retry_delay: int):
    """
    Log folder processing details to Excel with retry logic for locked files.
    
    Args:
        data: List of dictionaries containing log data
        excel_path: Path to Excel file
        retry_attempts: Number of retry attempts
        retry_delay: Delay between retries in seconds
    """
    df = pd.DataFrame(data)
    
    for attempt in range(retry_attempts):
        try:
            # Check if Excel file exists
            if os.path.exists(excel_path):
                # Read existing data
                existing_df = pd.read_excel(excel_path)
                # Append new data
                combined_df = pd.concat([existing_df, df], ignore_index=True)
            else:
                combined_df = df
            
            # Write to Excel
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Processing Log')
            
            print(f"Successfully logged {len(data)} entries to Excel.")
            return
            
        except PermissionError as e:
            if attempt < retry_attempts - 1:
                print(f"Excel file is locked. Retry {attempt + 1}/{retry_attempts} in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                print(f"ERROR: Could not write to Excel after {retry_attempts} attempts. File may be open.")
                raise e


def get_processed_folders(excel_path: str) -> set:
    """
    Get list of already processed folder names from Excel log.
    
    Args:
        excel_path: Path to Excel file
        
    Returns:
        Set of processed folder names
    """
    if not os.path.exists(excel_path):
        return set()
    
    try:
        df = pd.read_excel(excel_path)
        if 'Folder Name' in df.columns:
            return set(df['Folder Name'].unique())
    except Exception as e:
        print(f"Warning: Could not read existing log: {e}")
    
    return set()


def process_folders(config: dict):
    """
    Main processing function to scan, validate, move, and log folders.
    
    Args:
        config: Configuration dictionary
    """
    source_dir = Path(config['source_dir'])
    destination_dir = Path(config['destination_dir'])
    excel_path = config['excel_log_path']
    pattern = config['naming_pattern']
    retry_attempts = config['retry_attempts']
    retry_delay = config['retry_delay_seconds']
    
    # Ensure directories exist
    if not source_dir.exists():
        print(f"ERROR: Source directory does not exist: {source_dir}")
        return
    
    destination_dir.mkdir(parents=True, exist_ok=True)
    
    # Get already processed folders
    processed_folders = get_processed_folders(excel_path)
    print(f"Found {len(processed_folders)} already processed folders in log.")
    
    # Get all subfolders in source directory
    subfolders = [f for f in source_dir.iterdir() if f.is_dir()]
    
    if not subfolders:
        print("No subfolders found in source directory.")
        return
    
    print(f"Found {len(subfolders)} subfolders to process.")
    
    # Process each subfolder
    for folder in subfolders:
        folder_name = folder.name
        
        # Skip if already processed
        if folder_name in processed_folders:
            print(f"Skipping '{folder_name}' - already processed.")
            continue
        
        # Skip if already exists in destination
        if (destination_dir / folder_name).exists():
            print(f"Skipping '{folder_name}' - already exists in destination.")
            continue
        
        print(f"\nProcessing folder: {folder_name}")
        
        # Validate and normalize name
        normalized_name, flag = normalize_folder_name(folder_name, pattern)
        
        if flag == 'X':
            print(f"  WARNING: Folder name does not match convention. Flagged with 'X'.")
        else:
            print(f"  Name validated: OK")
        
        # Rename folder if normalized name is different
        if normalized_name != folder_name:
            new_folder_path = folder.parent / normalized_name
            folder.rename(new_folder_path)
            folder = new_folder_path
            print(f"  Renamed to: {normalized_name}")
        
        # Get all files in the folder
        files_info = get_folder_files(folder)
        print(f"  Found {len(files_info)} files in folder.")
        
        # Move folder to destination
        try:
            moved_path = move_folder(folder, destination_dir)
            print(f"  Moved to: {moved_path}")
        except Exception as e:
            print(f"  ERROR: Could not move folder: {e}")
            continue
        
        # Prepare log data
        processed_date = datetime.now()
        log_data = []
        
        if files_info:
            for file_info in files_info:
                log_data.append({
                    'Folder Name': normalized_name,
                    'Naming Flag': flag,
                    'Processed Date': processed_date,
                    'File Name': file_info['file_name'],
                    'File Path': file_info['file_path'],
                    'File Created Date': file_info['created_date']
                })
        else:
            # Log folder even if empty
            log_data.append({
                'Folder Name': normalized_name,
                'Naming Flag': flag,
                'Processed Date': processed_date,
                'File Name': '(empty)',
                'File Path': '',
                'File Created Date': None
            })
        
        # Write to Excel log
        try:
            log_to_excel(log_data, excel_path, retry_attempts, retry_delay)
            print(f"  Logged to Excel successfully.")
        except Exception as e:
            print(f"  ERROR: Could not log to Excel: {e}")


def main():
    """Main entry point for the script."""
    print("="*60)
    print("Folder Ingestion Script")
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    try:
        # Load configuration
        config = load_config()
        print(f"\nConfiguration loaded:")
        print(f"  Source: {config['source_dir']}")
        print(f"  Destination: {config['destination_dir']}")
        print(f"  Log: {config['excel_log_path']}")
        
        # Process folders
        process_folders(config)
        
        print("\n" + "="*60)
        print("Processing complete.")
        print("="*60)
        
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        raise


if __name__ == "__main__":
    main()
