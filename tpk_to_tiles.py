"""
TPK to TMS Tile Extractor
=========================
Converts an ArcMap Tile Package (.tpk) file into
a folder of tiles organized as {z}/{x}/{y}.png
so Tableau Desktop can use them as an offline map.

Usage:
    python tpk_to_tiles.py "C:\path\to\your_file.tpk" "C:\MapTiles\middle_east"

Requirements:
    - Python 2.7 or 3.x (ArcMap 10 ships with Python 2.7)
    - No extra libraries needed (uses only built-in modules)
"""

import os
import sys
import zipfile
import struct
import shutil
import math
import json


def extract_tpk_bundled(tpk_path, output_folder):
    """
    Extract tiles from a .tpk file (which is a ZIP archive)
    and reorganize them into {z}/{x}/{y}.png format.
    
    ArcGIS uses two tile storage formats inside .tpk:
      1. Exploded: individual tile files in _alllayers/L{zz}/R{row}/C{col}.png
      2. Bundled (v2): .bundle files containing multiple tiles
    
    This script handles BOTH formats.
    """
    
    if not os.path.exists(tpk_path):
        print("ERROR: File not found: " + tpk_path)
        print("Make sure you typed the full path correctly.")
        sys.exit(1)
    
    if not zipfile.is_zipfile(tpk_path):
        print("ERROR: This is not a valid .tpk file (not a ZIP archive).")
        sys.exit(1)
    
    # Create output folder
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    print("Opening TPK file: " + tpk_path)
    print("Output folder: " + output_folder)
    print("")
    
    tile_count = 0
    
    with zipfile.ZipFile(tpk_path, 'r') as zf:
        file_list = zf.namelist()
        
        # Print some info
        print("Files inside TPK: " + str(len(file_list)))
        
        # Check for conf.xml or conf.cdi to understand the tiling scheme
        for name in file_list:
            if name.lower().endswith('conf.xml'):
                print("Found config: " + name)
            if name.lower().endswith('conf.cdi'):
                print("Found CDI config: " + name)
        
        print("")
        
        # ============================================================
        # METHOD 1: Try exploded tile format first
        # Files like: tile/_alllayers/L00/R0000/C0000.png
        # ============================================================
        exploded_tiles = [f for f in file_list 
                          if '_alllayers' in f.lower() 
                          and (f.lower().endswith('.png') or f.lower().endswith('.jpg') or f.lower().endswith('.jpeg'))]
        
        if exploded_tiles:
            print("Found EXPLODED tile format (" + str(len(exploded_tiles)) + " tiles)")
            print("Extracting...")
            print("")
            
            for tile_path in exploded_tiles:
                try:
                    # Parse the path: .../L{zz}/R{row_hex}/C{col_hex}.ext
                    parts = tile_path.replace('\\', '/').split('/')
                    
                    level_part = None
                    row_part = None
                    col_part = None
                    
                    for i, part in enumerate(parts):
                        if part.upper().startswith('L') and len(part) == 3:
                            level_part = part
                        elif part.upper().startswith('R') and len(part) >= 5:
                            row_part = part
                        elif part.upper().startswith('C') and len(part) >= 5:
                            col_part = os.path.splitext(part)[0]
                            ext = os.path.splitext(part)[1]
                    
                    if not all([level_part, row_part, col_part]):
                        continue
                    
                    # Convert from ArcGIS format to TMS/Google format
                    zoom = int(level_part[1:])
                    row_arc = int(row_part[1:], 16)  # hex to int
                    col_arc = int(col_part[1:], 16)   # hex to int
                    
                    # ArcGIS uses top-left origin (same as Google/OSM for columns)
                    # TMS uses bottom-left origin for rows, but Tableau expects top-left (Google/OSM style)
                    # So we keep the row as-is (top-left origin = Google/OSM = what Tableau expects)
                    x = col_arc
                    y = row_arc
                    z = zoom
                    
                    # Create output path
                    out_dir = os.path.join(output_folder, str(z), str(x))
                    if not os.path.exists(out_dir):
                        os.makedirs(out_dir)
                    
                    out_file = os.path.join(out_dir, str(y) + ext)
                    
                    # Extract tile
                    tile_data = zf.read(tile_path)
                    with open(out_file, 'wb') as f:
                        f.write(tile_data)
                    
                    tile_count += 1
                    
                    if tile_count % 1000 == 0:
                        print("  Extracted " + str(tile_count) + " tiles...")
                
                except Exception as e:
                    print("  Warning: Could not process " + tile_path + " - " + str(e))
                    continue
        
        # ============================================================
        # METHOD 2: Try bundled tile format
        # Files like: tile/_alllayers/L00/R0000C0000.bundle
        # ============================================================
        bundle_files = [f for f in file_list 
                        if f.lower().endswith('.bundle')]
        
        if bundle_files and tile_count == 0:
            print("Found BUNDLED tile format (" + str(len(bundle_files)) + " bundle files)")
            print("Extracting (this may take a while)...")
            print("")
            
            for bundle_path in bundle_files:
                try:
                    # Parse bundle path to get zoom level and base row/col
                    parts = bundle_path.replace('\\', '/').split('/')
                    
                    level_part = None
                    bundle_name = None
                    
                    for part in parts:
                        if part.upper().startswith('L') and len(part) == 3:
                            level_part = part
                        if part.lower().endswith('.bundle'):
                            bundle_name = os.path.splitext(part)[0]
                    
                    if not level_part or not bundle_name:
                        continue
                    
                    zoom = int(level_part[1:])
                    
                    # Bundle name format: R{row_hex}C{col_hex}
                    r_idx = bundle_name.upper().index('R')
                    c_idx = bundle_name.upper().index('C')
                    base_row = int(bundle_name[r_idx+1:c_idx], 16)
                    base_col = int(bundle_name[c_idx+1:], 16)
                    
                    # Read bundle data
                    bundle_data = zf.read(bundle_path)
                    
                    # Bundle v2 format (compact cache v2):
                    # First 64 bytes: header
                    # Bytes 64-128: index (60 bytes of tile index info)
                    # After header: tile index entries
                    # Each bundle contains up to 128x128 tiles
                    
                    BUNDLE_SIZE = 128  # tiles per dimension in a bundle
                    INDEX_HEADER_SIZE = 64
                    
                    # Read tile index (each entry is 8 bytes)
                    num_entries = BUNDLE_SIZE * BUNDLE_SIZE
                    
                    for i in range(num_entries):
                        index_offset = INDEX_HEADER_SIZE + (i * 8)
                        
                        if index_offset + 8 > len(bundle_data):
                            break
                        
                        # Read 8-byte index entry (little-endian)
                        entry = struct.unpack('<Q', bundle_data[index_offset:index_offset+8])[0]
                        
                        # Offset is in lower 40 bits (5 bytes)
                        tile_offset = entry & 0xFFFFFFFFFF
                        # Size is in upper 24 bits (3 bytes)  
                        tile_size = (entry >> 40) & 0xFFFFFF
                        
                        if tile_size == 0 or tile_offset == 0:
                            continue
                        
                        if tile_offset + tile_size > len(bundle_data):
                            continue
                        
                        # Calculate row and col within bundle
                        row_in_bundle = i // BUNDLE_SIZE
                        col_in_bundle = i % BUNDLE_SIZE
                        
                        row = base_row + row_in_bundle
                        col = base_col + col_in_bundle
                        
                        # Extract tile data
                        tile_bytes = bundle_data[tile_offset:tile_offset+tile_size]
                        
                        if len(tile_bytes) < 8:
                            continue
                        
                        # Detect image format from magic bytes
                        ext = '.png'
                        if tile_bytes[:2] == b'\xff\xd8':
                            ext = '.jpg'
                        elif tile_bytes[:4] == b'\x89PNG':
                            ext = '.png'
                        
                        # Output as {z}/{x}/{y}.ext
                        x = col
                        y = row
                        z = zoom
                        
                        out_dir = os.path.join(output_folder, str(z), str(x))
                        if not os.path.exists(out_dir):
                            os.makedirs(out_dir)
                        
                        out_file = os.path.join(out_dir, str(y) + ext)
                        
                        with open(out_file, 'wb') as f:
                            f.write(tile_bytes)
                        
                        tile_count += 1
                        
                        if tile_count % 1000 == 0:
                            print("  Extracted " + str(tile_count) + " tiles...")
                
                except Exception as e:
                    print("  Warning: Could not process bundle " + bundle_path + " - " + str(e))
                    continue
    
    print("")
    print("=" * 50)
    print("DONE!")
    print("Total tiles extracted: " + str(tile_count))
    print("Output folder: " + output_folder)
    print("")
    
    if tile_count == 0:
        print("WARNING: No tiles were extracted!")
        print("Possible reasons:")
        print("  - The TPK file might be empty or corrupted")
        print("  - The tile format might not be recognized")
        print("  - Try re-exporting from ArcMap with different settings")
    else:
        print("Next steps:")
        print("  1. Copy the .tms file to: Documents\\My Tableau Repository\\Mapsources\\")
        print("  2. Make sure the path in the .tms file points to: " + output_folder)
        print("  3. Restart Tableau Desktop")
        print("  4. Go to Map > Background Maps > Middle East Offline Map")
    
    return tile_count


# ============================================================
# MAIN
# ============================================================
if __name__ == '__main__':
    
    if len(sys.argv) < 2:
        print("")
        print("TPK to TMS Tile Extractor")
        print("=" * 40)
        print("")
        print("Usage:")
        print('  python tpk_to_tiles.py "C:\\path\\to\\file.tpk" "C:\\MapTiles\\middle_east"')
        print("")
        print("Arguments:")
        print("  1st: Path to your .tpk file (required)")
        print("  2nd: Output folder (optional, defaults to C:\\MapTiles\\middle_east)")
        print("")
        sys.exit(0)
    
    tpk_file = sys.argv[1]
    
    if len(sys.argv) >= 3:
        out_folder = sys.argv[2]
    else:
        out_folder = "C:\\MapTiles\\middle_east"
    
    extract_tpk_bundled(tpk_file, out_folder)
