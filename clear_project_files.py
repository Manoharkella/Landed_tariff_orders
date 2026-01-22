import os

def remove_files_by_extension(root_folder, extensions):
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if any(file.lower().endswith(ext) for ext in extensions):
                try:
                    os.remove(os.path.join(root, file))
                    print(f"Removed: {os.path.join(root, file)}")
                except Exception as e:
                    print(f"Failed to remove {os.path.join(root, file)}: {e}")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    # Remove from Extraction and Download folders
    for folder in ["Extraction", "Download"]:
        folder_path = os.path.join(base_dir, folder)
        if os.path.exists(folder_path):
            remove_files_by_extension(folder_path, [".pdf", ".jsonl"])