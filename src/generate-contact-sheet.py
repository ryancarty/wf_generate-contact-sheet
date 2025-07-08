import os
import sys
import subprocess
import re
from pathlib import Path
import math
import platform

# ==== DEPENDENCY CHECK ====
required_module = "PIL"
install_name = "Pillow"
missing = []
try:
    __import__(required_module)
except ImportError:
    missing.append(install_name)

if missing:
    missing_list = ', '.join(missing)
    print(f"⚠️ Missing packages: {missing_list}")
    print()
    install = input("Do you want to install them now? (y/n): ").strip().lower()
    print()
    if install == 'y':
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
        print("✅ Packages installed. Re-running script...")
        print()
        subprocess.call([sys.executable, *sys.argv])
        sys.exit()
    else:
        print("❌ Cannot continue without required packages.")
        sys.exit()

# ==== SAFE IMPORT (after dependency check) ====
from PIL import Image, ImageDraw, ImageFont

# ==== CONFIGURATION ====
def get_latest_versioned_files(renders_folder):
    version_pattern = re.compile(r"^(.*?)(\.v\d{3})\.(\d{4})\.png$")
    render_versions = {}

    for file in os.listdir(renders_folder):
        if file.endswith(".png"):
            match = version_pattern.match(file)
            if match:
                base_name = match.group(1).strip(".")
                version_str = match.group(2)
                version = int(re.search(r"v(\d{3})", version_str).group(1))

                if base_name not in render_versions:
                    render_versions[base_name] = {}

                if version not in render_versions[base_name]:
                    render_versions[base_name][version] = []

                full_path = os.path.join(renders_folder, file)
                render_versions[base_name][version].append(full_path)

    latest_files = []
    for base_name, versions in render_versions.items():
        latest_version = max(versions.keys())
        latest_files.extend(versions[latest_version])

    return latest_files


def gather_all_latest_images(root_folder):
    latest_images = []
    for root, dirs, files in os.walk(root_folder):
        if root.endswith(os.path.join("CMP", "work")):
            cmp_renders = os.path.join(root, "renders")
            if os.path.isdir(cmp_renders):
                cmp_images = get_latest_versioned_files(cmp_renders)
                if cmp_images:
                    latest_images.extend([(img, "CMP") for img in cmp_images])
                    continue

            # Fallback to LGT if CMP fails
            lgt_renders = root.replace(os.path.join("CMP", "work"), os.path.join("LGT", "work"))
            lgt_renders = os.path.join(lgt_renders, "renders")
            if os.path.isdir(lgt_renders):
                lgt_images = get_latest_versioned_files(lgt_renders)
                if lgt_images:
                    latest_images.extend([(img, "LGT") for img in lgt_images])

    return latest_images


# ==== TEXT COLOR BASED ON BACKGROUND ====
def get_text_color(bg):
    luminance = 0.299 * bg[0] + 0.587 * bg[1] + 0.114 * bg[2]
    return (255, 255, 255) if luminance < 128 else (0, 0, 0)


# ==== IMAGE RESIZING ====
def resize_to_height(img, height):
    w, h = img.size
    new_width = int((w / h) * height)
    return img.resize((new_width, height))


# ==== EXTRACT SH CODE ====
def extract_sh_code(filename):
    match = re.search(r"(sh\d+)", filename)
    return match.group(1) if match else ""


# ==== CONTACT SHEET CREATION ====
def create_contact_sheet(image_tuples, output_path, title_text, labeled=False):
    fixed_height = 200
    padding = 10
    title_height = 60
    bg_color = (255, 255, 255)
    text_color = get_text_color(bg_color)

    num_images = len(image_tuples)
    if num_images == 0:
        raise ValueError("No images to process.")

    columns = math.ceil(math.sqrt(num_images))
    rows = math.ceil(num_images / columns)

    while columns < rows:
        columns += 1
        rows = math.ceil(num_images / columns)

    resized_images = []
    max_widths_per_column = [0] * columns

    for i, (img_path, dept) in enumerate(image_tuples):
        img = Image.open(img_path)
        img = resize_to_height(img, fixed_height)
        resized_images.append((img_path, img, dept))

        col = i % columns
        max_widths_per_column[col] = max(max_widths_per_column[col], img.width)

    x_offsets = [padding]
    for w in max_widths_per_column[:-1]:
        x_offsets.append(x_offsets[-1] + w + padding)

    sheet_width = sum(max_widths_per_column) + (columns + 1) * padding
    sheet_height = title_height + rows * (fixed_height + padding) + padding

    contact_sheet = Image.new('RGB', (sheet_width, sheet_height), color=bg_color)
    draw = ImageDraw.Draw(contact_sheet)

    try:
        font = ImageFont.truetype("Arial.ttf", 32)
    except:
        font = ImageFont.load_default()

    bbox = font.getbbox(title_text)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    title_x = (sheet_width - text_width) // 2
    title_y = (title_height - text_height) // 2
    draw.text((title_x, title_y), title_text, fill=text_color, font=font)

    for i, (img_path, img, dept) in enumerate(resized_images):
        row = i // columns
        col = i % columns
        col_width = max_widths_per_column[col]

        x = x_offsets[col] + (col_width - img.width) // 2
        y = title_height + padding + row * (fixed_height + padding)

        contact_sheet.paste(img, (x, y))

        if labeled:
            filename = os.path.basename(img_path)
            sh_code = extract_sh_code(filename)
            version_match = re.search(r"\.v(\d{3})\.", filename)
            version_str = f"_v{version_match.group(1)}" if version_match else ""
            dept_str = f"_{dept}" if dept else ""
            label = f"{sh_code}{dept_str}{version_str}" if sh_code else f"{dept_str}{version_str}"

            if label:
                try:
                    label_font = ImageFont.truetype("Arial.ttf", 16)
                except:
                    label_font = ImageFont.load_default()
                draw.text((x_offsets[col] + 5, y + 5), label, fill=text_color, font=label_font)

    contact_sheet.save(output_path)
    print(f"✅ Saved: {output_path}")


# ==== MAIN EXECUTION ====
if __name__ == "__main__":
    # Set base path depending on OS
    if platform.system() == "Windows":
        base_path = r"\\csnzoo.com\services\imagedata\3dContent\PostProduction_WorkingLocation\Ryan\CODE\Contact-Sheet-Generator\Dummy-Server"
    elif platform.system() == "Darwin":  # macOS
        base_path = "/Volumes/3dContent/PostProduction_WorkingLocation/Ryan/CODE/Contact-Sheet-Generator/Dummy-Server"
    else:
        print("❌ Unsupported operating system.")
        sys.exit()

    folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f)) and f.startswith("25")]

    print()
    print("\033[35mAvailable top-level folders:\033[33m")
    for i, folder in enumerate(folders):
        print(f"{i + 1}. {folder}")

    choice = input("\033[32mSelect a folder by number: \033[0m").strip()
    if not choice.isdigit() or int(choice) < 1 or int(choice) > len(folders):
        print("Invalid selection.")
        exit()
    print()

    selected_top_folder = folders[int(choice) - 1]
    sequences_path = os.path.join(base_path, selected_top_folder, "sequences")

    if not os.path.isdir(sequences_path):
        print(f"No 'sequences' folder found in {selected_top_folder}")
        exit()

    sequence_folders = [f for f in os.listdir(sequences_path) if os.path.isdir(os.path.join(sequences_path, f))]

    if not sequence_folders:
        print("No subfolders found in the 'sequences' directory.")
        exit()

    print("\033[35mAvailable sequence folders:\033[33m")
    for i, folder in enumerate(sequence_folders):
        print(f"{i + 1}. {folder}")

    sub_choice = input("\033[32mSelect a sequence folder by number: \033[0m").strip()
    if not sub_choice.isdigit() or int(sub_choice) < 1 or int(sub_choice) > len(sequence_folders):
        print("Invalid selection.")
        exit()
    print()

    selected_sequence_folder = sequence_folders[int(sub_choice) - 1]
    input_folder = os.path.join(sequences_path, selected_sequence_folder)

    folder_name = os.path.basename(os.path.normpath(input_folder))
    image_tuples = gather_all_latest_images(input_folder)

    title = f"Contact Sheet for {folder_name}"
    output_dir = os.path.join(input_folder, "Contact-Sheets")
    os.makedirs(output_dir, exist_ok=True)

    version = 1
    for file in os.listdir(output_dir):
        match = re.search(rf"Contact-Sheet_{re.escape(folder_name)}(?:_labeled)?\.(\d+)\.jpg", file)
        if match:
            version = max(version, int(match.group(1)) + 1)

    old_dir = os.path.join(output_dir, "old")
    os.makedirs(old_dir, exist_ok=True)
    for file in os.listdir(output_dir):
        match = re.search(rf"Contact-Sheet_{re.escape(folder_name)}(?:_labeled)?\.(\d+)\.jpg", file)
        if match and int(match.group(1)) < version:
            existing_file = os.path.join(output_dir, file)
            if os.path.isfile(existing_file):
                os.rename(existing_file, os.path.join(old_dir, file))

    output_normal = os.path.join(output_dir, f"Contact-Sheet_{folder_name}.{version:03}.jpg")
    output_labeled = os.path.join(output_dir, f"Contact-Sheet_{folder_name}_labeled.{version:03}.jpg")

    create_contact_sheet(image_tuples, output_normal, title, labeled=False)
    create_contact_sheet(image_tuples, output_labeled, title, labeled=True)

    print("✅ Contact sheets generated.")

    if platform.system() == "Windows":
        os.startfile(output_dir)
    elif platform.system() == "Darwin":  # macOS
        os.system(f"open '{output_dir}'")
    elif platform.system() == "Linux":
        os.system(f"xdg-open '{output_dir}'")