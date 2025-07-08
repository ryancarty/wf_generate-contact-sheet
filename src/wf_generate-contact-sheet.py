import os
import sys
import subprocess
import re
from pathlib import Path
import math
import platform

# ==== DEPENDENCY CHECK ====
dependencies = {
    "PIL": "Pillow",
    "colorama": "colorama"
}

missing = []
for module_name, install_name in dependencies.items():
    try:
        __import__(module_name)
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
from colorama import init, Fore, Style

# Initialize colorama
init()

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

            lgt_renders = root.replace(os.path.join("CMP", "work"), os.path.join("LGT", "work"))
            lgt_renders = os.path.join(lgt_renders, "renders")
            if os.path.isdir(lgt_renders):
                lgt_images = get_latest_versioned_files(lgt_renders)
                if lgt_images:
                    latest_images.extend([(img, "LGT") for img in lgt_images])

    return latest_images


def get_text_color(bg):
    luminance = 0.299 * bg[0] + 0.587 * bg[1] + 0.114 * bg[2]
    return (255, 255, 255) if luminance < 128 else (0, 0, 0)


def resize_to_height(img, height):
    w, h = img.size
    new_width = int((w / h) * height)
    return img.resize((new_width, height))


def extract_sh_code(filename):
    match = re.search(r"(sh\d+)", filename)
    return match.group(1) if match else ""


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
        if platform.system() == "Windows":
            label_font = ImageFont.truetype("arial.ttf", 32)
        elif platform.system() == "Darwin":
            label_font = ImageFont.truetype("Arial.ttf", 32)
        elif platform.system() == "Linux":
            label_font = ImageFont.truetype("Arial.ttf", 32)
        
    except Exception as e:
        print("⚠️ Font fallback activated:", e)
        label_font = ImageFont.load_default()

    bbox = label_font.getbbox(title_text)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    title_x = (sheet_width - text_width) // 2
    title_y = (title_height - text_height) // 2
    draw.text((title_x, title_y), title_text, fill=text_color, font=label_font)

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
                    if platform.system() == "Windows":
                        label_font = ImageFont.truetype("arial.ttf", 16)
                    elif platform.system() == "Darwin":
                        label_font = ImageFont.truetype("Arial.ttf", 16)
                    elif platform.system() == "Linux":
                        label_font = ImageFont.truetype("Arial.ttf", 16)
                except Exception as e:
                    print("⚠️ Font fallback activated:", e)
                    label_font = ImageFont.load_default()
                draw.text((x_offsets[col] + 5, y + 5), label, fill=text_color, font=label_font)

    contact_sheet.save(output_path)
    print(f"{Fore.GREEN}✅ Saved: {output_path}{Style.RESET_ALL}")


# ==== MAIN EXECUTION ====
if __name__ == "__main__":
    if platform.system() == "Windows":
        #base_path = r"\\csnzoo.com\services\imagedata\3dContent\PostProduction_WorkingLocation\Ryan\CODE\Contact-Sheet-Generator\Dummy-Server"
        base_path = r"\\csnzoo.com\services\imagedata\3dContent\sg_flow"
    elif platform.system() == "Darwin":
        #base_path = "/Volumes/3dContent/PostProduction_WorkingLocation/Ryan/CODE/Contact-Sheet-Generator/Dummy-Server"
        base_path = "/Volumes/3dContent/sg_flow"
    else:
        print("❌ Unsupported operating system.")
        sys.exit()

    folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f)) and f.startswith("25")]

    print()
    print(Fore.MAGENTA + "Available top-level folders:" + Fore.YELLOW)
    for i, folder in enumerate(folders):
        print(f"{i + 1}. {folder}")
    print(Style.RESET_ALL)

    choice = input(Fore.GREEN + "Select a folder by number: " + Style.RESET_ALL).strip()
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

    print(Fore.MAGENTA + "Available sequence folders:" + Fore.YELLOW)
    for i, folder in enumerate(sequence_folders):
        print(f"{i + 1}. {folder}")
    print(Style.RESET_ALL)

    sub_choice = input(Fore.GREEN + "Select a sequence folder by number: " + Style.RESET_ALL).strip()
    if not sub_choice.isdigit() or int(sub_choice) < 1 or int(sub_choice) > len(sequence_folders):
        print("Invalid selection.")
        exit()
    print()

    print("Generating contact sheet...folder will open when it is complete...")

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
            old_file_path = os.path.join(old_dir, file)
            if os.path.isfile(existing_file):
                if os.path.exists(old_file_path):
                    os.remove(old_file_path)  # Delete existing file to avoid WinError 183
                os.rename(existing_file, old_file_path)

            #if os.path.isfile(existing_file):
                #os.rename(existing_file, os.path.join(old_dir, file))

    output_normal = os.path.join(output_dir, f"Contact-Sheet_{folder_name}.{version:03}.jpg")
    output_labeled = os.path.join(output_dir, f"Contact-Sheet_{folder_name}_labeled.{version:03}.jpg")

    create_contact_sheet(image_tuples, output_normal, title, labeled=False)
    print("One more to go...")
    create_contact_sheet(image_tuples, output_labeled, title, labeled=True)

    print(Fore.GREEN + "✅ Contact sheets generated." + Style.RESET_ALL)

    if platform.system() == "Windows":
        os.startfile(output_dir)
    elif platform.system() == "Darwin":
        os.system(f"open '{output_dir}'")
    elif platform.system() == "Linux":
        os.system(f"xdg-open '{output_dir}'")



    # ==== OPTIONAL NUKE SCRIPT EXECUTION ====
    make_nuke = input(Fore.CYAN + "\nWould you like to generate a Nuke script from these images? (y/n): " + Style.RESET_ALL).strip().lower()
    if make_nuke == "y":
        print("Generating Nuke script...")

        if not image_tuples:
            print(Fore.RED + "❌ No images found to create a Nuke script." + Style.RESET_ALL)
            sys.exit(1)

        print(f"✅ Preparing {len(image_tuples)} images for the Nuke script...")

        nuke_script_name = f"Contact-Sheet_{folder_name}.{version:03}.nk"
        nuke_script_path = os.path.join(output_dir, nuke_script_name)

        # Ensure 'old' folder exists
        os.makedirs(old_dir, exist_ok=True)

        # Move any old .nk or .nk.autosave files
        for file in os.listdir(output_dir):
            if re.match(rf"Contact-Sheet_{re.escape(folder_name)}\.\d+\.(nk|nk\.autosave)$", file):
                existing_file = os.path.join(output_dir, file)
                old_file_path = os.path.join(old_dir, file)
                if os.path.isfile(existing_file):
                    if os.path.exists(old_file_path):
                        os.remove(old_file_path)
                    os.rename(existing_file, old_file_path)

                # ==== Nuke Node Layout ====
        spacing = 180
        start_x = 0
        lgt_ypos = 0
        cmp_ypos = 200
        contact_ypos = 400

        sorted_images = sorted(image_tuples, key=lambda x: os.path.basename(x[0]))
        print(f"🧪 Generating Nuke script for {len(sorted_images)} images...")

        read_positions = []
        nuke_lines = []

        for idx, (img_path, dept) in enumerate(sorted_images):
            file_path = img_path.replace("\\", "/")
            xpos = start_x + idx * spacing
            ypos = lgt_ypos if dept == "LGT" else cmp_ypos

            # Read node
            nuke_lines.append("Read {")
            nuke_lines.append(f" file \"{file_path}\"")
            nuke_lines.append(f" name Read{idx+1}")
            nuke_lines.append(f" xpos {xpos}")
            nuke_lines.append(f" ypos {ypos}")
            nuke_lines.append("}")
            nuke_lines.append("")

            # Dot node below the read
            nuke_lines.append("Dot {")
            nuke_lines.append(f" name Dot{idx+1}")
            nuke_lines.append(f" xpos {xpos+34}")
            nuke_lines.append(f" ypos {contact_ypos}")
            nuke_lines.append("}")
            nuke_lines.append("")

            read_positions.append(xpos)

        center_x = (read_positions[0] + read_positions[-1]) // 2 if read_positions else 0

        # ContactSheet node, connected to Dots
        nuke_lines.append("ContactSheet {")
        nuke_lines.append(f" inputs {len(sorted_images)}")
        nuke_lines.append(" width 4096 height 4096")
        nuke_lines.append(f" rows {math.ceil(math.sqrt(len(sorted_images)))}")
        nuke_lines.append(f" columns {math.ceil(math.sqrt(len(sorted_images)))}")
        nuke_lines.append(" center true")
        nuke_lines.append(f" xpos {center_x}")
        nuke_lines.append(f" ypos {contact_ypos + 20}")
        nuke_lines.append(" name ContactSheet1")
        nuke_lines.append("}")


        print("🧪 Writing Nuke script to:", nuke_script_path)

        with open(nuke_script_path, "w") as f:
            f.write("\n".join(nuke_lines))

        print(Fore.GREEN + f"✅ Nuke script created: {nuke_script_path}" + Style.RESET_ALL)
    else:
        print("Skipping Nuke script creation.")