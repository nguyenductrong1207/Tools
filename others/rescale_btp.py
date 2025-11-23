import os
from PIL import Image, UnidentifiedImageError

def rescale_images_in_subfolders(root_folder, output_root):
    # Duyệt qua từng thư mục con trong thư mục gốc
    for subdir, _, files in os.walk(root_folder):
        # Tạo thư mục đích tương ứng trong thư mục output
        relative_path = os.path.relpath(subdir, root_folder)
        output_folder = os.path.join(output_root, relative_path)
        os.makedirs(output_folder, exist_ok=True)

        # Biến đếm số thứ tự
        file_count = 0

        # Xử lý từng file trong thư mục con
        for filename in files:
            input_path = os.path.join(subdir, filename)
            output_path = os.path.join(output_folder, filename)
            
            try:
                # Kiểm tra nếu là file ảnh
                if os.path.isfile(input_path) and filename.lower().endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif', 'tiff')):
                    file_size = os.path.getsize(input_path)  # Lấy kích thước file (bytes)

                    # Tăng số thứ tự
                    file_count += 1
                    
                    # Mở ảnh bằng Pillow
                    try:
                        with Image.open(input_path) as img:   
                            # Kiểm tra kích thước file và rescale nếu cần
                            if file_size > 220 * 1024:  # 200 KB
                                # Rescale 50% kích thước
                                new_size = (img.width // 2, img.height // 2)
                                img = img.resize(new_size, Image.LANCZOS)
                                print(f"{file_count}. Rescaled: {filename} to {new_size} in {relative_path}")
                            else:
                                print(f"{file_count}. Copied: {filename} without rescaling in {relative_path}")

                            # Lưu ảnh vào thư mục mới
                            img.save(output_path)
                    except UnidentifiedImageError:
                        # Xử lý lỗi nếu không thể xác định file là ảnh
                        print(f"Không thể xác định file là ảnh: {input_path}")
                    except IOError as e:
                        print("This img can not open", img, "\n Error:",e)
            except Exception as e:
                print("This img can not read or rescale", img, "\n Error:",e)

# Ví dụ sử dụng

root_folder= "C:\\Users\\MSI\\Downloads\\img\\btp_gio_ga"
output_root = "C:\\Users\\MSI\\Downloads\\img\\btp_gio_ga\\resize"

rescale_images_in_subfolders(root_folder, output_root)
