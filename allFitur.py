import cv2
import pandas as pd
import os
import numpy as np

def extract_folder_features(folder_path, rows=300, cols=300, output_excel=r"D:\Kuliah\SEMESTER5\penPol\nilai.xlsx"):
    # Ambil semua file gambar di folder dengan ekstensi yang didukung
    image_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

    if not image_files:
        print("Tidak ada file gambar dalam folder.")
        return

    # List untuk menyimpan data fitur
    all_data = []

    for image_file in image_files:
        image_path = os.path.join(folder_path, image_file)
        image = cv2.imread(image_path)
        
        if image is None:
            print(f"Gambar tidak valid atau tidak bisa dibaca: {image_path}")
            continue
        else:
            print(f"Memproses gambar: {image_file}")

        # Resize gambar
        resized_image = cv2.resize(image, (cols, rows))

        # Ekstrak komponen warna RGB
        red_component = resized_image[:, :, 2]
        green_component = resized_image[:, :, 1]
        blue_component = resized_image[:, :, 0]

        # Rata-rata warna
        avg_red = np.mean(red_component)
        avg_green = np.mean(green_component)
        avg_blue = np.mean(blue_component)

        # Grayscale
        gray_image = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)

        # Menghitung fitur
        contrast = gray_image.std()
        homogenitas = 1 / (1 + gray_image.var())

        # Threshold biner
        _, binary_thresh = cv2.threshold(gray_image, 127, 1, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(binary_thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        area = sum(cv2.contourArea(cnt) for cnt in contours)
        perimeter = sum(cv2.arcLength(cnt, True) for cnt in contours)

        # Edge detection
        edges = cv2.Canny(gray_image, 100, 200)
        edges_binary = np.where(edges > 0, 1, 0)

        # Menambahkan data fitur ke list
        all_data.append({
            'Gambar': image_file,
            'Warna_R': avg_red,
            'Warna_G': avg_green,
            'Warna_B': avg_blue,
            'Kontras': contrast,
            'Homogenitas': homogenitas,
            'Area': area,
            'Keliling': perimeter
        })

    # Membuat DataFrame dari semua data
    df = pd.DataFrame(all_data)

    # Menyimpan DataFrame ke file Excel
    try:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # Simpan fitur gambar utama
            df.to_excel(writer, index=False, sheet_name='Fitur Gambar')

            # Menyimpan fitur tambahan di sheet terpisah
            for feature in ['Kontras', 'Homogenitas', 'Area', 'Keliling']:
                feature_df = pd.DataFrame({
                    'Gambar': df['Gambar'],
                    'Nilai': df[feature]
                })
                feature_df.to_excel(writer, index=False, sheet_name=feature)

        print(f"File Excel berhasil dibuat: {output_excel}")
    except Exception as e:
        print(f"Error saat mencoba menyimpan file Excel: {e}")

# Ubah path folder dan output sesuai kebutuhan
folder_path = r"D:\Kuliah\SEMESTER5\penPol\gambar"
output_excel = r"D:\Kuliah\SEMESTER5\penPol\nilai.xlsx"

extract_folder_features(folder_path, rows=300, cols=300, output_excel=output_excel)
