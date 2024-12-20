import cv2
import pandas as pd
import os
import numpy as np

def extract_color_components(image_path, rows=300, cols=300, output_excel=r"D:\Kuliah\SEMESTER5\pepol\nilai.xlsx"):
    image = cv2.imread(image_path)
    if image is None:
        print(f"Error: Gambar tidak ditemukan atau tidak terbaca di jalur {image_path}")
        return
    else:
        print(f"Gambar berhasil dimuat: {image_path}")

    # Resize gambar
    resized_image = cv2.resize(image, (cols, rows))

    # Ekstrak komponen warna
    red_component = resized_image[:, :, 2]
    green_component = resized_image[:, :, 1]
    blue_component = resized_image[:, :, 0]

    # Grayscale
    gray_image = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)

    # Menghitung fitur kontras dan homogenitas
    contrast = gray_image.std()
    homogenitas = 1 / (1 + gray_image.var())

    # Threshold biner: Ubah 255 menjadi 1
    _, binary_thresh = cv2.threshold(gray_image, 127, 1, cv2.THRESH_BINARY)
    contours, _ = cv2.findContours(binary_thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Menghitung area dan keliling kontur
    area = sum(cv2.contourArea(cnt) for cnt in contours)
    perimeter = sum(cv2.arcLength(cnt, True) for cnt in contours)

    # Deteksi tepi menggunakan Canny
    edges = cv2.Canny(gray_image, 100, 200)
    edges_binary = np.where(edges > 0, 1, 0)

    # Membuat DataFrame untuk setiap komponen
    columns = [f"Kolom {j+1}" for j in range(cols)]
    red_df = pd.DataFrame(red_component, columns=columns)
    green_df = pd.DataFrame(green_component, columns=columns)
    blue_df = pd.DataFrame(blue_component, columns=columns)
    gray_df = pd.DataFrame(gray_image, columns=columns)
    binary_df = pd.DataFrame(binary_thresh, columns=columns)
    edges_df = pd.DataFrame(edges_binary, columns=columns)

    # Simpan data ke file Excel
    try:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # Menyimpan komponen warna dan gambar hasil pemrosesan
            red_df.to_excel(writer, sheet_name='Komponen R (Red)', index=False)
            green_df.to_excel(writer, sheet_name='Komponen G (Green)', index=False)
            blue_df.to_excel(writer, sheet_name='Komponen B (Blue)', index=False)
            gray_df.to_excel(writer, sheet_name='Grayscale', index=False)
            binary_df.to_excel(writer, sheet_name='Binary Image', index=False)
            edges_df.to_excel(writer, sheet_name='Edge Detection', index=False)

            # Menyimpan fitur tambahan di sheet terpisah
            pd.DataFrame({
                'Properti': ['Kontras'],
                'Rumus': ['Standar deviasi dari matriks grayscale'],
                'Nilai': [contrast]
            }).to_excel(writer, sheet_name='Kontras', index=False)

            pd.DataFrame({
                'Properti': ['Homogenitas'],
                'Rumus': ['1 / (1 + variansi matriks grayscale)'],
                'Nilai': [homogenitas]
            }).to_excel(writer, sheet_name='Homogenitas', index=False)

            pd.DataFrame({
                'Properti': ['Area'],
                'Rumus': ['Jumlah area dari semua kontur (piksel putih pada gambar biner)'],
                'Nilai': [area]
            }).to_excel(writer, sheet_name='Area', index=False)

            pd.DataFrame({
                'Properti': ['Keliling'],
                'Rumus': ['Jumlah panjang perimeter dari semua kontur pada gambar biner'],
                'Nilai': [perimeter]
            }).to_excel(writer, sheet_name='Keliling', index=False)

        print(f"File Excel berhasil dibuat: {output_excel}")

    except Exception as e:
        print(f"Error saat mencoba menyimpan file Excel: {e}")

def check_file_existence(output_excel):
    if os.path.exists(output_excel):
        print(f"File berhasil dibuat di {output_excel}")
    else:
        print(f"File tidak ditemukan di {output_excel}")

# Ubah path gambar dan output sesuai kebutuhan
image_path = r"D:\Kuliah\SEMESTER5\pepol\gambar5.png"
output_excel = r"D:\Kuliah\SEMESTER5\pepol\nilai.xlsx"

extract_color_components(image_path, rows=300, cols=300, output_excel=output_excel)
check_file_existence(output_excel)
