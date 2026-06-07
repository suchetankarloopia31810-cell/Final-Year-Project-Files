"""
Convert HEIC images (mislabeled .jpeg) to real JPEG for Word insertion.
Outputs to /projects/sandbox/fyp_report/photos/ with clean names.
"""
import os
from PIL import Image, ImageOps
import pillow_heif

pillow_heif.register_heif_opener()

SRC = '/projects/sandbox/Final-Year-Project-Files'
DST = '/projects/sandbox/fyp_report/photos'
os.makedirs(DST, exist_ok=True)

# Map source filename -> clean output name
MAPPING = {
    'Fine_size_grinded_biomass.jpeg'                  : 'fine_biomass.jpg',
    'Coarse_size_grinded_biomass.JPEG'                : 'coarse_biomass.jpg',
    'Coarse_Sample.jpeg'                              : 'coarse_sample.jpg',
    'Sample_1_without_binder.jpeg'                    : 'sample1_nobinder.jpg',
    'Sample_2_With_90-10_ratio.jpeg'                  : 'sample2_9010.jpg',
    'Sample_3_with_70-30_ratio.jpeg'                  : 'sample3_7030.jpg',
    'Cylinder_sample_1_Without_Binder.jpeg'           : 'cyl1_nobinder.jpg',
    'Cylinder_sample_2_With_90-10 Ratio Binder.jpeg'  : 'cyl2_9010.jpg',
    'Cylinder_sample_3_With_70-30 ratio Binder.jpeg'  : 'cyl3_7030.jpg',
    'Heating_plate_for_corn_starch_gelatinisation.jpeg': 'starch_gelatinisation.jpg',
    'Thermal_property_analyser.jpeg'                  : 'thermal_analyser.jpg',
    'without_Binder_biocomposte_thermal test.jpeg'    : 'thermal_test_nobinder.jpg',
    '90_10_biocomposte_thermal test.jpeg'             : 'thermal_test_9010.jpg',
    '70_30_biocomposte_thermal test.jpeg'             : 'thermal_test_7030.jpg',
    'Mesh Machine.jpeg'                               : 'mesh_machine.jpg',
    'UCT_test_Rig.jpeg'                               : 'uct_rig.jpg',   # actually PNG
}

MAX_DIM = 1600   # downscale large photos for reasonable docx size

for src_name, dst_name in MAPPING.items():
    src_path = os.path.join(SRC, src_name)
    dst_path = os.path.join(DST, dst_name)
    if not os.path.exists(src_path):
        print(f"MISSING: {src_name}")
        continue
    try:
        img = Image.open(src_path)
        img = ImageOps.exif_transpose(img)   # honour rotation
        img = img.convert('RGB')
        # downscale
        w, h = img.size
        if max(w, h) > MAX_DIM:
            scale = MAX_DIM / max(w, h)
            img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
        img.save(dst_path, 'JPEG', quality=88, optimize=True)
        print(f"OK: {dst_name}  ({img.size[0]}x{img.size[1]})")
    except Exception as e:
        print(f"FAIL {src_name}: {e}")

print("\nConversion complete. Files in", DST)
