import os
import logging
import shutil

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s: %(message)s")


def rename_merged_files(id_map: dict, ext: str = '.xlsx', overwrite: bool = False, input_dir: str = 'merged'):
    for fname in os.listdir(input_dir):
        if fname.lower().endswith(ext) and fname.startswith("merged_"):
            identifier = fname[len("merged_"):].split('.')[0]
            district = id_map.get(identifier)
            if not district:
                logging.warning(
                    f"No district for identifier '{identifier}' in {fname}. Skipping.")
                continue
            new_name = f"{district}{ext}"
            src = os.path.join(input_dir, fname)
            dst = os.path.join(input_dir, new_name)
            if os.path.exists(dst):
                if overwrite:
                    logging.warning(f"{dst} exists, overwriting.")
                else:
                    logging.warning(f"{dst} already exists, skipping.")
                    continue
            try:
                shutil.move(src, dst)
                logging.info(f"Renamed {fname} -> {new_name}")
            except Exception as e:
                logging.error(f"Failed to rename {fname}: {e}")
