import os
import shutil
import tempfile
import zipfile
from os import path


def ppmextractor(fp):
    """Main function.

    :param fp: Filepath of the Powerpoint-file.
    :type  fp: str

    :returns: Filepath to the folder of the extracted files
    :rtype: str
    """
    assert path.isfile(fp), "File at `fp` does not exist."

    with tempfile.TemporaryDirectory() as tmp_dir:
        filename = path.splitext(path.basename(fp))[0]  # Name of file, without path or file-extension
        filename_zip = filename + ".zip"  # Add .zip file-extension
        tmp_fp = path.join(tmp_dir, filename_zip)  # Full path of zip-file

        shutil.copyfile(fp, tmp_fp)  # Copy file from current location to temporary folder

        with zipfile.ZipFile(tmp_fp, "r") as archive:
            for file in archive.namelist():  # List al files
                if file.startswith("ppt/media"):  # Extract only files in the media-folder
                    archive.extract(file, tmp_dir)

        tmp_media_dir = tmp_dir + "/ppt/media"
        output_dir = os.getcwd() + "/media"

        shutil.rmtree(output_dir)  # Dangerous function!
        shutil.copytree(tmp_media_dir, output_dir)

        return os.listdir(output_dir)


if __name__ == "__main__":
    output = ppmextractor("./data/test_s.ppsx")
    print(output)
