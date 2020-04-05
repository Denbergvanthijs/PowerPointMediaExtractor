import os
import shutil
import tempfile
import zipfile
from os import path
from typing import Tuple


def ppmextractor(fp: str, fp_out: str = os.getcwd() + "/media", filter_ext: Tuple[str] = []) -> str:
    """Main function.

    :param fp: Filepath of the Powerpoint-file.
    :type  fp: str
    :param fp_out: Filepath of the output-folder
    :type  fp_out: str, defaults to `os.getcwd() + '/media'`
    :param filter_ext: Tuple of file-extensions to extract
    :type  filter_ext: Tuple[str], defaults to `[]`

    :returns: Filepath to the folder of the extracted files
    :rtype: str
    """
    assert path.isfile(fp), "File at `fp` does not exist."
    assert not path.isdir(fp_out), "Folder should not exist"

    with tempfile.TemporaryDirectory() as tmp_dir:
        filename = path.splitext(path.basename(fp))[0]  # Name of file, without path or file-extension
        filename_zip = filename + ".zip"  # Add .zip file-extension
        tmp_fp = path.join(tmp_dir, filename_zip)  # Full path of zip-file

        shutil.copyfile(fp, tmp_fp)  # Copy file from current location to temporary folder

        with zipfile.ZipFile(tmp_fp, "r") as archive:
            files = archive.namelist()
            assert "ppt/media" in {path.dirname(f) for f in files}, "Zip-file does not contain an media-folder"

            for f in files:  # List of all files
                if f.startswith("ppt/media"):  # If file is in media folder
                    if filter_ext:
                        if f.endswith(filter_ext):  # If file-extention in `filter_ext`
                            archive.extract(f, tmp_dir)
                    else:
                        archive.extract(f, tmp_dir)

        tmp_media_dir = tmp_dir + "/ppt/media"
        shutil.copytree(tmp_media_dir, fp_out)

        return os.listdir(fp_out)


if __name__ == "__main__":
    # output = ppmextractor("./data/test_old.ppt")
    # output = ppmextractor("./data/test_s.ppsx", filter_ext=(".png", ".jpeg"))
    output = ppmextractor("./data/test_x.pptx", filter_ext=(".m4a"))
    print(output)
