#!/bin/python

import os
import math
import tempfile
import subprocess
import threading
import pickle
import shutil
from pathlib import Path
import argparse
import logging
import platform


logger = logging.getLogger(__name__)


PLATFORM = platform.system()
if PLATFORM == "Windows":
    EXECUTABLE = Path("C:/program files/LibreOffice/program/soffice.com")
elif PLATFORM == "Linux":
    try:
        EXECUTABLE = Path(shutil.which("libreoffice"))
    except TypeError:
        Exception("Libreoffice nicht vorhanden!\n Bitte Libreoffe (64bit) installieren")
else:
    Exception("Platform not supported!")


MACRO = r"""import os
import uno
import pickle
import logging
from com.sun.star.beans import PropertyValue

logger = logging.getLogger(f"LO-Converter{os.environ.get('threadId')}")


class Converter:
    def __init__(self, thread_id):
        self.thread_id = thread_id
        self.uno_args_load = (
            self.prop("Minimized", True),
            self.prop("ReadOnly", True),
        )
        self.uno_args_save = (
            self.prop("FilterName", "writer_pdf_Export"),
            self.prop("Overwrite", True),
        )
        self.logfile = os.environ.get("logfile")

    def prop(self, name, value):
        prop = PropertyValue()
        prop.Name = name
        prop.Value = value
        return prop

    def convert(self, files: list[str]):
        desktop = XSCRIPTCONTEXT.getDesktop()
        for f in files:
            logger.info(f"converting: {str(f)}")
            root, ext = os.path.splitext(f)
            try:
                fileUrl = uno.systemPathToFileUrl(str(f))
            except Exception as e:
                logger.error(e)

            try:
                document = desktop.loadComponentFromURL(fileUrl, "_default", 0, self.uno_args_load)
            except Exception as e:
                logger.error(e)

            newpath = f"{root}.pdf"
            fileUrl = uno.systemPathToFileUrl(os.path.realpath(newpath))
            try:
                document.storeToURL(fileUrl, self.uno_args_save)
                document.close(True)
                logger.info(newpath)
                # self.logger(newpath)
            except Exception:
                logger.error(f"saving {newpath} failed!")
                # self.logger(f"error: saving {newpath}")


def main():
    thread_id: str = os.environ.get("threadId")
    tmpfile: str = os.environ.get("tmpfile")
    logging.basicConfig(filename=os.environ.get("logfile"), level=logging.INFO)

    with open(tmpfile, "rb") as f:
        files: list[str]= pickle.load(f)

    with open(os.environ.get("logfile"), "a") as log_file:
        log_file.write(str(files) + "\n")

    Converter(thread_id).convert(files)
"""


def lo_profile_path():
    if PLATFORM == "Windows":
        return Path(os.environ.get("APPDATA"), "LibreOffice", "4", "user")
    if PLATFORM == "Linux":
        return Path("~/.config/libreoffice/4/user").expanduser()

def write_macro(filename):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    with open(filename, "w") as macro_file:
        macro_file.write(MACRO)

def start_convert(files: list[str], id: int, logfile: Path):
    with tempfile.TemporaryDirectory() as tmpdir:
        shutil.copytree(lo_profile_path(), Path(tmpdir, "user"))
        t_file = Path(tmpdir, "t_file")
        write_macro(Path(tmpdir, "user", "Scripts", "python", "macro.py"))

        with open(t_file, "wb") as tmpfile:
            pickle.dump(files, tmpfile)
            tmpfile.close()
            cmd = [
                EXECUTABLE,
                "vnd.sun.star.script:macro.py$main?language=Python&location=user",
                "--nolockcheck",
                "--headless",
                f"--pidfile={Path(tmpdir, 'pidfile').as_uri()}",
                # f"-env:UserInstallation={Path(tmpdir).as_uri()}",
            ]
            envs = os.environ | {
                "tmpfile": f"{tmpfile.name}",
                "logfile": f"{logfile}",
                "UserInstallation": f"{Path(tmpdir).as_uri()}",
                "threadId": f"{id}",
            }
            print(f"Starting LO-Converter {id}")
            subprocess.run(args=cmd, env=envs)
            print(f"LO-Converter {id} finished!")


def main():
    parser = argparse.ArgumentParser(
                        prog='PDF-Converter',
                        description='Converts files to pdf at the file-location')

    parser.add_argument("--path", default=Path("."), type=Path)
    parser.add_argument("--threads", default=os.cpu_count(), type=int)
    parser.add_argument("--logfile", default=Path(__file__).parent / "converter.log", type=Path)
    parser.add_argument("--extensions", default="docx,doc,rtf,txt")
    args = parser.parse_args()

    max_threads: int = args.threads
    conv_path: Path = args.path.expanduser()
    logfile: Path = args.logfile.expanduser()
    logging.basicConfig(filename=logfile, level=logging.INFO)
    extensions = [f".{e}" for e in args.extensions.split(",")]

    try:
        if not EXECUTABLE.is_file():
            Exception("Libreoffice not found! Install Libreoffice (64bit)")
    except TypeError:
        Exception("Libreoffice not found! Install Libreoffice (64bit)")

    all_files = [Path(p,f).resolve() for p, d, files in os.walk(str(conv_path)) for f in files if Path(f).suffix in extensions]
    print(f"converting {len(all_files)} files")

    num_threads = min(int(len(all_files) / 25), max_threads)

    split_files: list[list[Path]] = []
    for t in range(num_threads):
        leng = math.ceil(len(all_files) / num_threads )
        split_files.append(list(all_files[t*leng : min((t+1)*leng, len(all_files))]))


    logger.info("Files per thread:\n" + "\n".join(map(lambda f: str(f), split_files)))

    threads = []
    for i in range(num_threads):
        threads.append(threading.Thread(target=start_convert, args=(split_files[i], i, logfile)))
        threads[i].start()

    for i in range(num_threads):
        threads[i].join()


if __name__ == "__main__":
    main()
