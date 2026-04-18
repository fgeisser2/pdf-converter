import os
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
