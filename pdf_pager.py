'''
This command line program merges multiple pages into one, adds page numbers, and [nested] bookmarks.
'''
import argparse
from collections import namedtuple
import io
import logging.handlers
import os
import shutil
import sys
import time

from pdfminer.pdfpage import PDFPage
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
from reportlab.pdfgen import canvas


# Parse command line arguments
parser = argparse.ArgumentParser(description='Merge multiple PDFs into a single PDF.')
parser.add_argument('-i', action='append', help='PDF input path (you can specify multiple i arguments)', default=None)
parser.add_argument('-o', help='Destination for new PDF', default='output.pdf')
parser.add_argument('-m', help='Text that will be appended before the page number', default=None)
parser.add_argument('-t', help='Add total pages to the output (Y/N)', default='Y')
parser.add_argument('-a', help='Append date to end of the csv output? (Y/N)', default='Y')
parser.add_argument('-b', help='Space in points between the bottom of the page and the page number', default=10)
parser.add_argument('-p', help='Page number flag (Y/N).  Y = add page numbers.  This will overwrite -m and -t.', default='Y')
parser.add_argument('-r', help='Bookmarks flag (Y/N).  Y = add bookmarks', default="Y")
args = parser.parse_args()


def rotating_log(error_level):
    FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    formatter = logging.Formatter(FORMAT)
    handler = logging.handlers.TimedRotatingFileHandler('rotating.log', when='m', interval=1, backupCount=1)
    handler.setLevel(logging.DEBUG)  # change to ERROR for production
    handler.setFormatter(formatter)
    logger = logging.getLogger("Rotating Log")
    logger.setLevel(error_level)
    logger.addHandler(handler)
    return logger


def iso_date_str():
    return time.strftime("%Y-%m-%d")


class PdfTask:
    def __init__(self, inputs: list, output_path: str="destination.pdf"
            , page_number_mask: str='Page', total_pages_flag: str="N"
            , append_date_flag: str="Y", bottom_margin: int=10
            , page_numbers_flag: str="Y", bookmarks_flag: str="Y"):
        """


        """
        self.inputs = inputs
        self.page_number_mask = page_number_mask
        self.output_path = output_path
        self.total_pages_flag = total_pages_flag
        self.append_date_flag = append_date_flag
        self.bottom_margin = bottom_margin
        self.page_numbers_flag = page_numbers_flag
        self.bookmarks_flag = bookmarks_flag

        logger.info('Input paths: {}'.format(self.input_paths))

    def run_tasks(self):
        logger.info('Plan of attack: {}'.format(self.steps))
        for step in self.steps:
            step.function(*step.args)
        shutil.copy(src=step.output_path, dst=self.final_output_path)
        for step in self.steps:
            os.remove(step.output_path)
            logger.info('Intermediate files deleted')

    @property
    def steps(self) -> list:
        steps = []
        Step = namedtuple('Step', 'input_path, output_path, function, args')
        if len(self.input_paths) > 1:
            inputs = self.input_paths
            output_path = self.get_pdf_name(self.output_path, prefix='merged')
            step = Step(
                input_path=None
                , output_path=output_path
                , function=self.merge_pdfs
                , args=(inputs, output_path)
            )
            steps.append(step)
        if self.page_numbers_flag == 'Y':
            input_path = steps[-1].output_path or self.input_paths[0]
            output_path = self.get_pdf_name(self.output_path, prefix='numbered')
            step = Step(
                input_path=input_path
                , output_path=output_path
                , function=self.add_page_numbers
                , args=(
                    input_path
                    , output_path
                    , self.page_number_mask
                    , self.total_pages_flag
                    , self.bottom_margin
                )
            )
            steps.append(step)
        if self.bookmarks_flag == 'Y':
            input_path = steps[-1].output_path or self.input_paths[0]
            output_path = self.get_pdf_name(self.output_path, prefix='bookmarked')
            step = Step(
                input_path=input_path
                , output_path=output_path
                , function=self.add_bookmarks
                , args=(input_path, output_path)
            )
            steps.append(step)
    #   replace final step's output path with the final output path
    #     steps[-1]._replace(output_path=self.final_output_path)
        logger.info('Steps after substitution: {}'.format(steps))
        return steps

    @staticmethod
    def get_pdf_name(original_output_path: str, prefix: str):
        return original_output_path[:4] + '_' + prefix + '_' + iso_date_str() + '.pdf'

    @property
    def final_output_path(self):
        if self.append_date_flag == "Y":
            return self.output_path[:-4] + '_' + iso_date_str() + '.pdf'
        else:
            return self.output_path

    @staticmethod
    def get_page_rotations(input_path: str) -> dict:
        """This method loops through input PDFs and returns their rotation.

        :return: list(Tuple)
        """
        page_rotations = {}
        with open(input_path, 'rb') as fp:
            for page_number, page in enumerate(PDFPage.get_pages(fp)):
                page_rotations[page_number] = page.rotate
        logger.info('Page rotations: {}'.format(page_rotations))
        return page_rotations

    def get_page_numbers(self) -> dict:
        """Find the would-be starting page for each input into a merged PDF

        :return: dict(PageNumber())
        """
        page_num = 0
        page_numbers = {}
        for input_path in self.input_paths:
            reader = PdfFileReader(open(input_path, "rb"))
            page_numbers[input_path] = page_num
            page_num += reader.getNumPages()
        logger.info('Section Start Pages: {}'.format(page_numbers))
        return page_numbers

    @property
    def bookmarks(self) -> list:
        """i: (title, pagenum, parent)"""
        page_numbers = self.get_page_numbers()
        Bookmark = namedtuple('Bookmark', 'input_path, bookmark_name, parent_bookmark_name, page_number')
        bookmarks = [
            Bookmark(
                input_path=val.split('|')[0]
                , bookmark_name=val.split('|')[1]
                , parent_bookmark_name=val.split('|')[2]
                , page_number=page_numbers.get(val.split('|')[0])
            )
            for val in self.inputs
            if val.split('|')[0]
        ]
        return bookmarks

    @property
    def input_paths(self) -> list:
        return [val.split('|')[0] for val in self.inputs]

    def add_bookmarks(self, input_path: str, output_path: str) -> None:
        """This method loops through pages in a document and add [nested] bookmarks.

        PyPDF2 PdfFileWriter documentation: https://pythonhosted.org/PyPDF2/PdfFileWriter.html
        """

        merger = PdfFileMerger()
        input_pdf = open(input_path, "rb")
        reader = PdfFileReader(input_pdf)
        total_pages = reader.getNumPages()
        output_pdf = open(output_path, "wb")

        merger.append(fileobj=input_pdf, pages=(0, total_pages))
        logger.info('Bookmarks: {}'.format(self.bookmarks))
        page_numbers = self.get_page_numbers()
        parent_bookmarks = {}
        for val in self.bookmarks:
            if val.parent_bookmark_name:
                if val.parent_bookmark_name not in parent_bookmarks.keys():
                    parent_bookmarks[val.parent_bookmark_name] = merger.addBookmark(
                        title=val.parent_bookmark_name
                        , pagenum=page_numbers.get(val.input_path)
                        , parent=None
                    )
                merger.addBookmark(
                    title=val.bookmark_name
                    , pagenum=val.page_number
                    , parent=parent_bookmarks.get(val.parent_bookmark_name)
                )
            else:
                if val.bookmark_name:
                    merger.addBookmark(
                        title=val.bookmark_name
                        , pagenum=val.page_number
                        , parent=None
                    )
        logger.info('Parent bookmarks: {}'.format(parent_bookmarks))
        merger.write(output_pdf)
        input_pdf.close()
        output_pdf.close()

    def merge_pdfs(self, input_paths: list, output_path: str) -> None:
        merger = PdfFileMerger()
        output = open(output_path, "wb")
        for input_path in sorted(input_paths):
            logger.info('Appending pdf {}...'.format(input_path))
            input_pdf = open(input_path, "rb")
            merger.append(fileobj=input_pdf)

        merger.write(output)
        output.close()
        logger.debug('PDFs merged successfully')

    def add_page_numbers(self, input_path, output_path, mask, total_pages_flag, bottom_margin):
        page_rotations_dict = PdfTask.get_page_rotations(input_path)
        logger.info('Page rotations: {}'.format(page_rotations_dict))
        output = PdfFileWriter()
        input_pdf = open(input_path, "rb")
        reader = PdfFileReader(input_pdf)
        page_ct = reader.getNumPages()

        logger.info('doc info: ' + str(reader.documentInfo))
        logger.info('page layout:' + str(reader.getPageLayout()))
        logger.info('page mode:' + str(reader.getPageMode()))
        logger.info('xmp metadata:' + str(reader.getXmpMetadata()))

        for page_num in range(page_ct):
            #   inspect the current input PDF page
            page = reader.getPage(page_num)
            page_rect = page.mediaBox
            logger.info('page media box: Page {num}: {dim}'.format(num=page_num, dim=page_rect))
            #   dimensions for a letter sized sheet of paper are [0, 0, 612, 792]
            #   72 pt = 1 inch

            page_dimensions = {
                'lower_left':    page_rect.getUpperLeft()
                , 'lower_right': page_rect.getLowerRight()
                , 'upper_left':  page_rect.getUpperLeft()
                , 'upper_right': page_rect.getUpperRight()
            }
            logger.info('Page dimensions: {}'.format(page_dimensions))

            #   create a new PDF containing the page number as a watermark with Reportlab
            txt = str(page_num + 1)

            if mask:
                txt = mask + " " + txt
            if total_pages_flag == 'Y':
                txt = txt + " of " + str(page_ct)

            packet = io.BytesIO()

            page_width = page_rect.getWidth()
            page_height = page_rect.getHeight()

            c = canvas.Canvas(packet, pagesize=(0, 0))
            c.drawString(page_width / 2, bottom_margin, txt)

            c.save()
            packet.seek(0)
            new_pdf = PdfFileReader(packet)
            #   merge new watermark pdf with the original
            wm = new_pdf.getPage(0)
            page_rotation = page_rotations_dict.get(page_num) or 0
            page.mergeRotatedTranslatedPage(
                wm
                , rotation=page_rotation
                , tx=page_width / 2
                , ty=page_height / 2
                , expand=True
            )
            page.scaleTo(page_width, page_height)
            page.compressContentStreams()
            output.addPage(page)

        with open(output_path, "wb") as outputStream:
            output.write(outputStream)

        input_pdf.close()
        logger.debug('Successfully added page numbers to {}'.format(input_path))


if __name__ == '__main__':
    try:
        logger = rotating_log(error_level=logging.DEBUG)
        logger.info('input arguments: {}'.format(args))
        task = PdfTask(inputs=args.i, output_path=args.o, page_number_mask=args.m
            , total_pages_flag=args.t, append_date_flag=args.a, bookmarks_flag=args.r
            , page_numbers_flag=args.p, bottom_margin=args.b)

        #TEST
        # task = PdfTask(
        #     inputs=['input.pdf|child1|parent1', 'input2.pdf|child2|parent1', 'input3.pdf|child3|parent2', 'input.pdf||']
        #     , output_path='destination.pdf'
        #     , bottom_margin=10
        #     , page_number_mask=''
        #     , total_pages_flag="Y"
        #     , append_date_flag="Y"
        #     , page_numbers_flag="Y"
        #     , bookmarks_flag="Y"
        # )
        task.run_tasks()
        os.startfile(task.final_output_path)
        logger.info("Success")
        sys.exit(0)
    except Exception as e:
        logger.error('Error: {}'.format(e))
        sys.exit(1)
