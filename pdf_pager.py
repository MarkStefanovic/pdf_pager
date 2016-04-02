'''
This command line program merges multiple pages into one, adds page numbers, and [nested] bookmarks.
'''
import argparse
import io
import logging.handlers
import os
import sys
import time

from pdfminer.pdfpage import PDFPage

from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter

from reportlab.pdfgen import canvas


# Parse command line arguments
parser = argparse.ArgumentParser(description='Merge multiple PDFs into a single PDF.')
parser.add_argument('-i', action='append', help='PDF input path (you can specify multiple i arguments)', default=['input1.pdf|child1|parent1', 'input2.pdf||', 'input3.pdf|child3|parent2', 'input4.pdf|child4|parent2'])
parser.add_argument('-o', help='Destination for new PDF', default='destination.pdf')
parser.add_argument('-m', help='Text that will be appended before the page number', default=None)
parser.add_argument('-t', help='Add total pages to the output (Y/N)', default='Y')
parser.add_argument('-a', help='Append date to end of the csv output? (Y/N)', default='Y')
parser.add_argument('-b', help='Space in points between the bottom of the page and the page number', default=10)
args = parser.parse_args()

# Set up logging
FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
formatter = logging.Formatter(FORMAT)
handler = logging.handlers.TimedRotatingFileHandler('rotating.log', when='m', interval=1, backupCount=1)
handler.setLevel(logging.INFO)
handler.setFormatter(formatter)
logger = logging.getLogger("Rotating Log")
logger.setLevel(logging.INFO)
logger.addHandler(handler)


def iso_date_str():
    return time.strftime("%Y-%m-%d")


def add_bookmarks(task):
    """Loop through pages in a document and add [nested] bookmarks

    PyPDF2 PdfFileWriter documentation: https://pythonhosted.org/PyPDF2/PdfFileWriter.html
    """

    merger = PdfFileMerger()
    input_pdf = open(task.numbered_pdf_path, "rb")
    reader = PdfFileReader(input_pdf)
    total_pages = reader.getNumPages()
    output_pdf = open(task.final_output_path, "wb")

    merger.append(fileobj=input_pdf, pages=(0, total_pages))

    parent_bookmarks = {}
    child_bookmarks = {}
    for key, val in task.bookmarks.items():
        if val[2] not in parent_bookmarks.keys():
            parent_bookmarks[val[2]] = merger.addBookmark(
                title=val[2]
                , pagenum=task.bookmark_parents.get(val[2])
                , parent=None
            )
        merger.addBookmark(
            title=val[0]
            , pagenum=val[1]
            , parent=parent_bookmarks.get(val[2])
        )
        child_bookmarks[key] = {
            'title': val[0]
            , 'page number': val[1]
            , 'parent bookmark name': val[2]
            , 'parent bookmark object': parent_bookmarks.get(val[2])
        }
    logger.info('Added parent bookmarks: {}'.format(parent_bookmarks))
    logger.info('Added child bookmarks: {}'.format(child_bookmarks))

    merger.write(output_pdf)
    input_pdf.close()
    output_pdf.close()


def merge_pdfs(task):
    merger = PdfFileMerger()
    output = open(task.merged_pdf_path, "wb")
    for i, input_path in sorted(task.input_paths.items()):
        logger.info(str(i) + ': Appending ' + input_path)

        input_pdf = open(input_path, "rb")
        merger.append(fileobj=input_pdf)

    merger.write(output)
    output.close()
    logger.info('PDFs merged successfully')


def add_page_numbers(task):
    output = PdfFileWriter()
    input_pdf = open(task.merged_pdf_path, "rb")
    reader = PdfFileReader(input_pdf)
    page_ct = reader.getNumPages()
    page_rotations_dict = task.page_rotations

    logger.info('doc info: ' + str(reader.documentInfo))
    logger.info('page layout:' + str(reader.getPageLayout()))
    logger.info('page mode:' + str(reader.getPageMode()))
    logger.info('xmp metadata:' + str(reader.getXmpMetadata()))

    for page_num in range(page_ct):
    #   inspect the current input PDF page
        page = reader.getPage(page_num)
        page_rect = page.mediaBox
        logger.info('page media box: {}'.format(page_rect))
        #   dimensions for a letter sized sheet of paper are [0, 0, 612, 792]
        #   72 pt = 1 inch

        page_dimensions = {
            'lower_left': page_rect.getUpperLeft()
            , 'lower_right': page_rect.getLowerRight()
            , 'upper_left': page_rect.getUpperLeft()
            , 'upper_right': page_rect.getUpperRight()
        }
        logger.info('input page dimensions: {lower_left}, {lower_right}, {upper_left}, {upper_right}'.format(**page_dimensions))

    #   create a new PDF containing the page number as a watermark with Reportlab
        txt = str(page_num + 1)
        if args.m:
            txt = args.m + " " + txt
        if args.t == 'Y':
            txt = txt + " of " + str(page_ct)

        packet = io.BytesIO()

        page_width = page.mediaBox.getWidth()
        page_height = page.mediaBox.getHeight()

        c = canvas.Canvas(packet, pagesize=(0, 0))
        c.drawString(page_width/2, args.b, txt)

        c.save()
        packet.seek(0)
        new_pdf = PdfFileReader(packet)

    #   merge new watermark pdf with the original
        wm = new_pdf.getPage(0)

        page.mergeRotatedTranslatedPage(
            wm
            , rotation=page_rotations_dict.get(page_num)
            , tx=page_width/2
            , ty=page_height/2
            , expand=True
        )
        page.scaleTo(page_width, page_height)
        output.addPage(page)

    with open(task.numbered_pdf_path, "wb") as outputStream:
        output.write(outputStream)

    input_pdf.close()
    logger.info('Successfully added page numbers to {}'.format(task.numbered_pdf_path))

class PdfTask:
    def __init__(self, inputs, output_path="destination.pdf", mask=None, total_pages_flag="N", append_date_flag="Y", bottom_margin=10):
        self.inputs = inputs
        self.output_path = output_path
        self.total_pages_flag = total_pages_flag
        self.append_date_flag = append_date_flag
        self.bottom_margin = bottom_margin

    @property
    def merged_pdf_path(self):
        return self.output_path[:4] + '_merged_' + iso_date_str() + '.pdf'

    @property
    def numbered_pdf_path(self):
        return self.output_path[:4] + '_numbered_' + iso_date_str() + '.pdf'

    @property
    def final_output_path(self):
        if self.append_date_flag == "Y":
            return self.output_path[:-4] + '_' + iso_date_str() + '.pdf'
        else:
            return self.output_path

    @property
    def page_rotations(self):
        page_rotations_dict = {}
        fp = open(self.merged_pdf_path, 'rb')
        for i, page in enumerate(PDFPage.get_pages(fp)):
            page_rotations_dict[i] = page.rotate
        logger.info('Page rotations: {}'.format(page_rotations_dict))
        fp.close()
        return page_rotations_dict

    @property
    def page_numbers(self):
        """Find the would-be starting page for each input into a merged PDF

        returns: {i: pagenum}
        """
        page_num = 0
        page_numbers_dict = {}
        for i, input_path in self.input_paths.items():
            reader = PdfFileReader(open(input_path, "rb"))
            page_numbers_dict[i] = page_num
            page_num += reader.getNumPages()
        logger.info('Section Start Pages: {}'.format(page_numbers_dict))
        return page_numbers_dict

    @property
    def bookmarks(self):
        """i: (title, pagenum, parent)"""
        inputs = {key: val for key, val in enumerate(self.inputs)}
        bookmarks = {key: val.split('|')[1] for key, val in inputs.items()}
        bookmark_parents = {key: val.split('|')[2] for key, val in inputs.items()}
        bookmarks_dict = {}
        for key, val in bookmarks.items():
            if val:
                pagenum = self.page_numbers.get(key)
                parent = bookmark_parents.get(key)
                bookmarks_dict[key] = (val, pagenum, parent)
        logger.info('Bookmarks: {}'.format(bookmarks_dict))
        return bookmarks_dict

    @property
    def bookmark_parents(self):
        """returns {title: pagenum}"""
        bookmark_parents_dict = {}
        inputs = {key: val for key, val in enumerate(self.inputs)}
        bookmark_parents = {key: val.split('|')[2] for key, val in inputs.items()}
        for key, val in sorted(bookmark_parents.items()):
            if val not in bookmark_parents_dict.keys() and val:
                pagenum = self.page_numbers.get(key)
                bookmark_parents_dict[val] = pagenum
        logger.info('Bookmark Parents: {}'.format(bookmark_parents_dict))
        return bookmark_parents_dict

    @property
    def input_paths(self):
        inputs = {key: val for key, val in enumerate(self.inputs)}
        input_paths = {key: val.split('|')[0] for key, val in inputs.items()}
        logger.info('Input paths dict: {}'.format(input_paths))
        return input_paths

    def cleanup(self):
        os.remove(self.merged_pdf_path)
        os.remove(self.numbered_pdf_path)
        logger.info('Intermediate files deleted')

if __name__ == '__main__':
    try:
        # task = PdfTask(args.i, args.o, args.m, args.t, args.a, args.b)
        task = PdfTask(['input.pdf|child1|parent1', 'input2.pdf|child2|parent1', 'input3.pdf|child3|parent2', 'input.pdf||'], 'destination.pdf', 10, '', 'Y', 10)
        merge_pdfs(task)
        add_page_numbers(task)
        add_bookmarks(task)
        task.cleanup()
        os.open(task.final_output_path, os.O_RDONLY)
        logger.info("Success")
        sys.exit(0)
    except Exception as e:
        logger.error('Error: {}'.format(e))
        sys.exit(1)