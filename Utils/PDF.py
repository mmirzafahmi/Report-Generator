from PyPDF2 import PdfFileWriter, PdfFileReader
from pdfrw import PdfReader, PdfWriter, PageMerge


def Concat(input_files, output):
    input_streams = []
    try:
        # First open all the files, then produce the output file, and
        # finally close the input files. This is necessary because
        # the data isn't read from the input files until the write
        # operation. Thanks to
        # https://stackoverflow.com/questions/6773631/problem-with-closing-python-pypdf-writing-getting-a-valueerror-i-o-operation/6773733#6773733
        for input_file in input_files:
            input_streams.append(open(input_file, 'rb'))
        writer = PdfFileWriter()
        for reader in map(PdfFileReader, input_streams):
            for n in range(reader.getNumPages()):
                writer.addPage(reader.getPage(n))
        with open(output, 'wb') as out:
            writer.write(out)

    finally:
        for f in input_streams:
            f.close()


def Add_Watermark(input_file, output_file, watermark_file):

    # define the reader and writer objects
    reader_input = PdfReader(input_file)
    writer_output = PdfWriter()
    watermark_input = PdfReader(watermark_file)
    watermark = watermark_input.pages[8]

    # go through the pages one after the next
    for current_page in range(len(reader_input.pages)):

        merger = PageMerge(reader_input.pages[current_page])
        merger.add(watermark).render()

    # write the modified content to disk
    writer_output.write(output_file, reader_input)


def Add_Title_Page(input_file, title_file, output_file, page):

    # define the reader and writer objects
    reader_input = PdfReader(input_file)
    writer_output = PdfWriter()
    watermark_input = PdfReader(title_file)
    watermark = watermark_input.pages[page]

    writer_output.addpage(watermark)
    writer_output.addpages(reader_input.pages)

    writer_output.write(output_file)
