from shutil import copy
from PyPDF2 import PdfFileWriter, PdfFileReader
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

class ReportWriter:
    '''
    Contains report calculated and files with report template and report output

    Args:
        report_calculated (dict): Calculated report including figures to be written to output file
        report_template_path (str): Path to file with report template
        report_output_path (str): Path to file with report output
    '''

    def __init__(self, report_calculated, report_template_path, report_output_path):
        self.report_calculated = report_calculated
        self.report_template_path = report_template_path
        self.report_output_path = report_output_path

    def write_report_by_watermark(self):
        '''
        Reads given report template, creates output data based on give report calculated, writes it to
        report output and saves it to file in given path location.
        '''
        # Files manipulation and opening ###############################################################################
        copy(self.report_template_path, self.report_output_path)

        self.report_template_file = open(self.report_template_path, "rb")
        self.report_output_file = open(self.report_output_path, "wb")

        # Creating output ##############################################################################################
        self.report_template = PdfFileReader(self.report_template_file)
        self.no_pages_template = self.report_template._get_num_pages()

        self.output = PdfFileWriter()

        # Iterating through pages of report template
        for i in range(self.no_pages_template):
            # Creating packet for actual page
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFillColorRGB(0, 0, 0)
            can.setFont("Times-Roman", 14)

            # Getting sections and its rows
            for section in self.report_calculated:
                rows = self.report_calculated[section]
                for columns in rows.values():
                    for column in columns.values():
                        if column["method"] == "info":
                            continue

                        page_no_template = column["page"] - 1  # Pages in config file starting 1, not 0

                        if i == page_no_template:
                            # Drawing figure
                            figure = column["figure"]
                            figure_string = f"{figure:,}".replace(',', ' ')
                            can.drawRightString(column["x_position"], column["y_position"], figure_string)

            # Creating page output
            can.save()
            packet.seek(0)

            page_output = self.report_template.pages[i]

            if packet.getbuffer().nbytes > 0:
                report_input = PdfFileReader(packet)
                page_input = report_input.pages[0]
                page_output.merge_page(page_input)

            # Adding page output to output object
            self.output.add_page(page_output)

        # Writing output object to output file #########################################################################
        self.output.write(self.report_output_file)

        # File closing #################################################################################################
        self.report_template_file.close()
        self.report_output_file.close()
