from os.path import join

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch, mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import PageBreak

style = getSampleStyleSheet()

heading1 = ParagraphStyle(
    'Heading1',
    parent=style['Heading1'],
    alignment=TA_CENTER,
)

heading2 = ParagraphStyle(
    'Heading2',
    parent=style['Heading2'],
    alignment=TA_CENTER
)


class GeneratePDF:
    """Generate PDF"""

    @staticmethod
    def generate_student_marksheet(marksheet_dict, size, path):
        canvas = SimpleDocTemplate(join(path, 'Student Marksheet'), pagesize=A4)

        element = []

        for i in range(size):

            roll = Paragraph("Roll No.: " + str(marksheet_dict[i]['Roll No.']), style['Heading2'])
            name = Paragraph("Name: " + marksheet_dict[i]['Name'], style['Heading2'])

            data = [
                ['Subject', 'Unit', 'Unit_10', 'Oral', 'Term', 'Total', 'Total/2'],
                [
                    'English',
                    str(marksheet_dict[i]['English']['Unit']),
                    str(marksheet_dict[i]['English']['Unit_10']),
                    str(marksheet_dict[i]['English']['Oral']),
                    str(marksheet_dict[i]['English']['Term']),
                    str(marksheet_dict[i]['English']['Total']),
                    str(marksheet_dict[i]['English']['Total/2'])
                ],
                [
                    'Hindi',
                    str(marksheet_dict[i]['Hindi']['Unit']),
                    str(marksheet_dict[i]['Hindi']['Unit_10']),
                    str(marksheet_dict[i]['Hindi']['Oral']),
                    str(marksheet_dict[i]['Hindi']['Term']),
                    str(marksheet_dict[i]['Hindi']['Total']),
                    str(marksheet_dict[i]['Hindi']['Total/2'])
                ],
                [
                    'Odia',
                    str(marksheet_dict[i]['Odia']['Unit']),
                    str(marksheet_dict[i]['Odia']['Unit_10']),
                    str(marksheet_dict[i]['Odia']['Oral']),
                    str(marksheet_dict[i]['Odia']['Term']),
                    str(marksheet_dict[i]['Odia']['Total']),
                    str(marksheet_dict[i]['Odia']['Total/2'])
                ],
                [
                    'Maths',
                    str(marksheet_dict[i]['Maths']['Unit']),
                    str(marksheet_dict[i]['Maths']['Unit_10']),
                    str(marksheet_dict[i]['Maths']['Oral']),
                    str(marksheet_dict[i]['Maths']['Term']),
                    str(marksheet_dict[i]['Maths']['Total']),
                    str(marksheet_dict[i]['Maths']['Total/2'])
                ],
                [
                    'Science',
                    str(marksheet_dict[i]['Science']['Unit']),
                    str(marksheet_dict[i]['Science']['Unit_10']),
                    str(marksheet_dict[i]['Science']['Oral']),
                    str(marksheet_dict[i]['Science']['Term']),
                    str(marksheet_dict[i]['Science']['Total']),
                    str(marksheet_dict[i]['Science']['Total/2'])
                ],
                [
                    'SST',
                    str(marksheet_dict[i]['SST']['Unit']),
                    str(marksheet_dict[i]['SST']['Unit_10']),
                    str(marksheet_dict[i]['SST']['Oral']),
                    str(marksheet_dict[i]['SST']['Term']),
                    str(marksheet_dict[i]['SST']['Total']),
                    str(marksheet_dict[i]['SST']['Total/2'])
                ],
                ['G.K', ' ', ' ', ' ', ' ', marksheet_dict[i]['GK'], ' '],
                ['Computer', ' ', ' ', ' ', ' ', marksheet_dict[i]['Computer'], ' '],
                ['Total', ' ', ' ', ' ', ' ', marksheet_dict[i]['Total Mark'], ' '],
                ['Percent', ' ', ' ', ' ', ' ', marksheet_dict[i]['Percentage'], ' ']
            ]

            table = Table(data, 7 * [1 * inch], 11 * [0.25 * inch])
            table.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black)
            ]))

            element.append(roll)
            element.append(name)
            element.append(table)

            if i % 2 == 0:
                space = Paragraph("     ", style['Heading3'])
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
                element.append(space)
            else:
                element.append(PageBreak()), style['Heading2']
        canvas.build(element)
        print('Done!')

    @staticmethod
    def generate_unit_marksheet(marksheet_dict, size, path, class_name):
        canvas = SimpleDocTemplate(join(path, 'Unit Marksheet'), pagesize=(297 * mm, 210 * mm))

        element = []

        head = Paragraph(class_name + ' - Unit Test', style['Heading2'])

        data = [
            ['Roll No.', 'Name', 'English', 'Hindi', 'Odia', 'Maths', 'Science', 'SST', 'Total Mark', 'Percentage']
        ]
        for i in range(size):
            total = marksheet_dict[i]['English']['Unit'] + marksheet_dict[i]['Hindi']['Unit'] + \
                    marksheet_dict[i]['Odia']['Unit'] + marksheet_dict[i]['Maths']['Unit'] + \
                    marksheet_dict[i]['Science']['Unit'] + marksheet_dict[i]['SST']['Unit']
            percent = '%.2f' % ((total / 180) * 100)

            data.append([
                str(marksheet_dict[i]['Roll No.']),
                marksheet_dict[i]['Name'],
                str(marksheet_dict[i]['English']['Unit']),
                str(marksheet_dict[i]['Hindi']['Unit']),
                str(marksheet_dict[i]['Odia']['Unit']),
                str(marksheet_dict[i]['Maths']['Unit']),
                str(marksheet_dict[i]['Science']['Unit']),
                str(marksheet_dict[i]['SST']['Unit']),
                str(total),
                str(percent)
            ])

        table = Table(data)
        table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black)
        ]))

        element.append(head)
        element.append(table)
        canvas.build(element)
        print('Done!')

    @staticmethod
    def generate_term_marksheet(marksheet_dict, size, path, class_name):
        canvas = SimpleDocTemplate(join(path, 'Term Marksheet'), pagesize=(297 * mm, 210 * mm))

        element = []

        head = Paragraph(class_name + ' - Term Test', style['Heading2'])

        data = [['Roll No.', 'Name',
                 'English', 'Hindi', 'Odia', 'Maths', 'Science', 'SST', 'GK', 'Computer',
                 'Total Mark', 'Percentage']]

        for i in range(size):
            total = marksheet_dict[i]['English']['Term'] + marksheet_dict[i]['Hindi']['Term'] + \
                    marksheet_dict[i]['Odia']['Term'] + marksheet_dict[i]['Maths']['Term'] + \
                    marksheet_dict[i]['Science']['Term'] + marksheet_dict[i]['SST']['Term'] + \
                    marksheet_dict[i]['GK'] + marksheet_dict[i]['Computer']
            percent = '%.2f' % (total / 7)

            data.append([
                str(marksheet_dict[i]['Roll No.']),
                marksheet_dict[i]['Name'],
                str(marksheet_dict[i]['English']['Term']),
                str(marksheet_dict[i]['Hindi']['Term']),
                str(marksheet_dict[i]['Odia']['Term']),
                str(marksheet_dict[i]['Maths']['Term']),
                str(marksheet_dict[i]['Science']['Term']),
                str(marksheet_dict[i]['SST']['Term']),
                str(marksheet_dict[i]['GK']),
                str(marksheet_dict[i]['Computer']),
                str(total),
                str(percent)
            ])

        table = Table(data)
        table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black)
        ]))

        element.append(head)
        element.append(table)
        canvas.build(element)
        print('Done!')

    @staticmethod
    def generate_unit_marksheet_standalone(marksheet_dict, size, path, class_name):
        canvas = SimpleDocTemplate(
            join(path, 'Unit Marksheet'),
            pagesize=(297 * mm, 210 * mm),
            topMargin=15*mm,
            bottomMargin=15*mm
        )

        element = list()

        element.append(Paragraph('KHARIAR PUBLIC SCHOOL, KHARIAR', heading1))
        element.append(Paragraph('Result of Unit Test - 1 (2020-21)', heading2))
        element.append(Paragraph(class_name, heading2))

        data = [
            ['Roll No.', 'Name',
             'English', 'Hindi', 'Odia', 'Maths', 'Science', 'SST',
             'Mark Obtained', 'Total Mark', 'Percentage', 'Grade']
        ]

        for i in range(size):
            data.append([
                str(marksheet_dict[i]['Roll No.']),
                marksheet_dict[i]['Name'],
                str(marksheet_dict[i]['English']),
                str(marksheet_dict[i]['Hindi']),
                str(marksheet_dict[i]['Odia']),
                str(marksheet_dict[i]['Maths']),
                str(marksheet_dict[i]['Science']),
                str(marksheet_dict[i]['SST']),
                str(marksheet_dict[i]['Mark Obtained']),
                str(marksheet_dict[i]['Total Mark']),
                str(marksheet_dict[i]['Percentage']),
                str(marksheet_dict[i]['Grade'])
            ])

        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black)
        ]))

        element.append(table)
        canvas.build(element)
        print('Done!')

    @staticmethod
    def generate_term_marksheet_standalone(marksheet_dict, size, path, class_name):
        canvas = SimpleDocTemplate(
            join(path, 'Term Marksheet'),
            pagesize=(297 * mm, 210 * mm),
            topMargin=15 * mm,
            bottomMargin=15 * mm
        )

        element = list()

        element.append(Paragraph('KHARIAR PUBLIC SCHOOL, KHARIAR', heading1))
        element.append(Paragraph('Result of Term Test - 1 (2020-21)', heading2))
        element.append(Paragraph(class_name, heading2))

        data = [
            ['Roll No.', 'Name',
             'English', 'Hindi', 'Odia', 'Maths', 'Science', 'SST', 'GK', 'Computer',
             'Mark Obtained', 'Total Mark', 'Percentage', 'Grade']
        ]

        for i in range(size):
            data.append([
                str(marksheet_dict[i]['Roll No.']),
                marksheet_dict[i]['Name'],
                str(marksheet_dict[i]['English']),
                str(marksheet_dict[i]['Hindi']),
                str(marksheet_dict[i]['Odia']),
                str(marksheet_dict[i]['Maths']),
                str(marksheet_dict[i]['Science']),
                str(marksheet_dict[i]['SST']),
                str(marksheet_dict[i]['GK']),
                str(marksheet_dict[i]['Computer']),
                str(marksheet_dict[i]['Mark Obtained']),
                str(marksheet_dict[i]['Total Mark']),
                str(marksheet_dict[i]['Percentage']),
                str(marksheet_dict[i]['Grade'])
            ])

        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black)
        ]))

        element.append(table)
        canvas.build(element)
        print('Done!')
