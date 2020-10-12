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

heading3 = ParagraphStyle(
    'Heading3',
    parent=style['Heading3'],
    alignment=TA_CENTER
)


class GeneratePDF:
    """Generate PDF"""

    @staticmethod
    def generate_student_marksheet(marksheet_dict, size, path):
        canvas = SimpleDocTemplate(join(path, 'Student Marksheet'), pagesize=A4)

        element = list()

        for i in range(size):

            roll = Paragraph("Roll No.: " + str(marksheet_dict[i]['Roll No.']), style['Heading3'])
            name = Paragraph("Name: " + marksheet_dict[i]['Name'], style['Heading3'])

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
                ['Mark Obtained', ' ', ' ', ' ', ' ', marksheet_dict[i]['Mark Obtained'], ' '],
                ['Total Mark', ' ', ' ', ' ', ' ', marksheet_dict[i]['Total Mark'], ' '],
                ['Percent', ' ', ' ', ' ', ' ', marksheet_dict[i]['Percentage'], ' '],
                ['Grade', ' ', ' ', ' ', ' ', marksheet_dict[i]['Grade'], ' ']
            ]

            table = Table(
                data,
                colWidths=[1.25*inch, 0.75*inch, 0.75*inch, 0.75*inch, 0.75*inch, 0.75*inch, 0.75*inch]
            )
            table.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
                ('ALIGN', (5, 12), (5, 12), 'CENTER')
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
            else:
                element.append(PageBreak())
        canvas.build(element)
        print('Student Marksheet Created')

    @staticmethod
    def generate_unit_marksheet(marksheet_dict, size, path, number, class_name):
        canvas = SimpleDocTemplate(
            join(path, 'Unit Marksheet'),
            pagesize=(297 * mm, 210 * mm),
            topMargin=15 * mm,
            bottomMargin=15 * mm
        )

        element = list()

        element.append(Paragraph('KHARIAR PUBLIC SCHOOL, KHARIAR', heading1))
        element.append(Paragraph('Result of Unit Test - ' + number, heading3))
        element.append(Paragraph(class_name, heading3))
        element.append(Paragraph('', heading3))

        data = [
            ['Roll No.', 'Name',
             'English', 'Hindi', 'Odia', 'Maths', 'Science', 'SST',
             'Mark Obtained', 'Total Mark', 'Percentage', 'Grade']
        ]
        for i in range(size):
            mark_obtained = marksheet_dict[i]['English']['Unit'] + marksheet_dict[i]['Hindi']['Unit'] + \
                            marksheet_dict[i]['Odia']['Unit'] + marksheet_dict[i]['Maths']['Unit'] + \
                            marksheet_dict[i]['Science']['Unit'] + marksheet_dict[i]['SST']['Unit']
            total_mark = 180
            percent = float('%.2f' % (mark_obtained / 1.8))
            if percent >= 91:
                grade = 'A1'
            elif percent >= 81:
                grade = 'A2'
            elif percent >= 71:
                grade = 'B1'
            elif percent >= 61:
                grade = 'B2'
            elif percent >= 51:
                grade = 'C1'
            elif percent >= 41:
                grade = 'C2'
            elif percent >= 33:
                grade = 'D'
            else:
                grade = 'E'

            data.append([
                str(marksheet_dict[i]['Roll No.']),
                marksheet_dict[i]['Name'],
                str(marksheet_dict[i]['English']['Unit']),
                str(marksheet_dict[i]['Hindi']['Unit']),
                str(marksheet_dict[i]['Odia']['Unit']),
                str(marksheet_dict[i]['Maths']['Unit']),
                str(marksheet_dict[i]['Science']['Unit']),
                str(marksheet_dict[i]['SST']['Unit']),
                str(mark_obtained),
                str(total_mark),
                str(percent),
                grade
            ])

        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (2, 0), (-2, -1), 'RIGHT'),
            ('ALIGN', (11, 1), (11, -1), 'CENTER')
        ]))

        element.append(table)
        canvas.build(element)
        print('Unit Marksheet Created')

    @staticmethod
    def generate_term_marksheet(marksheet_dict, size, path, number, class_name):
        canvas = SimpleDocTemplate(
            join(path, 'Term Marksheet'),
            pagesize=(297 * mm, 210 * mm),
            topMargin=15 * mm,
            bottomMargin=15 * mm
        )

        element = list()

        element.append(Paragraph('KHARIAR PUBLIC SCHOOL, KHARIAR', heading1))
        element.append(Paragraph('Result of Term Test - ' + number, heading3))
        element.append(Paragraph(class_name, heading3))
        element.append(Paragraph('', heading3))

        data = [
            ['Roll No.', 'Name',
             'English', 'Hindi', 'Odia', 'Maths', 'Science', 'SST', 'GK', 'Computer',
             'Mark Obtained', 'Total Mark', 'Percentage', 'Grade']
        ]

        for i in range(size):
            mark_obtained = marksheet_dict[i]['English']['Term'] + marksheet_dict[i]['Hindi']['Term'] + \
                            marksheet_dict[i]['Odia']['Term'] + marksheet_dict[i]['Maths']['Term'] + \
                            marksheet_dict[i]['Science']['Term'] + marksheet_dict[i]['SST']['Term'] + \
                            marksheet_dict[i]['GK'] + marksheet_dict[i]['Computer']
            total_mark = 580
            percent = float('%.2f' % (mark_obtained / 7))
            if percent >= 91:
                grade = 'A1'
            elif percent >= 81:
                grade = 'A2'
            elif percent >= 71:
                grade = 'B1'
            elif percent >= 61:
                grade = 'B2'
            elif percent >= 51:
                grade = 'C1'
            elif percent >= 41:
                grade = 'C2'
            elif percent >= 33:
                grade = 'D'
            else:
                grade = 'E'

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
                str(mark_obtained),
                str(total_mark),
                str(percent),
                grade
            ])

        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (2, 0), (-2, -1), 'RIGHT'),
            ('ALIGN', (13, 1), (13, -1), 'CENTER')
        ]))

        element.append(table)
        canvas.build(element)
        print('Term Marksheet Created')

    @staticmethod
    def generate_unit_marksheet_standalone(marksheet_dict, size, path, number, class_name):
        canvas = SimpleDocTemplate(
            join(path, 'Unit Marksheet'),
            pagesize=(297 * mm, 210 * mm),
            topMargin=15 * mm,
            bottomMargin=15 * mm
        )

        element = list()

        element.append(Paragraph('KHARIAR PUBLIC SCHOOL, KHARIAR', heading1))
        element.append(Paragraph('Result of Unit Test - ' + number, heading3))
        element.append(Paragraph(class_name, heading3))
        element.append(Paragraph('', heading3))

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
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (2, 0), (-2, -1), 'RIGHT'),
            ('ALIGN', (11, 1), (11, -1), 'CENTER')
        ]))

        element.append(table)
        canvas.build(element)
        print('Unit Marksheet Created')

    @staticmethod
    def generate_term_marksheet_standalone(marksheet_dict, size, path, number, class_name):
        canvas = SimpleDocTemplate(
            join(path, 'Term Marksheet'),
            pagesize=(297 * mm, 210 * mm),
            topMargin=15 * mm,
            bottomMargin=15 * mm
        )

        element = list()

        element.append(Paragraph('KHARIAR PUBLIC SCHOOL, KHARIAR', heading1))
        element.append(Paragraph('Result of Term Test - ' + number, heading3))
        element.append(Paragraph(class_name, heading3))
        element.append(Paragraph('', heading3))

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
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (2, 0), (-2, -1), 'RIGHT'),
            ('ALIGN', (13, 1), (13, -1), 'CENTER')
        ]))

        element.append(table)
        canvas.build(element)
        print('Term Marksheet Created')
