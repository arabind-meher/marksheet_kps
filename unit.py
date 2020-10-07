import os
from os.path import join

import shutil
import pandas as pd

from generate_pdf import GeneratePDF


class UnitMarksheet:
    """Unit Marksheet"""

    def __init__(self, path, class_name):
        self.marksheet = pd.read_csv(join(path, 'Unit.csv'))
        self.size = len(self.marksheet)
        output_path = join(path, 'Unit')

        del (self.marksheet['Timestamp'])

        try:
            os.mkdir(output_path)
        except FileExistsError:
            shutil.rmtree(output_path)
            os.mkdir(output_path)

        self.marksheet.to_excel(join(output_path, 'Raw Marksheet.xlsx'))

        self.calculate_total_mark_percentage()
        self.marksheet_dict = self.create_marksheet_dictionary()

        self.marksheet.to_excel(join(output_path, 'Final Marksheet.xlsx'))

        GeneratePDF.generate_unit_marksheet_standalone(self.marksheet_dict, self.size, output_path, class_name)

    @staticmethod
    def calculate_grade(percentage):
        grade = list()
        for x in percentage:
            x = float(x)
            if x > 90:
                grade.append('A1')
            elif x > 80:
                grade.append('A2')
            elif x > 70:
                grade.append('B1')
            elif x > 60:
                grade.append('B2')
            elif x > 50:
                grade.append('C1')
            elif x > 40:
                grade.append('C2')
            elif x > 33:
                grade.append('D')
            else:
                grade.append('E')
        return grade

    def calculate_total_mark_percentage(self):
        total_mark = list()
        for i in range(self.size):
            total_mark.append(
                self.marksheet.loc[i]['U_English'] +
                self.marksheet.loc[i]['U_Hindi'] +
                self.marksheet.loc[i]['U_Odia'] +
                self.marksheet.loc[i]['U_Maths'] +
                self.marksheet.loc[i]['U_Science'] +
                self.marksheet.loc[i]['U_SST']
            )
        full_mark = [180] * self.size
        percent = ['%.2f' % (x / 1.8) for x in total_mark]
        grade = self.calculate_grade(percent)

        self.marksheet.insert(8, 'Mark Obtained', total_mark)
        self.marksheet.insert(9, 'Total Mark', full_mark)
        self.marksheet.insert(10, 'Percentage', percent)
        self.marksheet.insert(11, 'Grade', grade)

    def create_marksheet_dictionary(self):
        marksheet_dict = dict()
        for i in range(self.size):
            student_dict = dict()

            student_dict['Roll No.'] = self.marksheet.loc[i]['Roll No.']
            student_dict['Name'] = self.marksheet.loc[i]['Name']

            student_dict['English'] = self.marksheet.loc[i]['U_English']
            student_dict['Hindi'] = self.marksheet.loc[i]['U_Hindi']
            student_dict['Odia'] = self.marksheet.loc[i]['U_Odia']
            student_dict['Maths'] = self.marksheet.loc[i]['U_Maths']
            student_dict['Science'] = self.marksheet.loc[i]['U_Science']
            student_dict['SST'] = self.marksheet.loc[i]['U_SST']

            student_dict['Mark Obtained'] = self.marksheet.loc[i]['Mark Obtained']
            student_dict['Total Mark'] = self.marksheet.loc[i]['Total Mark']
            student_dict['Percentage'] = self.marksheet.loc[i]['Percentage']
            student_dict['Grade'] = self.marksheet.loc[i]['Grade']

            marksheet_dict[i] = student_dict
        return marksheet_dict
