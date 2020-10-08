import os
from os.path import join

import shutil
import pandas as pd

from generate_pdf import GeneratePDF


class FinalMarksheet:
    """Final Marksheet"""

    def __init__(self, path, text, class_name):
        self.marksheet = pd.read_csv(join(path, 'Final.csv'))
        self.size = len(self.marksheet)
        output_path = join(path, 'Final')

        del (self.marksheet['Timestamp'])

        try:
            os.mkdir(output_path)
        except FileExistsError:
            shutil.rmtree(output_path)
            os.mkdir(output_path)

        self.marksheet.to_excel(join(output_path, 'Raw Marksheet.xlsx'))

        self.create_unit_10()
        self.calculate_total_mark_percentage()
        self.marksheet_dict = self.create_marksheet_dictionary()

        self.marksheet.to_excel(join(output_path, 'Final Marksheet.xlsx'))

        GeneratePDF.generate_student_marksheet(self.marksheet_dict, self.size, output_path)
        GeneratePDF.generate_unit_marksheet(self.marksheet_dict, self.size, output_path, text, class_name)
        GeneratePDF.generate_term_marksheet(self.marksheet_dict, self.size, output_path, text, class_name)

    def create_unit_10(self):
        u_english_10 = ['%.2f' % (x / 3) for x in self.marksheet['U_English']]
        u_hindi_10 = ['%.2f' % (x / 3) for x in self.marksheet['U_Hindi']]
        u_odia_10 = ['%.2f' % (x / 3) for x in self.marksheet['U_Odia']]
        u_maths_10 = ['%.2f' % (x / 3) for x in self.marksheet['U_Maths']]
        u_science_10 = ['%.2f' % (x / 3) for x in self.marksheet['U_Science']]
        u_sst_10 = ['%.2f' % (x / 3) for x in self.marksheet['U_SST']]

        self.marksheet.insert(3, 'U_English_10', u_english_10)
        self.marksheet.insert(5, 'U_Hindi_10', u_hindi_10)
        self.marksheet.insert(7, 'U_Odia_10', u_odia_10)
        self.marksheet.insert(9, 'U_Maths_10', u_maths_10)
        self.marksheet.insert(11, 'U_Science_10', u_science_10)
        self.marksheet.insert(13, 'U_SST_10', u_sst_10)

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
        mark_obtained = list()
        for i in range(self.size):
            mark_obtained.append(
                float(self.marksheet.loc[i]['U_English_10']) +
                float(self.marksheet.loc[i]['U_Hindi_10']) +
                float(self.marksheet.loc[i]['U_Odia_10']) +
                float(self.marksheet.loc[i]['U_Maths_10']) +
                float(self.marksheet.loc[i]['U_Science_10']) +
                float(self.marksheet.loc[i]['U_SST_10']) +
                self.marksheet.loc[i]['O_English'] +
                self.marksheet.loc[i]['O_Hindi'] +
                self.marksheet.loc[i]['O_Odia'] +
                self.marksheet.loc[i]['O_Maths'] +
                self.marksheet.loc[i]['O_Science'] +
                self.marksheet.loc[i]['O_SST'] +
                self.marksheet.loc[i]['T_English'] +
                self.marksheet.loc[i]['T_Hindi'] +
                self.marksheet.loc[i]['T_Odia'] +
                self.marksheet.loc[i]['T_Maths'] +
                self.marksheet.loc[i]['T_Science'] +
                self.marksheet.loc[i]['T_SST'] +
                self.marksheet.loc[i]['T_GK'] +
                self.marksheet.loc[i]['T_Computer']
            )
        total_mark = [700] * self.size
        percent = ['%.2f' % (x / 7) for x in mark_obtained]
        grade = self.calculate_grade(percent)

        self.marksheet.insert(28, 'Mark Obtained', mark_obtained)
        self.marksheet.insert(29, 'Total Mark', total_mark)
        self.marksheet.insert(30, 'Percentage', percent)
        self.marksheet.insert(31, 'Grade', grade)

    def create_marksheet_dictionary(self):
        marksheet_dict = dict()
        for i in range(self.size):
            student_dict = dict()

            # Roll No.
            student_dict['Roll No.'] = self.marksheet.loc[i]['Roll No.']
            # Name
            student_dict['Name'] = self.marksheet.loc[i]['Name']

            # English
            english = dict()
            english['Unit'] = self.marksheet.loc[i]['U_English']
            english['Unit_10'] = '%.2f' % float(self.marksheet.loc[i]['U_English_10'])
            english['Oral'] = self.marksheet.loc[i]['O_English']
            english['Term'] = self.marksheet.loc[i]['T_English']
            english['Total'] = '%.2f' % (
                    float(self.marksheet.loc[i]['U_English_10']) +
                    self.marksheet.loc[i]['O_English'] +
                    self.marksheet.loc[i]['T_English']
            )
            english['Total/2'] = '%.2f' % (
                    (float(self.marksheet.loc[i]['U_English_10']) +
                     self.marksheet.loc[i]['O_English'] +
                     self.marksheet.loc[i]['T_English']) / 2
            )
            student_dict['English'] = english

            # Hindi
            hindi = dict()
            hindi['Unit'] = self.marksheet.loc[i]['U_Hindi']
            hindi['Unit_10'] = '%.2f' % float(self.marksheet.loc[i]['U_Hindi_10'])
            hindi['Oral'] = self.marksheet.loc[i]['O_Hindi']
            hindi['Term'] = self.marksheet.loc[i]['T_Hindi']
            hindi['Total'] = '%.2f' % (
                    float(self.marksheet.loc[i]['U_Hindi_10']) +
                    self.marksheet.loc[i]['O_Hindi'] +
                    self.marksheet.loc[i]['T_Hindi']
            )
            hindi['Total/2'] = '%.2f' % (
                    (float(self.marksheet.loc[i]['U_Hindi_10']) +
                     self.marksheet.loc[i]['O_Hindi'] +
                     self.marksheet.loc[i]['T_Hindi']) / 2
            )
            student_dict['Hindi'] = hindi

            # Odia
            odia = dict()
            odia['Unit'] = self.marksheet.loc[i]['U_Odia']
            odia['Unit_10'] = '%.2f' % float(self.marksheet.loc[i]['U_Odia_10'])
            odia['Oral'] = self.marksheet.loc[i]['O_Odia']
            odia['Term'] = self.marksheet.loc[i]['T_Odia']
            odia['Total'] = '%.2f' % (
                    float(self.marksheet.loc[i]['U_Odia_10']) +
                    self.marksheet.loc[i]['O_Odia'] +
                    self.marksheet.loc[i]['T_Odia']
            )
            odia['Total/2'] = '%.2f' % (
                    (float(self.marksheet.loc[i]['U_Odia_10']) +
                     self.marksheet.loc[i]['O_Odia'] +
                     self.marksheet.loc[i]['T_Odia']) / 2
            )
            student_dict['Odia'] = odia

            # Maths
            maths = dict()
            maths['Unit'] = self.marksheet.loc[i]['U_Maths']
            maths['Unit_10'] = '%.2f' % float(self.marksheet.loc[i]['U_Maths_10'])
            maths['Oral'] = self.marksheet.loc[i]['O_Maths']
            maths['Term'] = self.marksheet.loc[i]['T_Maths']
            maths['Total'] = '%.2f' % (
                    float(self.marksheet.loc[i]['U_Maths_10']) +
                    self.marksheet.loc[i]['O_Maths'] +
                    self.marksheet.loc[i]['T_Maths']
            )
            maths['Total/2'] = '%.2f' % (
                    (float(self.marksheet.loc[i]['U_Maths_10']) +
                     self.marksheet.loc[i]['O_Maths'] +
                     self.marksheet.loc[i]['T_Maths']) / 2
            )
            student_dict['Maths'] = maths

            # Science
            science = dict()
            science['Unit'] = self.marksheet.loc[i]['U_Science']
            science['Unit_10'] = '%.2f' % float(self.marksheet.loc[i]['U_Science_10'])
            science['Oral'] = self.marksheet.loc[i]['O_Science']
            science['Term'] = self.marksheet.loc[i]['T_Science']
            science['Total'] = '%.2f' % (
                    float(self.marksheet.loc[i]['U_Science_10']) +
                    self.marksheet.loc[i]['O_Science'] +
                    self.marksheet.loc[i]['T_Science']
            )
            science['Total/2'] = '%.2f' % (
                    (float(self.marksheet.loc[i]['U_Science_10']) +
                     self.marksheet.loc[i]['O_Science'] +
                     self.marksheet.loc[i]['T_Science']) / 2
            )
            student_dict['Science'] = science

            # SST
            sst = dict()
            sst['Unit'] = self.marksheet.loc[i]['U_SST']
            sst['Unit_10'] = '%.2f' % float(self.marksheet.loc[i]['U_SST_10'])
            sst['Oral'] = self.marksheet.loc[i]['O_SST']
            sst['Term'] = self.marksheet.loc[i]['T_SST']
            sst['Total'] = '%.2f' % (
                    float(self.marksheet.loc[i]['U_SST_10']) +
                    self.marksheet.loc[i]['O_SST'] +
                    self.marksheet.loc[i]['T_SST']
            )
            sst['Total/2'] = '%.2f' % (
                    (float(self.marksheet.loc[i]['U_SST_10']) +
                     self.marksheet.loc[i]['O_SST'] +
                     self.marksheet.loc[i]['T_SST']) / 2
            )
            student_dict['SST'] = sst

            # GK
            student_dict['GK'] = self.marksheet.loc[i]['T_GK']
            # Computer
            student_dict['Computer'] = self.marksheet.loc[i]['T_Computer']

            # Total Mark
            student_dict['Total Mark'] = '%.2f' % (self.marksheet.loc[i]['Total Mark'])
            # Percentage
            student_dict['Percentage'] = '%.2f' % float(self.marksheet.loc[i]['Percentage'])

            marksheet_dict[i] = student_dict
        return marksheet_dict
