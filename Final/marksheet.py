import os
from os.path import join

import shutil
import pandas as pd

from generate_pdf import GeneratePDF


class Marksheet:
    """Marksheet"""

    def __init__(self, path, class_name):
        self.marksheet = pd.read_csv(join(path, 'Marksheet.csv'))
        self.size = len(self.marksheet)
        output_path = join(path, 'Output')

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
        GeneratePDF.generate_unit_marksheet(self.marksheet_dict, self.size, output_path, class_name)
        GeneratePDF.generate_term_marksheet(self.marksheet_dict, self.size, output_path, class_name)

    def create_unit_10(self):
        u_english_10 = [x / 3 for x in self.marksheet['U_English']]
        u_hindi_10 = [x / 3 for x in self.marksheet['U_Hindi']]
        u_odia_10 = [x / 3 for x in self.marksheet['U_Odia']]
        u_maths_10 = [x / 3 for x in self.marksheet['U_Maths']]
        u_science_10 = [x / 3 for x in self.marksheet['U_Science']]
        u_sst_10 = [x / 3 for x in self.marksheet['U_SST']]

        self.marksheet.insert(3, 'U_English_10', u_english_10)
        self.marksheet.insert(5, 'U_Hindi_10', u_hindi_10)
        self.marksheet.insert(7, 'U_Odia_10', u_odia_10)
        self.marksheet.insert(9, 'U_Maths_10', u_maths_10)
        self.marksheet.insert(11, 'U_Science_10', u_science_10)
        self.marksheet.insert(13, 'U_SST_10', u_sst_10)

    def calculate_total_mark_percentage(self):
        total_mark = list()
        for i in range(self.size):
            total_mark.append(
                self.marksheet.loc[i]['U_English_10'] +
                self.marksheet.loc[i]['U_Hindi_10'] +
                self.marksheet.loc[i]['U_Odia_10'] +
                self.marksheet.loc[i]['U_Maths_10'] +
                self.marksheet.loc[i]['U_Science_10'] +
                self.marksheet.loc[i]['U_SST_10'] +
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
        percent = [x / 7 for x in total_mark]

        self.marksheet.insert(28, 'Total Mark', total_mark)
        self.marksheet.insert(29, 'Percentage', percent)

    def create_marksheet_dictionary(self):
        marksheet_dict = dict()
        for i in range(self.size):
            student_dict = dict()

            student_dict['Roll No.'] = self.marksheet.loc[i]['Roll No.']
            student_dict['Name'] = self.marksheet.loc[i]['Name']

            english = dict()
            english['Unit'] = self.marksheet.loc[i]['U_English']
            english['Unit_10'] = '%.2f' % (self.marksheet.loc[i]['U_English_10'])
            english['Oral'] = self.marksheet.loc[i]['O_English']
            english['Term'] = self.marksheet.loc[i]['T_English']
            english['Total'] = '%.2f' % (
                    self.marksheet.loc[i]['U_English_10'] +
                    self.marksheet.loc[i]['O_English'] +
                    self.marksheet.loc[i]['T_English']
            )
            english['Total/2'] = '%.2f' % (
                    (self.marksheet.loc[i]['U_English_10'] +
                     self.marksheet.loc[i]['O_English'] +
                     self.marksheet.loc[i]['T_English']) / 2
            )
            student_dict['English'] = english

            hindi = dict()
            hindi['Unit'] = self.marksheet.loc[i]['U_Hindi']
            hindi['Unit_10'] = '%.2f' % (self.marksheet.loc[i]['U_Hindi_10'])
            hindi['Oral'] = self.marksheet.loc[i]['O_Hindi']
            hindi['Term'] = self.marksheet.loc[i]['T_Hindi']
            hindi['Total'] = '%.2f' % (
                    self.marksheet.loc[i]['U_Hindi_10'] +
                    self.marksheet.loc[i]['O_Hindi'] +
                    self.marksheet.loc[i]['T_Hindi']
            )
            hindi['Total/2'] = '%.2f' % (
                    (self.marksheet.loc[i]['U_Hindi_10'] +
                     self.marksheet.loc[i]['O_Hindi'] +
                     self.marksheet.loc[i]['T_Hindi']) / 2
            )
            student_dict['Hindi'] = hindi

            odia = dict()
            odia['Unit'] = self.marksheet.loc[i]['U_Odia']
            odia['Unit_10'] = '%.2f' % (self.marksheet.loc[i]['U_Odia_10'])
            odia['Oral'] = self.marksheet.loc[i]['O_Odia']
            odia['Term'] = self.marksheet.loc[i]['T_Odia']
            odia['Total'] = '%.2f' % (
                    self.marksheet.loc[i]['U_Odia_10'] +
                    self.marksheet.loc[i]['O_Odia'] +
                    self.marksheet.loc[i]['T_Odia']
            )
            odia['Total/2'] = '%.2f' % (
                    (self.marksheet.loc[i]['U_Odia_10'] +
                     self.marksheet.loc[i]['O_Odia'] +
                     self.marksheet.loc[i]['T_Odia']) / 2
            )
            student_dict['Odia'] = odia

            maths = dict()
            maths['Unit'] = self.marksheet.loc[i]['U_Maths']
            maths['Unit_10'] = '%.2f' % (self.marksheet.loc[i]['U_Maths_10'])
            maths['Oral'] = self.marksheet.loc[i]['O_Maths']
            maths['Term'] = self.marksheet.loc[i]['T_Maths']
            maths['Total'] = '%.2f' % (
                    self.marksheet.loc[i]['U_Maths_10'] +
                    self.marksheet.loc[i]['O_Maths'] +
                    self.marksheet.loc[i]['T_Maths']
            )
            maths['Total/2'] = '%.2f' % (
                    (self.marksheet.loc[i]['U_Maths_10'] +
                     self.marksheet.loc[i]['O_Maths'] +
                     self.marksheet.loc[i]['T_Maths']) / 2
            )
            student_dict['Maths'] = maths

            science = dict()
            science['Unit'] = self.marksheet.loc[i]['U_Science']
            science['Unit_10'] = '%.2f' % (self.marksheet.loc[i]['U_Science_10'])
            science['Oral'] = self.marksheet.loc[i]['O_Science']
            science['Term'] = self.marksheet.loc[i]['T_Science']
            science['Total'] = '%.2f' % (
                    self.marksheet.loc[i]['U_Science_10'] +
                    self.marksheet.loc[i]['O_Science'] +
                    self.marksheet.loc[i]['T_Science']
            )
            science['Total/2'] = '%.2f' % (
                    (self.marksheet.loc[i]['U_Science_10'] +
                     self.marksheet.loc[i]['O_Science'] +
                     self.marksheet.loc[i]['T_Science']) / 2
            )
            student_dict['Science'] = science

            sst = dict()
            sst['Unit'] = self.marksheet.loc[i]['U_SST']
            sst['Unit_10'] = '%.2f' % (self.marksheet.loc[i]['U_SST_10'])
            sst['Oral'] = self.marksheet.loc[i]['O_SST']
            sst['Term'] = self.marksheet.loc[i]['T_SST']
            sst['Total'] = '%.2f' % (
                    self.marksheet.loc[i]['U_SST_10'] +
                    self.marksheet.loc[i]['O_SST'] +
                    self.marksheet.loc[i]['T_SST']
            )
            sst['Total/2'] = '%.2f' % (
                    (self.marksheet.loc[i]['U_SST_10'] +
                     self.marksheet.loc[i]['O_SST'] +
                     self.marksheet.loc[i]['T_SST']) / 2
            )
            student_dict['SST'] = sst

            student_dict['GK'] = self.marksheet.loc[i]['T_GK']
            student_dict['Computer'] = self.marksheet.loc[i]['T_Computer']

            student_dict['Total Mark'] = '%.2f' % (self.marksheet.loc[i]['Total Mark'])
            student_dict['Percentage'] = '%.2f' % (self.marksheet.loc[i]['Percentage'])

            marksheet_dict[i] = student_dict
        return marksheet_dict
