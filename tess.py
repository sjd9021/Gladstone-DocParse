import cv2
import pytesseract
import os
import openpyxl
import pandas as pd
from openpyxl.styles import Font, Alignment
import re


class Tesseract():
    print(0)
    def __init__(self, file):
        self.file = file
        print(self.file)
        img = cv2.imread(self.file)
        self.text1 = pytesseract.image_to_string(img)
        self.text_lines = self.text1.replace(" ", "")
        self.text_lines = self.text_lines.lower()
        self.text_lines = self.text_lines.splitlines()
        while "" in self.text_lines:
            self.text_lines.remove("")
    def date(self):
        survey_date = self.text_lines[5]
        print(survey_date)
        if survey_date == 'invoice':
            survey_date = self.text_lines[6]
        return survey_date

    def insurers(self):
        for i in self.text_lines:
            if 'ina/cwith:' in i:
                print(i)
                return i[10:].upper()

    def policy_num(self):
        for i in self.text_lines:
            if 'policyno.:' in i:
                print(i)
                return i[10:].upper()

    def total_amount(self):
        a = self.text1.rfind('$')
        return self.text1[a:a + 7]

    def branch(self):
        lists = []
        with open('sam1.csv', 'r') as file:
            x = file.read()
            x = os.linesep.join([s for s in x.splitlines() if s])
            x = x.splitlines()
            for i in x:
                if not i.isnumeric():
                    lists.append(i)
            insurer_branch = set(lists)
        for i in insurer_branch:
            if i.lower() in self.text1.lower():
                return i.upper()

    def our_ref(self):
        for i in self.text_lines:
            if 'ourref.:' in i:
                return i[8:].upper()
            if 'ourref:' in i:
                return i[7:].upper()


