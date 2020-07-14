import cv2
import pytesseract
import os
import csv


def image_to_text(img):
    text1 = pytesseract.image_to_string(img)
    a = text1.replace(" ", "")
    a = a.lower()
    return a


img = cv2.imread("sam.png")
text = image_to_text(img)
text_lines = text.splitlines()
while "" in text_lines:
    text_lines.remove("")


def date():
    survey_date = text_lines[5]
    if survey_date == 'invoice':
        survey_date = text_lines[6]
    return survey_date


def insurers():
    for i in text_lines:
        if 'ina/cwith:' in i:
            return i[10:].upper()


def policy_num():
    for i in text_lines:
        if 'policyno.:' in i:
            return i[10:].upper()


def total_amount():
    a = text.rfind('$')
    return text[a:a+7]



def branch():
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
        if i.lower() in text.lower():
            return i.upper()


def our_ref():
    for i in text_lines:
        if 'ourref.:' in i:
            return i[8:].upper()
        if 'ourref:' in i:
            return i[7:].upper()


our_ref = our_ref()
branch = branch()
insurers = insurers()
date = date()
total_amount = total_amount()
policy = policy_num()
