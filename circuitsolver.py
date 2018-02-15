"""
EMAIL:

Sir,

I have completed the program for solving DA-1. I would like to test it before submitting it. Could you please send me the excel sheet of phone numbers?

Thank you for your time.

Abitha K Thyagarajan

"""



import openpyxl
from math import pi
from cmath import rect, polar
from sys import argv
from easygui import *


# get name of the Excel file and name of the sheet

#filename, fieldValues[0], fieldValues[1] = argv
#fieldValues[0] = enterbox('Excel file: ')
#fieldValues[1] = enterbox('Sheet name: ')

msg = "Enter the details of the excel file.\n\nMake sure the excel file is in the same folder as this application."
title = "DA-1 Solver"
fieldNames = ["Excel file", "Sheet name"]
fieldValues = multenterbox(msg, title, fieldNames)
# make sure that none of the fields were left blank
while 1:
    errmsg = ""
    for i, name in enumerate(fieldNames):
        if fieldValues[i].strip() == "":
          errmsg += "{} is a required field.\n\n".format(name)
    if errmsg == "":
        break # no problems found
    fieldValues = multenterbox(errmsg, title, fieldNames, fieldValues)
    if fieldValues is None:
        break




# open the Excel file -- ok
wb = openpyxl.load_workbook(fieldValues[0])
sheet = wb[fieldValues[1]]


# replace zeroes with average digit -- ok
def davg(d):
    dav = sum(d) / 10
    if dav - 0.5 == int(dav):
        if (dav + 0.5) % 2 == 0: dav += 0.5
        else: dav -= 0.5
    else: dav = round(dav)
    for i in range(len(d)):
        if d[i] == 0: d[i] = dav
    return d


# question 1
# code ok
def q1(d):
    z = ((d[1] - 1j / d[2]) * d[3]) / (d[1] + d[3] - 1j / d[2])
    i1 = d[4] / (d[3] + d[8]) # dc value?
    i2 = polar(z * d[6] / (d[8] + z + complex(0, d[7])))
    i3 = polar(d[3] * d[5] / ( (d[1] + d[3] - complex(0, 1/(2*d[2]))) * (d[3] + d[8] + complex(0, 2*d[7])) - d[3]**2 ))
    return [round(i1, 3), str(round(i2[0], 3)) + "cos(" + str(round(i2[1], 3)) + ")", str(round(i3[0], 3)) + "cos(" + str(round(i3[1], 3)) + ")"]
    #return 0 # return a list of strings


# question 2 -- CHECK
# code ok
def q2(d):
    z = d[2] * (d[1] - complex(0, d[5])) / (d[1] + d[2] - complex(0, d[5]))     + d[4] * (d[3] + complex(0, d[6])) / (d[4] + d[3] + complex(0, d[6]))
    v = rect(d[10], pi) * ( (d[1] - complex(0, d[5])) / (d[1] + d[2] - complex(0, d[5]) - d[4] / (d[3] + d[4] + complex(0, d[6]) ))) # WRONG??
    p = abs(v) ** 2 / (8 * z.real)
    return ['%.3f %+.3fj' % (v.real, v.imag), '%.3f %+.3fj' % (z.real, z.imag), round(p, 3)]


# question 3 -- CHECK
# code ok
def q3(d):
    m = [[1, -1, 0, 0, 1, 0], [0, 1, 1, -1, 0, 0], [0, 0, 0, -1, 1, 1]]
    #zb = [[d[2], 0, 0, 0, 0, 0], [0, d[3], 0, 0, 0, 0], [0, 0, d[6], 0, 0, 0], [0, 0, 0, d[9], 0, 0], [0, 0, 0, 0, d[7], 0], [0, 0, 0, 0, 0, davg(d)]]
    #eb = [[d[1]], [d[4]], [d[5]], [d[8]], [0], [d[10]]]
    de = d[2]*d[3]*d[7]+d[2]*d[3]*d[0]+d[2]*d[3]*d[9]+d[2]*d[6]*d[7]+d[2]*d[6]*d[0]+d[2]*d[6]*d[9]+d[2]*d[7]*d[9]+d[2]*d[0]*d[9]+d[3]*d[6]*d[7]+d[3]*d[6]*d[0]+d[3]*d[6]*d[9]+d[3]*d[7]*d[0]+d[3]*d[0]*d[9]+d[6]*d[7]*d[0]+d[6]*d[7]*d[9]+d[7]*d[0]*d[9]
    de1 = d[1]*d[3]*d[7]+d[1]*d[3]*d[0]+d[1]*d[3]*d[9]+d[1]*d[6]*d[7]+d[1]*d[6]*d[0]+d[1]*d[6]*d[9]+d[1]*d[7]*d[9]+d[1]*d[0]*d[9]-d[3]*d[7]*d[10]+d[5]*d[3]*d[7]-d[3]*d[8]*d[0]-d[3]*d[10]*d[9]+d[5]*d[3]*d[0]+d[5]*d[3]*d[9]-d[4]*d[6]*d[7]-d[4]*d[6]*d[0]-d[4]*d[6]*d[9]-d[4]*d[0]*d[9]+d[6]*d[7]*d[8]-d[6]*d[7]*d[10]-d[7]*d[10]*d[9]+d[5]*d[7]*d[9]
    de2 = (d[7]+d[0]+d[9])*(d[3]*(d[1]-d[4])+(d[2]+d[3]+d[7])*(d[4]+d[5]-d[8]))+d[7]*(d[1]*d[9]-d[4]*d[7]-d[4]*d[9]-d[5]*d[7]+d[7]*d[8])-(d[10]-d[8])*(d[2]*d[9]+d[3]*d[7]+d[3]*d[9]+d[7]*d[9])
    de3 = -d[1]*d[3]*d[7]-d[1]*d[3]*d[9]-d[1]*d[6]*d[7]-d[1]*d[7]*d[9]-d[2]*d[3]*d[8]+d[2]*d[3]*d[10]-d[2]*d[4]*d[9]-d[2]*d[5]*d[9]-d[2]*d[6]*d[8]+d[2]*d[6]*d[10]+d[2]*d[10]*d[9]-d[3]*d[5]*d[7]-d[3]*d[5]*d[9]-d[3]*d[6]*d[8]+d[3]*d[6]*d[10]+d[3]*d[7]*d[10]+d[3]*d[10]*d[9]+d[4]*d[6]*d[7]-d[5]*d[7]*d[9]-d[6]*d[7]*d[8]+d[6]*d[7]*d[10]+d[7]*d[10]*d[9]
    ic = [de1 / de, de2 / de, de3 / de]
    ic = [round(x, 3) for x in ic]
    q = [[1, 1, -1, 0, 0, 0], [-1, 0, 0, 0, 1, -1], [0, 0, 1, 1, 0, 1]]
    ans = ic + m + q
    return ans


# writing the answers to the sheet

# headings
sheet[str("C" + str(1))] = "New phone number"
sheet[str("H" + str(1))], sheet[str("I" + str(1))], sheet[str("J" + str(1))] = "Q1 - I1", "Q1 - I2", "Q1 - I3"
sheet[str("K" + str(1))], sheet[str("L" + str(1))], sheet[str("M" + str(1))] = "Q2 - V", "Q2 - Z", "Q2 - P"
sheet[str("N" + str(1))], sheet[str("O" + str(1))], sheet[str("P" + str(1))], sheet[str("Q" + str(1))] = "Q3 - I1", "Q3 - I2", "Q3 - I3", "Q3 - Q"


for i in range(5, sheet.max_row + 1):
    # getting the phone number & modifying it with Davg -- ok
    d = list('0' + str(sheet["B" + str(i)].value))
    #print(d)
    d = [int(x) for x in d]
    d = davg(d)
    #print(d)

    # question 1 -- needs three cells, H, I, and J
    sheet[str("H" + str(i))], sheet[str("I" + str(i))], sheet[str("J" + str(i))] = q1(d)[0], q1(d)[1], q1(d)[2]
    # question 2 -- cells K, L, M
    sheet[str("K" + str(i))], sheet[str("L" + str(i))], sheet[str("M" + str(i))] = q2(d)[0], q2(d)[1], q2(d)[2]
    # question 3 -- cells N, O, P
    sheet[str("N" + str(i))], sheet[str("O" + str(i))], sheet[str("P" + str(i))], sheet[str("Q" + str(i))] = q3(d)[0], q3(d)[1], q3(d)[2], str(q3(d)[3]) + str(q3(d)[4]) + str(q3(d)[5])
    #sheet[str("Q" + str(i))] = str(d[1:])

    # new phone number
    sheet[str("C" + str(i))] = "".join((str(x) for x in d[1:]))

wb.save(fieldValues[0])
