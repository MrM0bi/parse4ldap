# -*- coding: utf-8 -*-
import sys
import re
import os
from time import sleep
import math

# # Maches the Rows of the inputfiles with the Rows the LDAP accepts
# # -1 : Not present/Do not set | -2 : Auto Fill **Not workin at the moment** | <num> Row of inputfile
# # ##### DO NOT CHANGE THE ORDER OF THE INDEX NAMES ######
# index = [["name", -2], ["firstname", 2], ["lastname", -1], ["email", -1], ["title", -1], ["company", 1],
#          ["workAddress", -1], ["workPostalCode", -1], ["workPhone", 3], ["cellPhone", 5], ["homePhone", 4], ["fax", -1],
#          ["notes", -2]]


index = [["name", -0], ["firstname", -1], ["lastname", -1], ["email", -1], ["title", -1], ["company", -1],
         ["workAddress", -1], ["workPostalCode", -1], ["workPhone", 1], ["cellPhone", 2], ["homePhone", -1], ["fax", -1],
         ["notes", -1]]


HASHEADER = True
IGNORECSVERRORS = False
REPLACEOUTPUTWITHPRECENT = False
CHECKLASTNAMEDUPES = True
QUOTEVALUES = False
EXPORTASLDIF = False
VALIDATEEMAIL = False
NOWORDSINTEXT = True

DELIMITER = ";"


### Client specific settings
ROTTENSTEINERFIX = False
LAUBENREISENFIX = False
TVLAJENFIX = True


### File Confgs

INPUTFILEPATH = OUTPUTFILEPATH = PREEXISTINGLIST = ALREADYPROCESSED = MIDEUPREFIXES = REGPREFIXES = LDAPUSER = ""

INPUTFILEPATH = "<path2file>"

# PREEXISTINGLIST = "<path2file>"
# ALREADYPROCESSED = "<path2file>"

OUTPUTFILEPATH = re.sub("\.csv$", "_output", INPUTFILEPATH, 1)

MIDEUPREFIXES = f"{os.getcwd()}/MidEuPrefixes.csv"
REGPREFIXES = f"{os.getcwd()}/ItalyRegionalPrefixes.csv"


### LDIF Configs

# Format: ou=benjamintest,dc=ldap,dc=vip,dc=rolvoice,dc=it
LDAPUSER = "ou=benjamintest,dc=ldap,dc=vip,dc=rolvoice,dc=it"

CNIDSTARTSWITH = 0
# CNIDPREFIX = "AG"
CNIDPREFIX = "AN"


### needed Arrays

prelist = []
warnings = []
mideuprefixes = []
regprefixes = []

header = []
lineex = []
dupelist = []
lastnames = []

skiplines = 0


# Returns a list of National prefixes (in this case only of Mid-Europe)
def getMidEuPrefixes():
    ret = []
    with open(MIDEUPREFIXES, "r", encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
        for l in lines:
            l = l.split(";")
            ret.append(str(l[1]).strip())
    return ret


# Returns a list of Regional prefixes
def getRegionalPrefixes():
    ret = []
    with open(REGPREFIXES, "r", encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
        for l in lines:
            l = l.split(";")
            ret.append(str(l[0]).strip())
    return ret


# Returns a list of lines, in (hopefully) the same Format as exported
def getPrelist():
    ret = []
    if len(PREEXISTINGLIST) > 3:
        with open(PREEXISTINGLIST, "r", encoding='iso-8859-15', errors='replace') as f:
            if len(f.readline().split(DELIMITER)) == 13:
                lines = f.readlines()
                for l in lines:
                    l = l.split(DELIMITER)
                    ret.append(l)
            else:
                print(
                    "-> [ERROR] Pre-List doesnt have the right amount of fields that an output list should have, aborting to read pre-list...")
    return ret


# # Saves all Lastnames in a dedicated array
# def getLastames():
#     ret = []
#
#     with open(INPUTFILEPATH, "r", encoding='iso-8859-15', errors='replace') as f:
#         lines = f.readlines()
#         for l in lines:
#             l = l.split(DELIMITER)
#             ret.append(str(l[index[0][1]]).strip().lower())
#
#     if len(PREEXISTINGLIST) > 0:
#         with open(PREEXISTINGLIST, "r", encoding='iso-8859-15', errors='replace') as f:
#             lines = f.readlines()
#             for l in lines:
#                 l = l.split(DELIMITER)
#                 ret.append(str(l[2]).strip().lower())
#
#     return ret


# Checks if a Text starts with one of the strings given in the Array
def startsWithValOfArr(text, array):
    for a in array:
        if str(text).startswith(a):
            return len(a)
    return None


# Trys to bring a given Phonenumber into a specific Format
def fixPhone(number):
    national = ""
    regional = ""

    if len(number) >= 3 and len(str(re.sub("\D", "", number))) > 0:
        ognum = number

        if re.search("^\+", number):
            number = "+" + re.sub("\D", "", number)
        else:            number = re.sub("\D", "", number)

        if re.search("^\+", number):
            national = "00" + number[1:3]
            number = number[3:]
        elif re.search("^39", number):
            national = "0039"
            number = re.sub("^39", "", number, 1)

        if TVLAJENFIX and re.search("^4", number):
            number = "0" + number

        if re.search("^00", number):
            if number[2:4] in mideuprefixes:
                national = number[0:4].strip()
                number = number[4:]

        if not re.search("^3|^8", number):
            if len(national) == 0 and number[:2] in mideuprefixes:
                national = "00" + number[:2].strip()
                number = number[2:]

        if re.search("^0", number) and not re.search("^..4", national):
            index = startsWithValOfArr(number[1:], regprefixes)
            if index is not None:
                regional = number[0: 1 + index].strip()
                number = number[1 + index:]
        else:
            index = startsWithValOfArr(number, regprefixes)
            if index is not None and not re.search("^..4", national):
                regional = "0" + number[0: index].strip()
                number = number[index:]

        # Reassemble number
        if len(regional) > 0:
            number = regional + number

        if len(national) > 1:
            number = national.replace("00", "+") + number

        if len(number) > 6 and len(national) == 0:
            number = "+39" + number

        # ToDo: Check for 3-digit Nationals

        # print("#   National: "+national+"   Regional: "+regional+"   Num: "+ognum)

    else:
        number = ""

    return number


# Returns only the digits of a Number (Removes any Text)
def fixNumOnly(number):
    ret = ""
    if re.search("^\+", number):
        number = "+" + re.sub("\D", "", number)
    else:
        number = re.sub("\D", "", number)
    ret = number
    return ret


# Fixes/Replaces specific Charachters or Patterns
def fixText(text):
    text = text.replace(DELIMITER, " ").replace("  ", " ").replace("\"", "").strip()

    if NOWORDSINTEXT and len(re.sub("\W", "", text)) <= 0:
        text = ""

    return text

# Fixes/Validates a E-Mail adress
def fixMail(mail):
    mail = str(mail)
    validmail = False

    if "@" in mail:
        mailparts = mail.split("@")

        if len(mailparts[0]) > 0 and "." in mailparts[1]:
            validmail = True

    if validmail:
        return mail
    else:
        return ""


# Returns only the non-digit Characters of a Number (Removes any Digits)
def randText(number):
    ret = ""

    number = re.sub("\d", "", number).strip()

    if len(number) > 0 and re.search("\w", number):
        ret = number

    return ret


# Checks if one of the Numbers is longer then 17 charachters
def isKaputt(line):
    kaputt = False
    for nlen in [line[8], line[9], line[10], line[11]]:
        if len(nlen) > 17:
            if len(header) != 0:
                kaputt = True
                break

    return kaputt


# Places the columns into a given Squence like its defined in the index array
def columnManager(line, uncnt):
    temp = [""] * 13

    # RotternsteinerFix
    if ROTTENSTEINERFIX:
        if len(line) < 5:
            line.insert(2, '')
            line.insert(3, '')
        # print(str(line) +"  |  "+str(len(line)))

    if len(ALREADYPROCESSED) > 3:
        temp = line
    else:
        # name
        if index[0][1] >= 0:
            temp[0] = fixText(line[index[0][1]]).title()

        # fistname
        if index[1][1] >= 0:
            temp[1] = fixText(line[index[1][1]]).title()

        # lastname
        if index[2][1] >= 0:
            temp[2] = fixText(line[index[2][1]]).title()

        # email
        if index[3][1] >= 0:
            temp[3] = fixText(line[index[3][1]])
            if VALIDATEEMAIL:
                temp[3] = fixMail(temp[3])

        # title
        if index[4][1] >= 0:
            temp[4] = fixText(line[index[4][1]]).title()

        # company
        if index[5][1] >= 0:
            temp[5] = fixText(line[index[5][1]]).title()

        # workAddress
        if index[6][1] >= 0:
            temp[6] = fixText(line[index[6][1]])

        # workPostalCode
        if index[7][1] >= 0:
            temp[7] = fixText(line[index[7][1]])

        # workPhone
        if index[8][1] >= 0:
            temp[8] = fixPhone(line[index[8][1]])
            # temp[8] = line[index[8][1]]

        # cellPhone
        if index[9][1] >= 0:
            temp[9] = fixPhone(line[index[9][1]])
            # temp[9] = line[index[9][1]]

        # homePhone
        if index[10][1] >= 0:
            # Tries the WorkPhone field first
            # if len(temp[8]) <= 1:
            #     temp[8] = fixPhone(line[index[10][1]], linecounter)
            # else:

            temp[10] = fixPhone(line[index[10][1]])
            # temp[10] = line[index[10][1]]

        # fax
        if index[11][1] >= 0:
            # Tries the other number fields first
            # if len(temp[8]) <= 1:
            #     temp[8] = fixText(line[index[11][1]])
            # elif len(index[9]) <= 1:
            #     temp[9] = fixText(line[index[11][1]])
            # elif len(index[10]) <= 1:
            #     temp[10] = fixText(line[index[11][1]])
            # else:

            temp[11] = fixPhone(line[index[11][1]])

        # notes
        if index[12][1] >= 0:
            temptext = line[index[12][1]]
            temp[12] = temptext

            # checks for numbers in the Note and puts them in the Phone-Fields if present
            temptext = temptext.lower()

            if len(temptext) > 0:
                temptext = temptext.replace(' ', '')

                if "fax" not in temptext and "uhr" not in temptext and not re.search("\d+(\/|\-|\.|\-|\:|\,)\d+(\/|\-|\.|\-|\:|\,)\d+", temptext) and not re.search("\d+(\.|\:)\d+",temptext):
                    number = fixNumOnly(line[index[12][1]])
                    number = fixPhone(number)

                    if len(str(number)) > 6 and str(number).rjust(5, '.')[3:5] != "00":
                        # print("Note:  {0}   |->   {1}".format(temp[12].ljust(50), number))
                        if len(temp[8]) <= 1:
                            temp[8] = number
                        elif len(temp[9]) <= 1:
                            temp[9] = number
                        elif len(temp[10]) <= 1:
                            temp[10] = number


        elif index[12][1] == -2:
            temp[12] = str(
                randText(line[index[8][1]]) + " " + randText(line[index[9][1]]) + " " + randText(line[index[10][1]]))
            temp[12] = re.sub("[\+\-\/\"]", "", temp[12]).replace("  ", " ").strip()

    # Empties the array if certain fields are missing
    kill = False
    lenOfRelevant = len(str(temp[0] + temp[1] + temp[2] + temp[3] + temp[5] + temp[8] + temp[9] + temp[10] + temp[12]))
    if INPUTFILEPATH.endswith("arbeitnehmer.csv") and lenOfRelevant <= 5:
        # print("Len: {0}  |  {1}".format(lenOfRelevant, str(temp)))
        kill = True

    # Set Lastname if empty otherwise kill
    if not kill:
        if len(str(temp[2])) <= 1:
            if len(temp[0] + temp[1]) > 1:
                temp[2] = str(temp[0] + " " + temp[1]).strip()
            elif len(str(temp[5])) > 1:
                temp[2] = temp[5]
            else:
                if INPUTFILEPATH.endswith("arbeitnehmer.csv") and len(temp[11]) > 0:
                    temp[2] = "Kunde_{0}".format(temp[11])
                else:
                    temp[2] = "Unnamed_" + str(uncnt)
                    uncnt += 1

        return temp
    else:
        return [""] * 13


# Prints the progress or all the data of the Normalization
def normalizationOutput():
    if not REPLACEOUTPUTWITHPRECENT:
        # Original
        ogl = "{0} | {1} | {2} | {3}".format(fixNumOnly(line[index[8][1]]).ljust(20, " "),
                                             fixNumOnly(line[index[9][1]]).ljust(20, " "),
                                             fixNumOnly(line[index[10][1]]).ljust(20, " "), str(
                randText(line[index[8][1]]) + " " + randText(line[index[9][1]]) + " " + randText(
                    line[index[10][1]])))
        # Proccessed
        mod = "{0} | {1} | {2} | {3}".format(temp[8].ljust(20, " "), temp[9].ljust(20, " "),
                                             temp[10].ljust(20, " "), temp[12], 1)

        if ogl != mod:
            print("Numbers: " + ogl)
            print("     --> " + mod)
    else:
        if linecounter % int(math.sqrt(len(lines))) == 0:
            lod = int(round((linecounter / len(lines)) * 10))
            sys.stdout.write("\r-> [" + "".ljust(lod, "#").ljust(10, " ") + "] (" + str(linecounter) + ")")


# Checks the whole list, warnings and pre-existing list to check for duplicates
def lastnameDupeMitigation(temp):
    found = True
    foundnum = 1
    og_ned = (str(temp[1].strip()) + str(temp[2].strip())).replace(' ', '').lower()
    combined = lineex + warnings + prelist

    # if "simonettacafiero" in og_ned: #ToDo REMOVE
    indupes = checkDupes(og_ned)

    if indupes == -1:
        while found:
            found = False
            needle = og_ned

            if foundnum != 1:
                needle += " " + str(foundnum)

            for cl in combined:
                if needle in (str(cl[1].strip()) + str(cl[2].strip())).replace(' ', '').lower():
                    found = True
                    foundnum += 1
                    break

        if foundnum > 1:
            dupelist.append([og_ned, 2])
            temp[2] += " 2"
    else:
        temp[2] += " " + str(indupes)

    # print("\niD: {} | fn: {} | og: {} | tm: {}".format(str(indupes).ljust(5), str(foundnum).ljust(5), str(og_ned).ljust(30), str(temp)))

    return temp


# Exports the collected and processed data into a .csv file
def fileExport():
    if EXPORTASLDIF:
        extension = ".ldif"
    else:
        extension = ".csv"

    with open(OUTPUTFILEPATH+extension, "w") as e:
        cnidcnt = CNIDSTARTSWITH

        # Writes Header into the File
        if not EXPORTASLDIF:
            for ind in index:
                e.write("\""+ind[0]+"\"")
                if index.index(ind) != len(index) - 1:
                    e.write(DELIMITER)
            e.write("\n")

        # Writes the Data into the File
        for exline in [lineex, warnings]:
            if len(exline) != 0:
                for x in exline:
                    if REPLACEOUTPUTWITHPRECENT:
                        if exline.index(x) % int(math.sqrt(len(exline))) == 0:
                            lod = int(round((exline.index(x) / len(exline)) * 10))
                            sys.stdout.write("\r-> [" + "".ljust(lod, "#").ljust(10, " ") + "]")
                    else:
                        print(x)

                    if not EXPORTASLDIF:
                        if QUOTEVALUES:
                            e.write("\"" + "\""+DELIMITER+"\"".join(x) + "\"\n")
                        else:
                            e.write(DELIMITER.join(x) + "\n")
                    else:
                        if len(LDAPUSER) <= 0:
                            raise Exception("LDAPUSER has to be set if LDIFEXPORT is enabled")


                        cnid = "{}_{}".format(CNIDPREFIX, cnidcnt)
                        cnidcnt += 1

                        # Example: name;firstname;lastname;email;title;company;workAddress;workPostalCode;workPhone;cellPhone;homePhone;fax;notes
                        ldif = "dn: cn={13},{14}\n" \
                               "cn: {1} {2}\n" \
                               "givenName: {1}\n" \
                               "sn: {2}\n" \
                               "mail: {3}\n" \
                               "title: {4}\n" \
                               "o: {5}\n" \
                               "postalAddress: {6}\n" \
                               "postalCode: {7}\n" \
                               "telephoneNumber: {8}\n" \
                               "mobile: {9}\n" \
                               "homePhone: {10}\n" \
                               "facsimileTelephoneNumber: {11}\n" \
                               "description: {12}\n" \
                               "objectclass: top\n" \
                               "objectclass: inetOrgPerson\n\n".format(*x, cnid, LDAPUSER)

                        # print(x)

                        e.write(ldif)

        # if len(warnings) != 0:
        #     for x in warnings:
        #         if REPLACEOUTPUTWITHPRECENT:
        #             if warnings.index(x) % int(math.sqrt(len(warnings))) == 0:
        #                 lod = int(round((warnings.index(x) / len(warnings)) * 10))
        #                 sys.stdout.write("\r-> [" + "".ljust(lod, "#").ljust(10, " ") + "]")
        #         else:
        #             print(x)
        #
        #         if not EXPORTASLDIF:
        #             if QUOTEVALUES:
        #                 e.write("\""+"\"+DELIMITER+\"".join(x)+"\"\n")
        #             else:
        #                 e.write(DELIMITER.join(x) + "\n")
        #         else:
        #             #To Do LDIF Export
    return OUTPUTFILEPATH+extension


# Prints all the Warnings (numbers longer than 17)
def printWarnings():
    headernum = 0
    if len(header) != 0:
        headernum = 1

    for w in warnings:
        print("-> Line: {0} | Work: {1} | Cell: {2} | Home: {3}".format(
            str(len(lineex) + warnings.index(w) + 1 + headernum).ljust(12, " "), w[8].ljust(25, " "),
            w[9].ljust(25, " "), w[10].ljust(25, " ")))


# Checks the dupelist to see if it contains a certain value
def checkDupes(name):
    ret = -1

    for dl in dupelist:
        if name in dl[0]:
            ret = int(dl[1]) + 1
            dupelist[dupelist.index(dl)][1] = ret

    return ret


def checkIfEmpty(temp):
    empty = True
    for t in temp:
        if len(t) > 0:
            empty = False

    return empty


### START ########

# Reads the Preexisting list if given
if len(PREEXISTINGLIST) > 3:
    print("Reading pre-existing List: ")
    prelist = getPrelist()
    print("-> read {0} lines \n-> done".format(str(len(prelist))))

# Checks if an already processed File is given
if len(ALREADYPROCESSED) > 3:
    file = ALREADYPROCESSED
else:
    file = INPUTFILEPATH

# Reading the Prefix-tables
if len(REGPREFIXES) > 0 and len(MIDEUPREFIXES) > 0:
    print("Reading National Prefixes: ")
    mideuprefixes = getMidEuPrefixes()
    print("-> read {0} lines \n-> done".format(str(len(mideuprefixes))))

    print("Reading Regional Prefixes: ")
    regprefixes = getRegionalPrefixes()
    print("-> read {0} lines \n-> done".format(str(len(regprefixes))))

# # Getting all the Lastnames
# print("Reading all Lastnames: ")
# if len(INPUTFILEPATH) > 0:
#     lastnames = getLastames()
#     print("-> read {0} lastnames \n-> done".format(str(len(lastnames))))
# else:
#     print("-> Inputfile not given")


print("Reading and Normalizing Numbers: ")
with open(file, "r", encoding='iso-8859-15') as f:
    export = ""
    linecounter = 0
    avgfields = 0.0
    fixavgfields = False
    unnamedcnt = 1
    saveline = ""

    lines = f.readlines()

    if HASHEADER:
        tmp = lines[0].split(DELIMITER)
        avgfields += len(tmp)
        fixavgfields = True
        header = lines[0]
        lines.pop(0)
        linecounter += 1

    for line in lines:
        errorfound = False
        linecounter += 1

        if True:
            # try:
            if skiplines == 0 or linecounter % skiplines == 0:
                temp = [""] * 13

                ### Check if a line was saved
                if len(saveline) > 0:
                    line = saveline.strip() +" "+ line
                    line = line.replace("  ", " ")
                    saveline = ""


                ### Fix newlines before Splitting
                line = line.replace("\n", " ").strip()

                og_line = line
                line = line.split(DELIMITER)

                cur = len(line)
                # print("Line length: "+str(cur)+"   Linecounter: "+str(linecounter))


                ### ToDo Was once intended to track the number if columns if no Header exists, to see if a line doesn't match
                # if not fixavgfields:
                #     fieldlenmissmatch = False
                #     if cur != int(round(avgfields/linecounter)):
                #         fieldlenmissmatch = True
                #
                #     # avgfields += curato
                #     print("##> AvgField: "+str(avgfields))

                if LAUBENREISENFIX and "Non specificato" not in og_line:
                    print("#####> " + og_line[:50])
                    saveline = og_line
                    linecounter -= 1
                    print("Line " + str(linecounter) + " to short (Avg: " + str(avgfields) + ")!: " + str(cur))
                elif not LAUBENREISENFIX and cur < int(round(avgfields)):
                    saveline = og_line
                    linecounter -= 1
                    print("Line "+str(linecounter)+" to short (Avg: "+str(avgfields)+")!: "+str(cur))
                else:

                    #### Does the main sorting and processing
                    temp = columnManager(line, unnamedcnt)

                    #### Checkts if temp is empty and drops the line in case, otherwise it is added to the list
                    if not checkIfEmpty(temp):

                        ### Output:
                        normalizationOutput()

                        ### Decide if adding to Warnings
                        kaputt = isKaputt(temp)

                        ### Check if Lastname exists and replace in case
                        if CHECKLASTNAMEDUPES:
                            lastnameDupeMitigation(temp)

                        if kaputt:
                            warnings.append(temp)
                        else:
                            lineex.append(temp)

        # except Exception as e:
        #     pass
        #     errorfound = True

        ### Normalization Error Output
        if errorfound and IGNORECSVERRORS:
            print(
                "[ERROR] Something is wrong while reading line {0}. Please check the .csv formatting \nLine: \'{1}\'".format(
                    linecounter, str(og_line)))
        elif errorfound:
            raise Exception(
                "[ERROR] Something is wrong while reading line {0}. Please check the .csv formatting \nLine: \'{1}\'".format(
                    linecounter, str(og_line)))

    sleep(0.25)
    print("")

    ### EXPORT ########
    print("\nExporting to File:")
    path = fileExport()
    print("\n-> Exported to \"" +path+ "\"")

    ### WARNINGS ########
    print("\nWarnings: " + str(len(warnings)))
    print("-> Moved the Warnings to the end of the File")
    printWarnings()

print("\n- Finished -")
exit(0)
