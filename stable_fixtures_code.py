# -*- coding: utf-8 -*-
"""
Created on Wed Oct 13 19:23:22 2021

@author: Zahid Ur Rahman
"""

import os
import copy
from xlrd import open_workbook
import xlsxwriter
book = open_workbook("1_102.xlsx")
sheet = book.sheet_by_index(0)
# taking the data from EXCEL file
list_from_excel = []
for i in range(1, sheet.nrows):
    row = []
    for j in range(sheet.ncols):
        if sheet.cell_value(i, j) != '':
            row.append(sheet.cell_value(i, j))
    list_from_excel.append(row)
##############################################################################
# creating a list with preferences
preferenceList = []
frequency = []
for i in range(len(list_from_excel)):
    subPreference = [list_from_excel[i][1]]
    frequency.append(list_from_excel[i][0])
    subOfContent = []
    for j in range(len(list_from_excel[i])):
        if j > 1:
            subOfPerson = [list_from_excel[i][j]]
            subOfContent.append(subOfPerson)
    subPreference.append(subOfContent)
    preferenceList.append(subPreference)
##############################################################################
# Creating dictionary that contains all the preferences of every person, no numbers
sortedPreference = {}
for i in range(len(preferenceList)):
    preferences = []    
    for j in range(len(preferenceList[i][1])):
        preferences.append(preferenceList[i][1][j][0])
    sortedPreference[preferenceList[i][0]] = preferences
    input_data = {"preferences": sortedPreference}
    o_data = copy.deepcopy(input_data)
    f_data = copy.deepcopy(input_data)
    a_data = copy.deepcopy(input_data)
##############################################################################
def step1(inputList):
    proposals = {}
    numProposals = {}
    queue = []    
    for i in inputList["preferences"]:        
        queue.append(i)
        proposals[i] = None
        numProposals[i] = 0
    queu = queue
    tmpPreferences = copy.deepcopy(inputList["preferences"])
    while not len(queu) == 0:        
        i = queu[0]
        numProposals[i] += 1
        for j in inputList["preferences"][i]:
            if proposals[j] == None:
                del queu[0]
                proposals[j] = i
                break
            elif proposals[j] != i:
                if frequency[int(j-1)] == 1:
                    current_index = inputList["preferences"][j].index(i)
                    other_index = inputList["preferences"][j].index(proposals[j])                
                    if current_index < other_index:
                        del queu[0]
                        queu.insert(0, proposals[j])
                        # Remove old proposal symmetrically
                        tmpPreferences[proposals[j]].remove(j)
                        tmpPreferences[j].remove(proposals[j])
                        proposals[j] = i
                        break
                    else:
                        # Remove invalid proposal symmetrically
                        tmpPreferences[i].remove(j)
                        tmpPreferences[j].remove(i)
                if frequency[int(j-1)] == 2:
                    if len(str(proposals[j])) < 4:
                        del queu[0]
                        proposals[j] = proposals[j],i
                        break 
                    else:
                        # Remove invalid proposal symmetrically
                        tmpPreferences[i].remove(j)
                        tmpPreferences[j].remove(i)
        inputList["preferences"] = copy.deepcopy(tmpPreferences)
    return (proposals, inputList)
##############################################################################

def step2(proposals, inputList):
    tmpPreferences = copy.deepcopy(inputList["preferences"])
    for i in inputList["preferences"]:
        if frequency[int(i)-1] == 2 and len(str(proposals[i])) >= 10 and proposals[i] != None:
        #  Remove the right hand side of the preferred element
            proposalIndex = tmpPreferences[i].index(proposals[i][-1])
            rem = tmpPreferences[i][proposalIndex + 1:]
            tmpPreferences[i] = tmpPreferences[i][:proposalIndex + 1]
        # Remove all other instances of the given element
            for j in inputList["preferences"]:
                if rem and j in rem:
                    tmpPreferences[i].remove(j)
        elif frequency[int(i)-1] == 1 and len(str(proposals[i])) <= 4 and proposals[i] != None:
            proposalIndex = tmpPreferences[i].index(proposals[i])
            rem = tmpPreferences[i][proposalIndex + 1:]
            tmpPreferences[i] = tmpPreferences[i][:proposalIndex + 1]
        # Remove all other instances of the given element
            for j in inputList["preferences"]:
                if rem and j in rem:
                    tmpPreferences[j].remove(i)

    # for i in inputList["preferences"]:
    #    pass
    tmpPreferences = {"preferences": tmpPreferences}
    return tmpPreferences
##############################################################################
def step3(remaining1):
    queue = [] 
    numProposals = {}
    tmpPreferences = copy.deepcopy(remaining1["preferences"])
    for i in inputList["preferences"]:        
        queue.append(i)
        numProposals[i] = 0
    while not len(queue) == 0:        
        i = queue[0]
        numProposals[i] += 1
        for j in remaining1["preferences"][i]:
            if proposals[j] == None:
                del queue[0]
                proposals[j] = i
                break
            if len(str(proposals[j])) >= 6:
                current_index = remaining1["preferences"][j].index(i)
                if proposals[j][0] != i and proposals[j][1] != i:
                    other_index1 = remaining1["preferences"][j].index(proposals[j][0])
                    other_index2 = remaining1["preferences"][j].index(proposals[j][1])
                    if i == 10.0:
                        del queue[0]
                    if current_index < other_index1 or current_index < other_index2:
                        del queue[0]
                        # Remove old proposal symmetrically
                        tmpPreferences[proposals[j][1]].remove(j)
                        tmpPreferences[j].remove(proposals[j][1])
                        proposals[j] = proposals[j][0],i
                        break
                    else:
                        # Remove invalid proposal symmetrically
                        tmpPreferences[i].remove(j)
                        tmpPreferences[j].remove(i)
            if len(str(proposals[j])) <= 4:
                if frequency[int(j-1)] == 2:
                    if proposals[j] != i:
                        del queue[0]
                        proposals[j] = proposals[j],i
                        break
                    if i==7.0:
                        del queue[0]
                if frequency[int(j-1)] == 1:
                    if proposals[j] != i:
                        current_index = inputList["preferences"][j].index(i)
                        other_index = inputList["preferences"][j].index(proposals[j])                
                        if current_index < other_index:
                            del queue[0]
                            # Remove old proposal symmetrically
                            tmpPreferences[proposals[j]].remove(j)
                            tmpPreferences[j].remove(proposals[j])
                            proposals[j] = i
                            break
                        else:
                            # Remove invalid proposal symmetrically
                            tmpPreferences[i].remove(j)
                            tmpPreferences[j].remove(i)
                
        remaining1["preferences"] = copy.deepcopy(tmpPreferences)
    return (proposals, remaining1)
##############################################################################
def step4(proposals, inputist):
    tmpPreferences = copy.deepcopy(inputist["preferences"])
    for i in inputist["preferences"]:
        if frequency[int(i)-1] == 2 and len(str(proposals[i])) >= 10 and proposals[i] != None:
        #  Remove the right hand side of the preferred element
            proposalIndex1 = tmpPreferences[i].index(proposals[i][0])
            proposalIndex2 = tmpPreferences[i].index(proposals[i][-1])
            proposalIndex = max(proposalIndex1, proposalIndex2)
            rem = tmpPreferences[i][proposalIndex + 1:]
            #tmpPreferences[i] = tmpPreferences[i][:proposalIndex + 1]
        # Remove all other instances of the given element
            for j in inputist["preferences"]:
                if rem and j in rem:
                    tmpPreferences[i].remove(j)
                    tmpPreferences[j].remove(i)
        if frequency[int(i)-1] == 1 and len(str(proposals[i])) <= 4 and proposals[i] != None:
            proposalIndex = tmpPreferences[i].index(proposals[i])
            rem = tmpPreferences[i][proposalIndex + 1:]
            #tmpPreferences[i] = tmpPreferences[i][:proposalIndex + 1]
        # Remove all other instances of the given element
            for j in inputist["preferences"]:
                if rem and j in rem:
                    tmpPreferences[j].remove(i)

    # for i in inputList["preferences"]:
    #    pass
    tmpPreferences = {"preferences": tmpPreferences}
    return tmpPreferences
##############################################################################
def step5(remaining2):
    tmpPreferences = copy.deepcopy(remaining2["preferences"])
    for j in remaining2["preferences"]:
        if remaining2["preferences"][j]:
            if frequency[int(j-1)] == 2:
                if len(tmpPreferences[j])>2:
                    i = tmpPreferences[j][-1]
                    tmpPreferences[i].remove(j)
                    tmpPreferences[j].remove(i)
            if frequency[int(j-1)] == 1:    
                while(len(tmpPreferences[j])>1):
                    i = tmpPreferences[j][-1]
                    tmpPreferences[i].remove(j)
                    tmpPreferences[j].remove(i)
    for k in remaining2["preferences"]:
        if frequency[int(k-1)] == 2:
            if len(tmpPreferences[k])>2:
                i = tmpPreferences[k][-1]
                tmpPreferences[i].remove(k)
                tmpPreferences[k].remove(i)
    return tmpPreferences
##############################################################################
proposals, inputList = step1(input_data)
remaining1 = step2(proposals, inputList)
prop, ilist = step3(remaining1)
remaining2 = step4(prop, ilist)
results = step5(remaining2)
print(results)
##############################################################################

