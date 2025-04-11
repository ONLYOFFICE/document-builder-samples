'''
(c) Copyright Ascensio System SIA 2025

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
'''

# import docbuilder
import sys
sys.path.append('../../out/python')
import constants
sys.path.append(constants.BUILDER_DIR)
import docbuilder
# import other standard libraries
import json
import os

def addTextToParagraph(paragraph, text, font_size, is_bold=False, jc='left'):
    paragraph.Call('AddText', text)
    paragraph.Call('SetFontSize', font_size)
    paragraph.Call('SetBold', is_bold)
    paragraph.Call('SetJc', jc)

def createTable(api, rows, cols, border_color=200):
    # create table
    table = api.Call('CreateTable', cols, rows)
    # set table properties
    table.Call('SetWidth', 'percent', 100)
    table.Call('SetTableCellMarginTop', 200)
    table.Call('GetRow', 0).Call('SetBackgroundColor', 245, 245, 245)
    # set table borders
    table.Call('SetTableBorderTop', 'single', 4, 0, border_color, border_color, border_color)
    table.Call('SetTableBorderBottom', 'single', 4, 0, border_color, border_color, border_color)
    table.Call('SetTableBorderLeft', 'single', 4, 0, border_color, border_color, border_color)
    table.Call('SetTableBorderRight', 'single', 4, 0, border_color, border_color, border_color)
    table.Call('SetTableBorderInsideV', 'single', 4, 0, border_color, border_color, border_color)
    table.Call('SetTableBorderInsideH', 'single', 4, 0, border_color, border_color, border_color)
    return table

def getTableCellParagraph(table, row, col):
    return table.Call('GetCell', row, col).Call('GetContent').Call('GetElement', 0)

def fillTableHeaders(table, data, font_size):
    for i in range(len(data)):
        paragraph = getTableCellParagraph(table, 0, i)
        addTextToParagraph(paragraph, data[i], font_size, True)

def fillTableBody(table, data, keys, font_size, start_row=1):
    for row in range(len(data)):
        for col, key in enumerate(keys):
            paragraph = getTableCellParagraph(table, row + start_row, col)
            addTextToParagraph(paragraph, str(data[row][key]), font_size)

def createNumbering(api, data, numbering_type, font_size):
    document = api.Call('GetDocument')
    numbering = document.Call('CreateNumbering', numbering_type)
    numbering_level = numbering.Call('GetLevel', 0)
    for entry in data:
        paragraph = api.Call('CreateParagraph')
        paragraph.Call('SetNumbering', numbering_level)
        addTextToParagraph(paragraph, str(entry), font_size)
        document.Call('Push', paragraph)
    # return the last paragraph in numbering
    return paragraph

if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/hrms_response.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new docx file
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Document.DOCX)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    document = api.Call('GetDocument')

    # TITLE PAGE
    # header
    paragraph = document.Call('GetElement', 0)
    addTextToParagraph(paragraph, 'Employee Development Plan for 2024', 48, True, 'center')
    paragraph.Call('SetSpacingBefore', 5000)
    paragraph.Call('SetSpacingAfter', 500)
    # employee name
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, str(data['employee']['name']), 36, False, 'center')
    document.Call('Push', paragraph)
    # employee position and department
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Position: %s\nDepartment: %s' % (str(data['employee']['position']), str(data['employee']['department'])), 24, False, 'center')
    paragraph.Call('AddPageBreak')
    document.Call('Push', paragraph)

    # COMPETENCIES SECION
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Competencies', 32, True)
    document.Call('Push', paragraph)
    # technical skills sub-header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Technical skills:', 24)
    document.Call('Push', paragraph)
    # technical skills table
    technical_skills = data['competencies']['technical_skills']
    table = createTable(api, len(technical_skills) + 1, 2)
    fillTableHeaders(table, ['Skill', 'Level'], 22)
    fillTableBody(table, technical_skills, ['name', 'level'], 22)
    document.Call('Push', table)
    # soft skills sub-header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Soft skills:', 24)
    document.Call('Push', paragraph)
    # soft skills table
    soft_skills = data['competencies']['soft_skills']
    table = createTable(api, len(soft_skills) + 1, 2)
    fillTableHeaders(table, ['Skill', 'Level'], 22)
    fillTableBody(table, soft_skills, ['name', 'level'], 22)
    document.Call('Push', table)

    # DEVELOPMENT AREAS section
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Development areas', 32, True)
    document.Call('Push', paragraph)
    # list
    createNumbering(api, data['development_areas'], 'numbered', 22)

    # GOALS section
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Goals for next year', 32, True)
    document.Call('Push', paragraph)
    # numbering
    paragraph = createNumbering(api, data['goals_next_year'], 'numbered', 22)
    # add a page break after the last paragraph
    paragraph.Call('AddPageBreak')

    # RESOURCES section
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Recommended resources', 32, True)
    document.Call('Push', paragraph)
    # table
    resources = data['resources']
    table = createTable(api, len(resources) + 1, 3)
    fillTableHeaders(table, ['Name', 'Provider', 'Duration'], 22)
    fillTableBody(table, resources, ['name', 'provider', 'duration'], 22)
    document.Call('Push', table)

    # FEEDBACK section
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Feedback', 32, True)
    document.Call('Push', paragraph)
    # manager's feedback
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Manager\'s feedback:', 24, False)
    document.Call('Push', paragraph)
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, '_' * 280, 24, False)
    document.Call('Push', paragraph)
    # employees's feedback
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Employee\'s feedback:', 24, False)
    document.Call('Push', paragraph)
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, '_' * 280, 24, False)
    document.Call('Push', paragraph)

    # save and close
    result_path = os.getcwd() + '/result.docx'
    builder.SaveFile(docbuilder.FileTypes.Document.DOCX, result_path)
    builder.CloseFile()
