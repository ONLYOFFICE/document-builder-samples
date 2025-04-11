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
    with open(os.path.join(resources_dir, 'data/financial_system_response.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new docx file
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Document.DOCX)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    document = api.Call('GetDocument')

    # DOCUMENT HEADER
    paragraph = document.Call('GetElement', 0)
    addTextToParagraph(paragraph, 'Annual Report for %d' % data['year'], 44, True, 'center')

    # FINANCIAL section
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Financial performance', 32, True)
    document.Call('Push', paragraph)
    # quarterly data
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Quarterly data:', 24)
    document.Call('Push', paragraph)
    # chart
    paragraph = api.Call('CreateParagraph')
    chart_names = ['revenue', 'expenses', 'net_profit']
    chart_data = [[entry[key] for entry in data['financials']['quarterly_data']] for key in chart_names]
    chart = api.Call('CreateChart', 'lineNormal', chart_data, ['Revenue', 'Expenses', 'Net Profit'], ['Q1', 'Q2', 'Q3', 'Q4'])
    chart.Call('SetSize', 170 * 36000, 90 * 36000)
    paragraph.Call('AddDrawing', chart)
    document.Call('Push', paragraph)
    # expenses
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Expenses:', 24)
    document.Call('Push', paragraph)
    # pie chart
    paragraph = api.Call('CreateParagraph')
    r_d_expenses = data['financials']['r_d_expenses']
    marketing_expenses = data['financials']['marketing_expenses']
    total_expenses = data['financials']['total_expenses']
    chart = api.Call('CreateChart', 'pie', [[r_d_expenses, marketing_expenses, total_expenses - (r_d_expenses + marketing_expenses)]], [], ['Research and Development', 'Marketing', 'Other'])
    chart.Call('SetSize', 170 * 36000, 90 * 36000)
    paragraph.Call('AddDrawing', chart)
    document.Call('Push', paragraph)
    # year totals
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Year total numbers:', 24)
    document.Call('Push', paragraph)
    # table
    table = createTable(api, 2, 3)
    fillTableHeaders(table, ['Total revenue', 'Total expenses', 'Total net profit'], 22)
    paragraph = getTableCellParagraph(table, 1, 0)
    addTextToParagraph(paragraph, str(data['financials']['total_revenue']), 22)
    paragraph = getTableCellParagraph(table, 1, 1)
    addTextToParagraph(paragraph, str(data['financials']['total_expenses']), 22)
    paragraph = getTableCellParagraph(table, 1, 2)
    addTextToParagraph(paragraph, str(data['financials']['net_profit']), 22)
    document.Call('Push', table)

    # ACHIEVEMENTS section
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Achievements this year', 32, True)
    document.Call('Push', paragraph)
    # list
    createNumbering(api, data['achievements'], 'numbered', 22)

    # PLANS section
    # header
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Plans for the next year', 32, True)
    document.Call('Push', paragraph)
    # projects
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Projects:', 24)
    document.Call('Push', paragraph)
    # table
    projects = data['plans']['projects']
    table = createTable(api, len(projects) + 1, 2)
    fillTableHeaders(table, ['Name', 'Deadline'], 22)
    fillTableBody(table, projects, ['name', 'deadline'], 22)
    document.Call('Push', table)
    # financial goals
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Financial goals:', 24)
    document.Call('Push', paragraph)
    # table
    goals = data['plans']['financial_goals']
    table = createTable(api, len(goals) + 1, 2)
    fillTableHeaders(table, ['Goal', 'Value'], 22)
    fillTableBody(table, goals, ['goal', 'value'], 22)
    document.Call('Push', table)
    # marketing initiatives
    paragraph = api.Call('CreateParagraph')
    addTextToParagraph(paragraph, 'Marketing initiatives:', 24)
    document.Call('Push', paragraph)
    # list
    createNumbering(api, data['plans']['marketing_initiatives'], 'bullet', 22)

    # save and close
    result_path = os.getcwd() + '/result.docx'
    builder.SaveFile(docbuilder.FileTypes.Document.DOCX, result_path)
    builder.CloseFile()
