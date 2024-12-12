'''
(c) Copyright Ascensio System SIA 2024

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

if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/investment_data.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new xlsx file
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Spreadsheet.XLSX)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    worksheet = api.Call('GetActiveSheet')

    # initialize financial data from JSON
    init_amount = data['initial_amount']
    rate = data['return_rate']
    term = data['term']

    # fill years
    start_cell = worksheet.Call('GetRangeByNumber', 1, 0)
    end_cell = worksheet.Call('GetRangeByNumber', term + 1, 0)
    worksheet.Call('GetRange', start_cell, end_cell).Call('SetValue', [[i] for i in range(term + 1)])
    # fill initial amount
    worksheet.Call('GetRangeByNumber', 1, 1).Call('SetValue', init_amount)
    # fill remaining cells
    start_cell = worksheet.Call('GetRangeByNumber', 2, 1)
    end_cell = worksheet.Call('GetRangeByNumber', term + 1, 1)
    worksheet.Call('GetRange', start_cell, end_cell).Call('SetValue', [['=$B$2*POWER((1+0.12),A%d)' % (i + 1)] for i in range(2, term + 2)])

    # create chart
    chart = worksheet.Call('AddChart', 'Sheet1!$A$1:$B$%d' % (term + 2), False, 'lineNormal', 2, 135.38 * 36000, 81.28 * 36000)
    chart.Call('SetPosition', 3, 0, 2, 0)
    chart.Call('SetTitle', 'Capital Growth Over Time', 22)
    color = api.Call('CreateRGBColor', 134, 134, 134)
    fill = api.Call('CreateSolidFill', color)
    stroke = api.Call('CreateStroke', 1, fill)
    chart.Call('SetMinorVerticalGridlines', stroke)
    chart.Call('SetMajorHorizontalGridlines', stroke)
    # fill table headers
    worksheet.Call('GetRangeByNumber', 0, 0).Call('SetValue', 'Year')
    worksheet.Call('GetRangeByNumber', 0, 1).Call('SetValue', 'Amount')

    # save and close
    result_path = os.getcwd() + '/result.xlsx'
    builder.SaveFile(docbuilder.FileTypes.Spreadsheet.XLSX, result_path)
    builder.CloseFile()
