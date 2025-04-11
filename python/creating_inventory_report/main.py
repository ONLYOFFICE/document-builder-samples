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

if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/ims_response.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new xlsx file
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Spreadsheet.XLSX)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    worksheet = api.Call('GetActiveSheet')

    # fill table headers
    worksheet.Call('GetRangeByNumber', 0, 0).Call('SetValue', 'Item')
    worksheet.Call('GetRangeByNumber', 0, 1).Call('SetValue', 'Quantity')
    worksheet.Call('GetRangeByNumber', 0, 2).Call('SetValue', 'Status')
    # make headers bold
    start_cell = worksheet.Call('GetRangeByNumber', 0, 0)
    end_cell = worksheet.Call('GetRangeByNumber', 0, 2)
    worksheet.Call('GetRange', start_cell, end_cell).Call('SetBold', True)
    # fill table data
    inventory = data['inventory']
    for i, entry in enumerate(inventory):
        cell = worksheet.Call('GetRangeByNumber', i + 1, 0)
        cell.Call('SetValue', str(entry['item']))
        cell = worksheet.Call('GetRangeByNumber', i + 1, 1)
        cell.Call('SetValue', str(entry['quantity']))
        cell = worksheet.Call('GetRangeByNumber', i + 1, 2)
        status = str(entry['status'])
        cell.Call('SetValue', status)
        # fill cell with color corresponding to status
        if status == 'In Stock':
            cell.Call('SetFillColor', api.Call('CreateColorFromRGB', 0, 194, 87))
        elif status == 'Reserved':
            cell.Call('SetFillColor', api.Call('CreateColorFromRGB', 255, 255, 0))
        else:
            cell.Call('SetFillColor', api.Call('CreateColorFromRGB', 255, 79, 79))
    # tweak cells width
    worksheet.Call('GetRange', 'A1').Call('SetColumnWidth', 40)
    worksheet.Call('GetRange', 'C1').Call('SetColumnWidth', 15)

    # save and close
    result_path = os.getcwd() + '/result.xlsx'
    builder.SaveFile(docbuilder.FileTypes.Spreadsheet.XLSX, result_path)
    builder.CloseFile()
