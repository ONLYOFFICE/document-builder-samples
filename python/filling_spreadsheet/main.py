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

import os
import sys
sys.path.append('../../out/python')
import constants
sys.path.append(constants.BUILDER_DIR)
import docbuilder

if __name__ == '__main__':
    data = [
        [ 'Id', 'Product', 'Price', 'Available' ],
        [ '1001', 'Item A', "12.2", 'true' ],
        [ '1002', 'Item B', "18.8", 'true' ],
        [ '1003', 'Item C', "70.1", 'false' ],
        [ '1004', 'Item D', "60.6", 'true' ],
        [ '1005', 'Item E', "32.6", 'true' ],
        [ '1006', 'Item F', "28.3", 'false' ],
        [ '1007', 'Item G', "11.1", 'false' ],
        [ '1008', 'Item H', "41.4", 'true' ]
    ]

    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Spreadsheet.XLSX)

    context = builder.GetContext()
    globalObj = context.GetGlobal()
    api = globalObj['Api']

    # Get current worksheet
    worksheet = api.Call('GetActiveSheet')

    # Create array from data
    array = docbuilder.CDocBuilderValue(data)

    # First cell in the range (A1) is equal to (0,0)
    startCell = worksheet.Call('GetRangeByNumber', 0, 0)

    # Last cell in the range is equal to array length -1
    endCell = worksheet.Call('GetRangeByNumber', len(data) - 1, len(data[0]) - 1)
    worksheet.Call('GetRange', startCell, endCell).Call('SetValue', array)

    # Save and close
    resultPath = os.getcwd() + '/result.xlsx'
    builder.SaveFile(docbuilder.FileTypes.Spreadsheet.XLSX, resultPath)
    builder.CloseFile()
