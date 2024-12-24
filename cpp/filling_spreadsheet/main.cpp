/**
 *
 * (c) Copyright Ascensio System SIA 2024
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

#include <string>
#include "common.h"
#include "docbuilder.h"

#include "out/cpp/builder_path.h"

using namespace std;
using namespace NSDoctRenderer;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.xlsx";

// Main function
int main()
{
    string data[9][4] = {
        { "Id", "Product", "Price", "Available" },
        { "1001", "Item A", "12.2", "true" },
        { "1002", "Item B", "18.8", "true" },
        { "1003", "Item C", "70.1", "false" },
        { "1004", "Item D", "60.6", "true" },
        { "1005", "Item E", "32.6", "true" },
        { "1006", "Item F", "28.3", "false" },
        { "1007", "Item G", "11.1", "false" },
        { "1008", "Item H", "41.4", "true" }
    };

    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;
    builder.CreateFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX);

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];

    // Get current worksheet
    CValue worksheet = api.Call("GetActiveSheet");

    // Create CValue array from data
    int rowsLen = sizeof data / sizeof data[0];
    int colsLen = sizeof data[0] / sizeof(string);
    CValue array = context.CreateArray(rowsLen);

    for (int row = 0; row < rowsLen; row++)
    {
        CValue arrayCol = context.CreateArray(colsLen);

        for (int col = 0; col < colsLen; col++)
        {
            arrayCol[col] = data[row][col].c_str();
        }
        array[row] = arrayCol;
    }

    // First cell in the range (A1) is equal to (0,0)
    CValue startCell = worksheet.Call("GetRangeByNumber", 0, 0);

    // Last cell in the range is equal to array length -1
    CValue endCell = worksheet.Call("GetRangeByNumber", array.GetLength() - 1, array[0].GetLength() - 1);
    worksheet.Call("GetRange", startCell, endCell).Call("SetValue", array);

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
