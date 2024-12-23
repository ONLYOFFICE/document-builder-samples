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
#include "resources/utils/utils.h"

using namespace std;
using namespace NSDoctRenderer;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.xlsx";

// Helper functions
void CheckCell(CValue worksheet, wstring cell, int row, int col)
{
    if (cell.find('#') != std::wstring::npos)
    {
        wstring commentMsg = L"Error: " + cell;
        CValue errorCell = worksheet.Call("GetRangeByNumber", row, col);
        errorCell.Call("AddComment", commentMsg.c_str());
    }
}

// Main function
int main()
{
    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;

    // Open file and get context
    wstring templatePath = NSUtils::GetResourcesDirectory() + L"/docs/spreadsheet_with_errors.xlsx";
    builder.OpenFile(templatePath.c_str(), L"");
    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];

    // Find and comment formula errors
    CValue worksheet = api.Call("GetActiveSheet");
    CValue range = worksheet.Call("GetUsedRange");
    CValue data = range.Call("GetValue");

    for (int row = 0; row < (int)data.GetLength(); row++)
    {
        for (int col = 0; col < (int)data[0].GetLength(); col++)
        {
            CheckCell(worksheet, data[row][col].ToString().c_str(), row, col);
        }
    }

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
