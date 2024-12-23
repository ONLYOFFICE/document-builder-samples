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

#include <fstream>
#include <string>

#include "common.h"
#include "docbuilder.h"

#include "out/cpp/builder_path.h"
#include "resources/utils/utils.h"
#include "resources/utils/json/json.hpp"

using namespace std;
using namespace NSDoctRenderer;
using json = nlohmann::json;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.xlsx";

// Main function
int main()
{
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/ims_response.json";
    ifstream fs(jsonPath);
    json data = json::parse(fs);

    // init docbuilder and create new xlsx file
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;
    builder.CreateFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX);

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];
    CValue worksheet = api.Call("GetActiveSheet");

    // fill table headers
    worksheet.Call("GetRangeByNumber", 0, 0).Call("SetValue", "Item");
    worksheet.Call("GetRangeByNumber", 0, 1).Call("SetValue", "Quantity");
    worksheet.Call("GetRangeByNumber", 0, 2).Call("SetValue", "Status");
    // make headers bold
    CValue startCell = worksheet.Call("GetRangeByNumber", 0, 0);
    CValue endCell = worksheet.Call("GetRangeByNumber", 0, 2);
    worksheet.Call("GetRange", startCell, endCell).Call("SetBold", true);
    // fill table data
    const json& inventory = data["inventory"];
    for (int i = 0; i < inventory.size(); i++)
    {
        const json& entry = inventory[i];
        CValue cell = worksheet.Call("GetRangeByNumber", i + 1, 0);
        cell.Call("SetValue", entry["item"].get<string>().c_str());
        cell = worksheet.Call("GetRangeByNumber", i + 1, 1);
        cell.Call("SetValue", to_string(entry["quantity"].get<int>()).c_str());
        cell = worksheet.Call("GetRangeByNumber", i + 1, 2);
        string status = entry["status"].get<string>();
        cell.Call("SetValue", status.c_str());
        // fill cell with color corresponding to status
        if (status == "In Stock")
            cell.Call("SetFillColor", api.Call("CreateColorFromRGB", 0, 194, 87));
        else if (status == "Reserved")
            cell.Call("SetFillColor", api.Call("CreateColorFromRGB", 255, 255, 0));
        else
            cell.Call("SetFillColor", api.Call("CreateColorFromRGB", 255, 79, 79));
    }
    // tweak cells width
    worksheet.Call("GetRange", "A1").Call("SetColumnWidth", 40);
    worksheet.Call("GetRange", "C1").Call("SetColumnWidth", 15);

    // save and close
    builder.SaveFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
