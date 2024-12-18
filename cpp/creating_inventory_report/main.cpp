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
    CDocBuilder oBuilder;
    oBuilder.CreateFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX);

    CContext oContext = oBuilder.GetContext();
    CValue oGlobal = oContext.GetGlobal();
    CValue oApi = oGlobal["Api"];
    CValue oWorksheet = oApi.Call("GetActiveSheet");

    // fill table headers
    oWorksheet.Call("GetRangeByNumber", 0, 0).Call("SetValue", "Item");
    oWorksheet.Call("GetRangeByNumber", 0, 1).Call("SetValue", "Quantity");
    oWorksheet.Call("GetRangeByNumber", 0, 2).Call("SetValue", "Status");
    // make headers bold
    CValue oStartCell = oWorksheet.Call("GetRangeByNumber", 0, 0);
    CValue oEndCell = oWorksheet.Call("GetRangeByNumber", 0, 2);
    oWorksheet.Call("GetRange", oStartCell, oEndCell).Call("SetBold", true);
    // fill table data
    const json& inventory = data["inventory"];
    for (int i = 0; i < inventory.size(); i++)
    {
        const json& entry = inventory[i];
        CValue oCell = oWorksheet.Call("GetRangeByNumber", i + 1, 0);
        oCell.Call("SetValue", entry["item"].get<string>().c_str());
        oCell = oWorksheet.Call("GetRangeByNumber", i + 1, 1);
        oCell.Call("SetValue", to_string(entry["quantity"].get<int>()).c_str());
        oCell = oWorksheet.Call("GetRangeByNumber", i + 1, 2);
        string status = entry["status"].get<string>();
        oCell.Call("SetValue", status.c_str());
        // fill cell with color corresponding to status
        if (status == "In Stock")
            oCell.Call("SetFillColor", oApi.Call("CreateColorFromRGB", 0, 194, 87));
        else if (status == "Reserved")
            oCell.Call("SetFillColor", oApi.Call("CreateColorFromRGB", 255, 255, 0));
        else
            oCell.Call("SetFillColor", oApi.Call("CreateColorFromRGB", 255, 79, 79));
    }
    // tweak cells width
    oWorksheet.Call("GetRange", "A1").Call("SetColumnWidth", 40);
    oWorksheet.Call("GetRange", "C1").Call("SetColumnWidth", 15);

    // save and close
    oBuilder.SaveFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX, resultPath);
    oBuilder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
