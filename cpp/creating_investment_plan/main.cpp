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
#include <vector>

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

// Helper functions
CValue createColumnData(const vector<string>& data)
{
    CValue arrColumnData = CValue::CreateArray((int)data.size());
    for (int i = 0; i < (int)data.size(); i++)
    {
        CValue arrRow = CValue::CreateArray(1);
        arrRow[0] = data[i].c_str();
        arrColumnData[i] = arrRow;
    }
    return arrColumnData;
}

// Main function
int main()
{
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/investment_data.json";
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

    // initialize financial data from JSON
    int initAmount = data["initial_amount"].get<int>();
    double rate = data["return_rate"].get<double>();
    int term = data["term"].get<int>();

    // fill years
    CValue startCell = worksheet.Call("GetRangeByNumber", 1, 0);
    CValue endCell = worksheet.Call("GetRangeByNumber", term + 1, 0);
    vector<string> years((size_t)term + 1);
    for (int year = 0; year <= term; year++)
    {
        years[year] = to_string(year);
    }
    worksheet.Call("GetRange", startCell, endCell).Call("SetValue", createColumnData(years));

    // fill initial amount
    worksheet.Call("GetRangeByNumber", 1, 1).Call("SetValue", initAmount);
    // fill remaining cells
    startCell = worksheet.Call("GetRangeByNumber", 2, 1);
    endCell = worksheet.Call("GetRangeByNumber", term + 1, 1);
    vector<string> amounts(term);
    for (int year = 0; year < term; year++)
    {
        amounts[year] = "=$B$2*POWER((1+" + to_string(rate) + "),A" + to_string(year + 3) + ")";
    }
    worksheet.Call("GetRange", startCell, endCell).Call("SetValue", createColumnData(amounts));

    // create chart
    string chartDataRange = "Sheet1!$A$1:$B$" + to_string(term + 2);
    CValue chart = worksheet.Call("AddChart", chartDataRange.c_str(), false, "lineNormal", 2, 135.38 * 36000, 81.28 * 36000);
    chart.Call("SetPosition", 3, 0, 2, 0);
    chart.Call("SetTitle", "Capital Growth Over Time", 22);
    CValue color = api.Call("CreateRGBColor", 134, 134, 134);
    CValue fill = api.Call("CreateSolidFill", color);
    CValue stroke = api.Call("CreateStroke", 1, fill);
    chart.Call("SetMinorVerticalGridlines", stroke);
    chart.Call("SetMajorHorizontalGridlines", stroke);
    // fill table headers
    worksheet.Call("GetRangeByNumber", 0, 0).Call("SetValue", "Year");
    worksheet.Call("GetRangeByNumber", 0, 1).Call("SetValue", "Amount");

    // save and close
    builder.SaveFile(OFFICESTUDIO_FILE_SPREADSHEET_XLSX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
