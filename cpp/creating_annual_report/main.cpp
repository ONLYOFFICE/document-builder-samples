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
const wchar_t* resultPath = L"result.docx";

// Helper functions
void addTextToParagraph(CValue oParagraph, string text, int fontSize, bool isBold = false, string jc = "left")
{
    oParagraph.Call("AddText", text.c_str());
    oParagraph.Call("SetFontSize", fontSize);
    oParagraph.Call("SetBold", isBold);
    oParagraph.Call("SetJc", jc.c_str());
}

CValue createTable(CValue oApi, int rows, int cols, int borderColor = 200)
{
    // create table
    CValue oTable = oApi.Call("CreateTable", cols, rows);
    // set table properties;
    oTable.Call("SetWidth", "percent", 100);
    oTable.Call("SetTableCellMarginTop", 200);
    oTable.Call("GetRow", 0).Call("SetBackgroundColor", 245, 245, 245);
    // set table borders;
    oTable.Call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
    oTable.Call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
    oTable.Call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
    oTable.Call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
    oTable.Call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
    oTable.Call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
    return oTable;
}

CValue getTableCellParagraph(CValue oTable, int row, int col)
{
    return oTable.Call("GetCell", row, col).Call("GetContent").Call("GetElement", 0);
}

void fillTableHeaders(CValue oTable, const vector<string>& data, int fontSize)
{
    for (int i = 0; i < data.size(); i++)
    {
        CValue oParagraph = getTableCellParagraph(oTable, 0, i);
        addTextToParagraph(oParagraph, data[i], fontSize, true);
    }
}

void fillTableBody(CValue oTable, const json& data, const vector<string>& keys, int fontSize, int startRow = 1)
{
    for (int row = 0; row < data.size(); row++)
    {
        for (int col = 0; col < keys.size(); col++)
        {
            CValue oParagraph = getTableCellParagraph(oTable, row + startRow, col);
            const string& key = keys[col];
            addTextToParagraph(oParagraph, data[row][key].get<string>(), fontSize);
        }
    }
}

CValue createNumbering(CValue oApi, const json& data, string numberingType, int fontSize)
{
    CValue oDocument = oApi.Call("GetDocument");
    CValue oNumbering = oDocument.Call("CreateNumbering", numberingType.c_str());
    CValue oNumberingLevel = oNumbering.Call("GetLevel", 0);

    CValue oParagraph;
    for (const auto& entry : data)
    {
        oParagraph = oApi.Call("CreateParagraph");
        oParagraph.Call("SetNumbering", oNumberingLevel);
        addTextToParagraph(oParagraph, entry.get<string>().c_str(), fontSize);
        oDocument.Call("Push", oParagraph);
    }
    // return the last oParagraph in numbering
    return oParagraph;
}

CValue createStringArray(const vector<string>& values)
{
    CValue arrResult = CValue::CreateArray((int)values.size());
    for (int i = 0; i < values.size(); i++)
    {
        arrResult[i] = values[i].c_str();
    }

    return arrResult;
}

CValue createIntegerArray(const vector<int>& values)
{
    CValue arrResult = CValue::CreateArray((int)values.size());
    for (int i = 0; i < values.size(); i++)
    {
        arrResult[i] = values[i];
    }

    return arrResult;
}

// Main function
int main()
{
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/financial_system_response.json";
    ifstream fs(jsonPath);
    json data = json::parse(fs);

    // init docbuilder and create new docx file
    CDocBuilder::Initialize(workDir);
    CDocBuilder oBuilder;
    oBuilder.CreateFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX);

    CContext oContext = oBuilder.GetContext();
    CValue oGlobal = oContext.GetGlobal();
    CValue oApi = oGlobal["Api"];
    CValue oDocument = oApi.Call("GetDocument");

    // DOCUMENT HEADER
    CValue oParagraph = oDocument.Call("GetElement", 0);
    addTextToParagraph(oParagraph, "Annual Report for " + to_string(data["year"].get<int>()), 44, true, "center");

    // FINANCIAL section
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Financial performance", 32, true);
    oDocument.Call("Push", oParagraph);
    // quarterly data
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Quarterly data:", 24);
    oDocument.Call("Push", oParagraph);
    // chart
    oParagraph = oApi.Call("CreateParagraph");
    vector<string> chartKeys = { "revenue", "expenses", "net_profit" };
    const json& quarterlyData = data["financials"]["quarterly_data"];
    CValue arrChartData = CValue::CreateArray((int)chartKeys.size());
    for (int i = 0; i < chartKeys.size(); i++)
    {
        arrChartData[i] = CValue::CreateArray((int)quarterlyData.size());
        for (int j = 0; j < quarterlyData.size(); j++)
        {
            arrChartData[i][j] = quarterlyData[j][chartKeys[i]].get<int>();
        }
    }
    CValue arrChartNames = createStringArray({ "Revenue", "Expenses", "Net Profit" });
    CValue arrHorValues = createStringArray({ "Q1", "Q2", "Q3", "Q4" });
    CValue oChart = oApi.Call("CreateChart", "lineNormal", arrChartData, arrChartNames, arrHorValues);
    oChart.Call("SetSize", 170 * 36000, 90 * 36000);
    oParagraph.Call("AddDrawing", oChart);
    oDocument.Call("Push", oParagraph);
    // expenses
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Expenses:", 24);
    oDocument.Call("Push", oParagraph);
    // pie chart
    oParagraph = oApi.Call("CreateParagraph");
    int rdExpenses = data["financials"]["r_d_expenses"].get<int>();
    int marketingExpenses = data["financials"]["marketing_expenses"].get<int>();
    int totalExpenses = data["financials"]["total_expenses"];
    arrChartData = CValue::CreateArray(1);
    arrChartData[0] = createIntegerArray({ rdExpenses, marketingExpenses, totalExpenses - (rdExpenses + marketingExpenses) });
    arrChartNames = createStringArray({ "Research and Development", "Marketing", "Other" });
    oChart = oApi.Call("CreateChart", "pie", arrChartData, CValue::CreateArray(0), arrChartNames);
    oChart.Call("SetSize", 170 * 36000, 90 * 36000);
    oParagraph.Call("AddDrawing", oChart);
    oDocument.Call("Push", oParagraph);
    // year totals
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Year total numbers:", 24);
    oDocument.Call("Push", oParagraph);
    // table
    CValue oTable = createTable(oApi, 2, 3);
    fillTableHeaders(oTable, { "Total revenue", "Total expenses", "Total net profit" }, 22);
    oParagraph = getTableCellParagraph(oTable, 1, 0);
    addTextToParagraph(oParagraph, to_string(data["financials"]["total_revenue"].get<int>()), 22);
    oParagraph = getTableCellParagraph(oTable, 1, 1);
    addTextToParagraph(oParagraph, to_string(data["financials"]["total_expenses"].get<int>()), 22);
    oParagraph = getTableCellParagraph(oTable, 1, 2);
    addTextToParagraph(oParagraph, to_string(data["financials"]["net_profit"].get<int>()), 22);
    oDocument.Call("Push", oTable);

    // ACHIEVEMENTS section
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Achievements this year", 32, true);
    oDocument.Call("Push", oParagraph);
    // list
    createNumbering(oApi, data["achievements"], "numbered", 22);

    // PLANS section
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Plans for the next year", 32, true);
    oDocument.Call("Push", oParagraph);
    // projects
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Projects:", 24);
    oDocument.Call("Push", oParagraph);
    // table
    const json& projects = data["plans"]["projects"];
    oTable = createTable(oApi, (int)projects.size() + 1, 2);
    fillTableHeaders(oTable, { "Name", "Deadline" }, 22);
    fillTableBody(oTable, projects, { "name", "deadline" }, 22);
    oDocument.Call("Push", oTable);
    // financial goals
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Financial goals:", 24);
    oDocument.Call("Push", oParagraph);
    // table
    const json& goals = data["plans"]["financial_goals"];
    oTable = createTable(oApi, (int)goals.size() + 1, 2);
    fillTableHeaders(oTable, { "Goal", "Value" }, 22);
    fillTableBody(oTable, goals, { "goal", "value" }, 22);
    oDocument.Call("Push", oTable);
    // marketing initiatives
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Marketing initiatives:", 24);
    oDocument.Call("Push", oParagraph);
    // list
    createNumbering(oApi, data["plans"]["marketing_initiatives"], "bullet", 22);

    // save and close
    oBuilder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    oBuilder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
