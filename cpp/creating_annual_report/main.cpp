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
void addTextToParagraph(CValue paragraph, string text, int fontSize, bool isBold = false, string jc = "left")
{
    paragraph.Call("AddText", text.c_str());
    paragraph.Call("SetFontSize", fontSize);
    paragraph.Call("SetBold", isBold);
    paragraph.Call("SetJc", jc.c_str());
}

CValue createTable(CValue api, int rows, int cols, int borderColor = 200)
{
    // create table
    CValue table = api.Call("CreateTable", cols, rows);
    // set table properties;
    table.Call("SetWidth", "percent", 100);
    table.Call("SetTableCellMarginTop", 200);
    table.Call("GetRow", 0).Call("SetBackgroundColor", 245, 245, 245);
    // set table borders;
    table.Call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
    return table;
}

CValue getTableCellParagraph(CValue table, int row, int col)
{
    return table.Call("GetCell", row, col).Call("GetContent").Call("GetElement", 0);
}

void fillTableHeaders(CValue table, const vector<string>& data, int fontSize)
{
    for (int i = 0; i < data.size(); i++)
    {
        CValue paragraph = getTableCellParagraph(table, 0, i);
        addTextToParagraph(paragraph, data[i], fontSize, true);
    }
}

void fillTableBody(CValue table, const json& data, const vector<string>& keys, int fontSize, int startRow = 1)
{
    for (int row = 0; row < data.size(); row++)
    {
        for (int col = 0; col < keys.size(); col++)
        {
            CValue paragraph = getTableCellParagraph(table, row + startRow, col);
            const string& key = keys[col];
            addTextToParagraph(paragraph, data[row][key].get<string>(), fontSize);
        }
    }
}

CValue createNumbering(CValue api, const json& data, string numberingType, int fontSize)
{
    CValue document = api.Call("GetDocument");
    CValue numbering = document.Call("CreateNumbering", numberingType.c_str());
    CValue numberingLevel = numbering.Call("GetLevel", 0);

    CValue paragraph;
    for (const auto& entry : data)
    {
        paragraph = api.Call("CreateParagraph");
        paragraph.Call("SetNumbering", numberingLevel);
        addTextToParagraph(paragraph, entry.get<string>().c_str(), fontSize);
        document.Call("Push", paragraph);
    }
    // return the last paragraph in numbering
    return paragraph;
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
    CDocBuilder builder;
    builder.CreateFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX);

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];
    CValue document = api.Call("GetDocument");

    // DOCUMENT HEADER
    CValue paragraph = document.Call("GetElement", 0);
    addTextToParagraph(paragraph, "Annual Report for " + to_string(data["year"].get<int>()), 44, true, "center");

    // FINANCIAL section
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Financial performance", 32, true);
    document.Call("Push", paragraph);
    // quarterly data
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Quarterly data:", 24);
    document.Call("Push", paragraph);
    // chart
    paragraph = api.Call("CreateParagraph");
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
    CValue chart = api.Call("CreateChart", "lineNormal", arrChartData, arrChartNames, arrHorValues);
    chart.Call("SetSize", 170 * 36000, 90 * 36000);
    paragraph.Call("AddDrawing", chart);
    document.Call("Push", paragraph);
    // expenses
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Expenses:", 24);
    document.Call("Push", paragraph);
    // pie chart
    paragraph = api.Call("CreateParagraph");
    int rdExpenses = data["financials"]["r_d_expenses"].get<int>();
    int marketingExpenses = data["financials"]["marketing_expenses"].get<int>();
    int totalExpenses = data["financials"]["total_expenses"];
    arrChartData = CValue::CreateArray(1);
    arrChartData[0] = createIntegerArray({ rdExpenses, marketingExpenses, totalExpenses - (rdExpenses + marketingExpenses) });
    arrChartNames = createStringArray({ "Research and Development", "Marketing", "Other" });
    chart = api.Call("CreateChart", "pie", arrChartData, CValue::CreateArray(0), arrChartNames);
    chart.Call("SetSize", 170 * 36000, 90 * 36000);
    paragraph.Call("AddDrawing", chart);
    document.Call("Push", paragraph);
    // year totals
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Year total numbers:", 24);
    document.Call("Push", paragraph);
    // table
    CValue table = createTable(api, 2, 3);
    fillTableHeaders(table, { "Total revenue", "Total expenses", "Total net profit" }, 22);
    paragraph = getTableCellParagraph(table, 1, 0);
    addTextToParagraph(paragraph, to_string(data["financials"]["total_revenue"].get<int>()), 22);
    paragraph = getTableCellParagraph(table, 1, 1);
    addTextToParagraph(paragraph, to_string(data["financials"]["total_expenses"].get<int>()), 22);
    paragraph = getTableCellParagraph(table, 1, 2);
    addTextToParagraph(paragraph, to_string(data["financials"]["net_profit"].get<int>()), 22);
    document.Call("Push", table);

    // ACHIEVEMENTS section
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Achievements this year", 32, true);
    document.Call("Push", paragraph);
    // list
    createNumbering(api, data["achievements"], "numbered", 22);

    // PLANS section
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Plans for the next year", 32, true);
    document.Call("Push", paragraph);
    // projects
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Projects:", 24);
    document.Call("Push", paragraph);
    // table
    const json& projects = data["plans"]["projects"];
    table = createTable(api, (int)projects.size() + 1, 2);
    fillTableHeaders(table, { "Name", "Deadline" }, 22);
    fillTableBody(table, projects, { "name", "deadline" }, 22);
    document.Call("Push", table);
    // financial goals
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Financial goals:", 24);
    document.Call("Push", paragraph);
    // table
    const json& goals = data["plans"]["financial_goals"];
    table = createTable(api, (int)goals.size() + 1, 2);
    fillTableHeaders(table, { "Goal", "Value" }, 22);
    fillTableBody(table, goals, { "goal", "value" }, 22);
    document.Call("Push", table);
    // marketing initiatives
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Marketing initiatives:", 24);
    document.Call("Push", paragraph);
    // list
    createNumbering(api, data["plans"]["marketing_initiatives"], "bullet", 22);

    // save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
