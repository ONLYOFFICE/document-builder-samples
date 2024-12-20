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

// Main function
int main()
{
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/hrms_response.json";
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

    // TITLE PAGE
    // header
    CValue oParagraph = oDocument.Call("GetElement", 0);
    addTextToParagraph(oParagraph, "Employee Development Plan for 2024", 48, true, "center");
    oParagraph.Call("SetSpacingBefore", 5000);
    oParagraph.Call("SetSpacingAfter", 500);
    // employee name
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, data["employee"]["name"].get<string>(), 36, false, "center");
    oDocument.Call("Push", oParagraph);
    // employee position and department
    oParagraph = oApi.Call("CreateParagraph");
    string employeeInfo = "Position: " + data["employee"]["position"].get<string>();
    employeeInfo += "\nDepartment: " + data["employee"]["department"].get<string>();
    addTextToParagraph(oParagraph, employeeInfo, 24, false, "center");
    oParagraph.Call("AddPageBreak");
    oDocument.Call("Push", oParagraph);

    // COMPETENCIES SECION
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Competencies", 32, true);
    oDocument.Call("Push", oParagraph);
    // technical skills sub-header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Technical skills:", 24);
    oDocument.Call("Push", oParagraph);
    // technical skills table
    const json& technicalSkills = data["competencies"]["technical_skills"];
    CValue oTable = createTable(oApi, (int)technicalSkills.size() + 1, 2);
    fillTableHeaders(oTable, { "Skill", "Level" }, 22);
    fillTableBody(oTable, technicalSkills, { "name", "level" }, 22);
    oDocument.Call("Push", oTable);
    // soft skills sub-header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Soft skills:", 24);
    oDocument.Call("Push", oParagraph);
    // soft skills table
    const json& softSkills = data["competencies"]["soft_skills"];
    oTable = createTable(oApi, (int)softSkills.size() + 1, 2);
    fillTableHeaders(oTable, { "Skill", "Level" }, 22);
    fillTableBody(oTable, softSkills, { "name", "level" }, 22);
    oDocument.Call("Push", oTable);

    // DEVELOPMENT AREAS section
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Development areas", 32, true);
    oDocument.Call("Push", oParagraph);
    // list
    createNumbering(oApi, data["development_areas"], "numbered", 22);

    // GOALS section
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Goals for next year", 32, true);
    oDocument.Call("Push", oParagraph);
    // numbering
    oParagraph = createNumbering(oApi, data["goals_next_year"], "numbered", 22);
    // add a page break after the last paragraph
    oParagraph.Call("AddPageBreak");

    // RESOURCES section
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Recommended resources", 32, true);
    oDocument.Call("Push", oParagraph);
    // table
    const json& resources = data["resources"];
    oTable = createTable(oApi, (int)resources.size() + 1, 3);
    fillTableHeaders(oTable, { "Name", "Provider", "Duration" }, 22);
    fillTableBody(oTable, resources, { "name", "provider", "duration" }, 22);
    oDocument.Call("Push", oTable);

    // FEEDBACK section
    // header
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Feedback", 32, true);
    oDocument.Call("Push", oParagraph);
    // manager's feedback
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Manager's feedback:", 24, false);
    oDocument.Call("Push", oParagraph);
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, string(280, '_'), 24, false);
    oDocument.Call("Push", oParagraph);
    // employees's feedback
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, "Employee's feedback:", 24, false);
    oDocument.Call("Push", oParagraph);
    oParagraph = oApi.Call("CreateParagraph");
    addTextToParagraph(oParagraph, string(280, '_'), 24, false);
    oDocument.Call("Push", oParagraph);

    // save and close
    oBuilder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    oBuilder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
