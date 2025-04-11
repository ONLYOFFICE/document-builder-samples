/**
 *
 * (c) Copyright Ascensio System SIA 2025
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
    for (int i = 0; i < (int)data.size(); i++)
    {
        CValue paragraph = getTableCellParagraph(table, 0, i);
        addTextToParagraph(paragraph, data[i], fontSize, true);
    }
}

void fillTableBody(CValue table, const json& data, const vector<string>& keys, int fontSize, int startRow = 1)
{
    for (int row = 0; row < (int)data.size(); row++)
    {
        for (int col = 0; col < (int)keys.size(); col++)
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

// Main function
int main()
{
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/hrms_response.json";
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

    // TITLE PAGE
    // header
    CValue paragraph = document.Call("GetElement", 0);
    addTextToParagraph(paragraph, "Employee Development Plan for 2024", 48, true, "center");
    paragraph.Call("SetSpacingBefore", 5000);
    paragraph.Call("SetSpacingAfter", 500);
    // employee name
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, data["employee"]["name"].get<string>(), 36, false, "center");
    document.Call("Push", paragraph);
    // employee position and department
    paragraph = api.Call("CreateParagraph");
    string employeeInfo = "Position: " + data["employee"]["position"].get<string>();
    employeeInfo += "\nDepartment: " + data["employee"]["department"].get<string>();
    addTextToParagraph(paragraph, employeeInfo, 24, false, "center");
    paragraph.Call("AddPageBreak");
    document.Call("Push", paragraph);

    // COMPETENCIES SECION
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Competencies", 32, true);
    document.Call("Push", paragraph);
    // technical skills sub-header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Technical skills:", 24);
    document.Call("Push", paragraph);
    // technical skills table
    const json& technicalSkills = data["competencies"]["technical_skills"];
    CValue table = createTable(api, (int)technicalSkills.size() + 1, 2);
    fillTableHeaders(table, { "Skill", "Level" }, 22);
    fillTableBody(table, technicalSkills, { "name", "level" }, 22);
    document.Call("Push", table);
    // soft skills sub-header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Soft skills:", 24);
    document.Call("Push", paragraph);
    // soft skills table
    const json& softSkills = data["competencies"]["soft_skills"];
    table = createTable(api, (int)softSkills.size() + 1, 2);
    fillTableHeaders(table, { "Skill", "Level" }, 22);
    fillTableBody(table, softSkills, { "name", "level" }, 22);
    document.Call("Push", table);

    // DEVELOPMENT AREAS section
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Development areas", 32, true);
    document.Call("Push", paragraph);
    // list
    createNumbering(api, data["development_areas"], "numbered", 22);

    // GOALS section
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Goals for next year", 32, true);
    document.Call("Push", paragraph);
    // numbering
    paragraph = createNumbering(api, data["goals_next_year"], "numbered", 22);
    // add a page break after the last paragraph
    paragraph.Call("AddPageBreak");

    // RESOURCES section
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Recommended resources", 32, true);
    document.Call("Push", paragraph);
    // table
    const json& resources = data["resources"];
    table = createTable(api, (int)resources.size() + 1, 3);
    fillTableHeaders(table, { "Name", "Provider", "Duration" }, 22);
    fillTableBody(table, resources, { "name", "provider", "duration" }, 22);
    document.Call("Push", table);

    // FEEDBACK section
    // header
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Feedback", 32, true);
    document.Call("Push", paragraph);
    // manager's feedback
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Manager's feedback:", 24, false);
    document.Call("Push", paragraph);
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, string(280, '_'), 24, false);
    document.Call("Push", paragraph);
    // employees's feedback
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Employee's feedback:", 24, false);
    document.Call("Push", paragraph);
    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, string(280, '_'), 24, false);
    document.Call("Push", paragraph);

    // save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
