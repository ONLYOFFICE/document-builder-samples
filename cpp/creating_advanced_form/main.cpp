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
const wchar_t* resultPath = L"result.docx";

// Helper functions
void setTableBorders(CValue table, int borderColor)
{
    table.Call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
    table.Call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
}

CValue createFullWidthTable(CValue api, int rows, int cols, int borderColor)
{
    CValue table = api.Call("CreateTable", cols, rows);
    table.Call("SetWidth", "percent", 100);
    setTableBorders(table, borderColor);
    return table;
}

CValue getTableCellParagraph(CValue table, int row, int col)
{
    return table.Call("GetCell", row, col).Call("GetContent").Call("GetElement", 0);
}

void addTextToParagraph(CValue paragraph, std::string text, int fontSize, bool isBold)
{
    paragraph.Call("AddText", text.c_str());
    paragraph.Call("SetFontSize", fontSize);
    paragraph.Call("SetBold", isBold);
}

void setPictureFormProperties(CValue pictureForm, std::string key, std::string tip, bool required, std::string placeholder, std::string scaleFlag, bool lockAspectRatio, bool respectBorders, int shiftX, int shiftY)
{
    pictureForm.Call("SetFormKey", key.c_str());
    pictureForm.Call("SetTipText", tip.c_str());
    pictureForm.Call("SetRequired", required);
    pictureForm.Call("SetPlaceholderText", placeholder.c_str());
    pictureForm.Call("SetScaleFlag", scaleFlag.c_str());
    pictureForm.Call("SetLockAspectRatio", lockAspectRatio);
    pictureForm.Call("SetRespectBorders", respectBorders);
    pictureForm.Call("SetPicturePosition", shiftX, shiftY);
}

void setTextFormProperties(CValue textForm, string key, string tip, bool required, string placeholder, bool comb, int maxCharacters, int cellWidth, bool multiLine, bool autoFit)
{
    textForm.Call("SetFormKey", key.c_str());
    textForm.Call("SetTipText", tip.c_str());
    textForm.Call("SetRequired", required);
    textForm.Call("SetPlaceholderText", placeholder.c_str());
    textForm.Call("SetComb", comb);
    textForm.Call("SetCharactersLimit", maxCharacters);
    textForm.Call("SetCellWidth", cellWidth);
    textForm.Call("SetCellWidth", multiLine);
    textForm.Call("SetMultiline", autoFit);
}

void addTextFormToParagraph(CValue paragraph, CValue textForm, int fontSize, string jc, bool hasBorder, int borderColor)
{
    if (hasBorder)
    {
        textForm.Call("SetBorderColor", borderColor, borderColor, borderColor);
    }
    paragraph.Call("AddElement", textForm);
    paragraph.Call("SetFontSize", fontSize);
    paragraph.Call("SetJc", jc.c_str());
}

// Main function
int main()
{
    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;
    builder.CreateFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX);

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];

    // Create advanced form
    CValue document = api.Call("GetDocument");
    CValue table = createFullWidthTable(api, 1, 2, 255);
    CValue paragraph = getTableCellParagraph(table, 0, 0);
    addTextToParagraph(paragraph, "PURCHASE ORDER", 36, true);
    paragraph = getTableCellParagraph(table, 0, 1);
    addTextToParagraph(paragraph, "Serial # ", 25, true);

    CValue textForm = api.Call("CreateTextForm");
    setTextFormProperties(textForm, "Serial", "Enter serial number", false, "Serial", true, 6, 1, false, false);
    addTextFormToParagraph(paragraph, textForm, 25, "left", true, 255);
    document.Call("Push", table);

    CValue pictureForm = api.Call("CreatePictureForm");
    setPictureFormProperties(pictureForm, "Photo", "Upload company logo", false, "Photo", "tooBig", false, false, 0, 0);
    paragraph = api.Call("CreateParagraph");
    paragraph.Call("AddElement", pictureForm);
    document.Call("Push", paragraph);

    textForm = api.Call("CreateTextForm");
    setTextFormProperties(textForm, "Company Name", "Enter company name", false, "Company Name", true, 20, 1, false, false);
    paragraph = api.Call("CreateParagraph");
    addTextFormToParagraph(paragraph, textForm, 35, "left", false, 255);
    document.Call("Push", paragraph);

    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "Date: ", 25, true);
    textForm = api.Call("CreateTextForm");
    setTextFormProperties(textForm, "Date", "Date", false, "DD.MM.YYYY", true, 10, 1, false, false);
    addTextFormToParagraph(paragraph, textForm, 25, "left", true, 255);
    document.Call("Push", paragraph);

    paragraph = api.Call("CreateParagraph");
    addTextToParagraph(paragraph, "To:", 35, true);
    document.Call("Push", paragraph);

    table = createFullWidthTable(api, 1, 1, 200);
    paragraph = getTableCellParagraph(table, 0, 0);
    textForm = api.Call("CreateTextForm");
    setTextFormProperties(textForm, "Recipient", "Recipient", false, "Recipient", true, 25, 1, false, false);
    addTextFormToParagraph(paragraph, textForm, 32, "left", false, 255);
    document.Call("Push", table);

    table = createFullWidthTable(api, 10, 2, 200);
    table.Call("GetRow", 0).Call("SetBackgroundColor", 245, 245, 245, false);
    CValue cell = table.Call("GetCell", 0, 0);
    cell.Call("SetWidth", "percent", 30);
    paragraph = getTableCellParagraph(table, 0, 0);
    addTextToParagraph(paragraph, "Qty.", 30, true);
    paragraph = getTableCellParagraph(table, 0, 1);
    addTextToParagraph(paragraph, "Description", 30, true);

    for (int i = 1; i < 10; i++)
    {
        CValue tempParagraph = getTableCellParagraph(table, i, 0);
        CValue tempTextForm = api.Call("CreateTextForm");
        setTextFormProperties(tempTextForm, "Qty" + std::to_string(i), "Qty" + std::to_string(i), false, " ", true, 9, 1, false, false);
        addTextFormToParagraph(tempParagraph, tempTextForm, 30, "left", false, 255);
        tempParagraph = getTableCellParagraph(table, i, 1);
        tempTextForm = api.Call("CreateTextForm");
        setTextFormProperties(tempTextForm, "Description" + std::to_string(i), "Description" + std::to_string(i), false, " ", true, 22, 1, false, false);
        addTextFormToParagraph(tempParagraph, tempTextForm, 30, "left", false, 255);
    }

    document.Call("Push", table);
    document.Call("RemoveElement", 0);
    document.Call("RemoveElement", 1);

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
