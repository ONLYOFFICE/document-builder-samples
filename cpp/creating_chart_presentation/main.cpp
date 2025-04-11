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

#include <string>
#include <vector>
#include "common.h"
#include "docbuilder.h"

#include "out/cpp/builder_path.h"
#include "resources/utils/utils.h"

using namespace std;
using namespace NSDoctRenderer;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.pptx";

void addText(CValue api, int fontSize, string text, CValue slide, CValue shape, CValue paragraph, CValue fill, string jc)
{
    CValue run = api.Call("CreateRun");
    CValue textPr = run.Call("GetTextPr");
    textPr.Call("SetFontSize", fontSize);
    textPr.Call("SetFill", fill);
    textPr.Call("SetFontFamily", "Tahoma");
    paragraph.Call("SetJc", jc.c_str());
    run.Call("AddText", text.c_str());
    run.Call("AddLineBreak");
    paragraph.Call("AddElement", run);
    slide.Call("AddObject", shape);
}

// Main function
int main()
{
    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;

    // Read chart data from xlsx
    wstring templatePath = NSUtils::GetResourcesDirectory() + L"/docs/chart_data.xlsx";
    builder.OpenFile(templatePath.c_str(), L"");
    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];
    CValue worksheet = api.Call("GetActiveSheet");
    CValue values = worksheet.Call("GetUsedRange").Call("GetValue");

    int sizeX = values.GetLength();
    int sizeY = values[0].GetLength();
    vector<vector<wstring>> data(sizeX, vector<wstring>(sizeY));
    for (int i = 0; i < sizeX; i++)
    {
        for (int j = 0; j < sizeY; j++)
        {
            data[i][j] = values[i][j].ToString().c_str();
        }
    }
    builder.CloseFile();

    // Create chart presentation
    builder.CreateFile(OFFICESTUDIO_FILE_PRESENTATION_PPTX);
    context = builder.GetContext();
    global = context.GetGlobal();
    api = global["Api"];
    CValue presentation = api.Call("GetPresentation");
    CValue slide = presentation.Call("GetSlideByIndex", 0);
    slide.Call("RemoveAllObjects");

    CValue rGBColor = api.Call("CreateRGBColor", 255, 244, 240);
    CValue fill = api.Call("CreateSolidFill", rGBColor);
    slide.Call("SetBackground", fill);

    CValue stroke = api.Call("CreateStroke", 0, api.Call("CreateNoFill"));
    CValue shapeTitle = api.Call("CreateShape", "rect", 300 * 36000, 20 * 36000, api.Call("CreateNoFill"), stroke);
    CValue shapeText = api.Call("CreateShape", "rect", 120 * 36000, 80 * 36000, api.Call("CreateNoFill"), stroke);
    shapeTitle.Call("SetPosition", 20 * 36000, 20 * 36000);
    shapeText.Call("SetPosition", 210 * 36000, 50 * 36000);
    CValue paragraphTitle = shapeTitle.Call("GetDocContent").Call("GetElement", 0);
    CValue paragraphText = shapeText.Call("GetDocContent").Call("GetElement", 0);
    rGBColor = api.Call("CreateRGBColor", 115, 81, 68);
    fill = api.Call("CreateSolidFill", rGBColor);

    string titleContent = "Price Type Report";
    string textContent = "This is an overview of price types. As we can see, May was the price peak, but even in June the price went down, the annual upward trend persists.";
    addText(api, 80, titleContent, slide, shapeTitle, paragraphTitle, fill, "center");
    addText(api, 42, textContent, slide, shapeText, paragraphText, fill, "left");

    // Transform 2d array into cols names, rows names and data
    CValue cols = context.CreateArray(sizeY - 1);
    for (int col = 1; col < sizeY; col++)
    {
        cols[col - 1] = data[0][col].c_str();
    }

    CValue rows = context.CreateArray(sizeX - 1);
    for (int row = 1; row < sizeX; row++)
    {
        rows[row - 1] = data[row][0].c_str();
    }

    CValue vals = context.CreateArray(sizeY - 1);
    for (int row = 1; row < sizeY; row++)
    {
        CValue row_data = context.CreateArray(sizeX - 1);
        for (int col = 1; col < sizeX; col++)
        {
            row_data[col - 1] = data[col][row].c_str();
        }
        vals[row - 1] = row_data;
    }

    // Pass CValue data to the CreateChart method
    CValue chart = api.Call("CreateChart", "lineStacked", vals, cols, rows);
    chart.Call("SetSize", 180 * 36000, 100 * 36000);
    chart.Call("SetPosition", 20 * 36000, 50 * 36000);
    chart.Call("ApplyChartStyle", 24);
    chart.Call("SetLegendFontSize", 12);
    chart.Call("SetLegendPos", "top");
    slide.Call("AddObject", chart);

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_PRESENTATION_PPTX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
