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

#include <map>
#include <string>
#include "common.h"
#include "docbuilder.h"

#include "out/cpp/builder_path.h"
#include "resources/utils/utils.h"

using namespace std;
using namespace NSDoctRenderer;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.docx";

// Main function
int main()
{
    std::map<wstring, wstring> formData;
    formData[L"Photo"] = L"https://static.onlyoffice.com/assets/docs/samples/img/onlyoffice_logo.png";
    formData[L"Serial"] = L"A1345";
    formData[L"Company Name"] = L"ONLYOFFICE";
    formData[L"Date"] = L"25.12.2023";
    formData[L"Recipient"] = L"Space Corporation";
    formData[L"Qty1"] = L"25";
    formData[L"Description1"] = L"Frame";
    formData[L"Qty2"] = L"2";
    formData[L"Description2"] = L"Stack";
    formData[L"Qty3"] = L"34";
    formData[L"Description3"] = L"Shifter";

    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;
    wstring templatePath = NSUtils::GetResourcesDirectory() + L"/docs/form.docx";
    builder.OpenFile(templatePath.c_str(), L"");

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];

    // Fill form
    CValue document = api.Call("GetDocument");
    CValue aForms = document.Call("GetAllForms");

    int formNum = 0;
    while (formNum < (int)aForms.GetLength())
    {
        CValue form = aForms[formNum];
        wstring type = aForms[formNum].Call("GetFormType").ToString().c_str();
        wstring value = formData[aForms[formNum].Call("GetFormKey").ToString().c_str()];
        if (type == L"textForm") form.Call("SetText", value.c_str());
        if (type == L"pictureForm") form.Call("SetImage", value.c_str());
        formNum++;
    }

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
