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

using namespace NSDoctRenderer;

const wchar_t* workDir = BUILDER_DIR;
const wchar_t* resultPath = L"result.docx";

// Helper functions
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
void setTextFormProperties(CValue textForm, std::string key, std::string tip, bool required, std::string placeholder, bool comb, int maxCharacters, int cellWidth, bool multiLine, bool autoFit)
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

    // Create basic form
    CValue document = api.Call("GetDocument");
    CValue paragraph = document.Call("GetElement", 0);
    CValue headingStyle = document.Call("GetStyle", "Heading 3");

    paragraph.Call("AddText", "Employee pass card");
    paragraph.Call("SetStyle", headingStyle);
    document.Call("Push", paragraph);

    CValue pictureForm = api.Call("CreatePictureForm");
    setPictureFormProperties(pictureForm, "Photo", "Upload your photo", false, "Photo", "tooBig", true, false, 50, 50);
    paragraph = api.Call("CreateParagraph");
    paragraph.Call("AddElement", pictureForm);
    document.Call("Push", paragraph);

    CValue textForm = api.Call("CreateTextForm");
    setTextFormProperties(textForm, "First name", "Enter your first name", false, "First name", true, 13, 3, false, false);
    paragraph = api.Call("CreateParagraph");
    paragraph.Call("AddElement", textForm);
    document.Call("Push", paragraph);

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
