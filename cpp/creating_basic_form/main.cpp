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
void setPictureFormProperties(CValue oPictureForm, std::string key, std::string tip, bool required, std::string placeholder, std::string scaleFlag, bool lockAspectRatio, bool respectBorders, int shiftX, int shiftY)
{
    oPictureForm.Call("SetFormKey", key.c_str());
    oPictureForm.Call("SetTipText", tip.c_str());
    oPictureForm.Call("SetRequired", required);
    oPictureForm.Call("SetPlaceholderText", placeholder.c_str());
    oPictureForm.Call("SetScaleFlag", scaleFlag.c_str());
    oPictureForm.Call("SetLockAspectRatio", lockAspectRatio);
    oPictureForm.Call("SetRespectBorders", respectBorders);
    oPictureForm.Call("SetPicturePosition", shiftX, shiftY);
}
void setTextFormProperties(CValue oTextForm, std::string key, std::string tip, bool required, std::string placeholder, bool comb, int maxCharacters, int cellWidth, bool multiLine, bool autoFit)
{
    oTextForm.Call("SetFormKey", key.c_str());
    oTextForm.Call("SetTipText", tip.c_str());
    oTextForm.Call("SetRequired", required);
    oTextForm.Call("SetPlaceholderText", placeholder.c_str());
    oTextForm.Call("SetComb", comb);
    oTextForm.Call("SetCharactersLimit", maxCharacters);
    oTextForm.Call("SetCellWidth", cellWidth);
    oTextForm.Call("SetCellWidth", multiLine);
    oTextForm.Call("SetMultiline", autoFit);
}

// Main function
int main()
{
    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder oBuilder;
    oBuilder.SetProperty("--work-directory", workDir);
    oBuilder.CreateFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX);

    CContext oContext = oBuilder.GetContext();
    CContextScope oScope = oContext.CreateScope();
    CValue oGlobal = oContext.GetGlobal();
    CValue oApi = oGlobal["Api"];

    // Create basic form
    CValue oDocument = oApi.Call("GetDocument");
    CValue oParagraph = oDocument.Call("GetElement", 0);
    CValue oHeadingStyle = oDocument.Call("GetStyle", "Heading 3");

    oParagraph.Call("AddText", "Employee pass card");
    oParagraph.Call("SetStyle", oHeadingStyle);
    oDocument.Call("Push", oParagraph);

    CValue oPictureForm = oApi.Call("CreatePictureForm");
    setPictureFormProperties(oPictureForm, "Photo", "Upload your photo", false, "Photo", "tooBig", true, false, 50, 50);
    oParagraph = oApi.Call("CreateParagraph");
    oParagraph.Call("AddElement", oPictureForm);
    oDocument.Call("Push", oParagraph);

    CValue oTextForm = oApi.Call("CreateTextForm");
    setTextFormProperties(oTextForm, "First name", "Enter your first name", false, "First name", true, 13, 3, false, false);
    oParagraph = oApi.Call("CreateParagraph");
    oParagraph.Call("AddElement", oTextForm);
    oDocument.Call("Push", oParagraph);

    // Save and close
    oBuilder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    oBuilder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
