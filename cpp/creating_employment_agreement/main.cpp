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
const wchar_t* resultPath = L"result.pdf";

const int defaultFontSize = 24;
const string defaultJc = "both";
const string signerData[] = {
    "Name: __________________________",
    "Signature: _______________________",
    "Date: ___________________________"
};

// Helper functions
CValue createParagraph(CValue api, string text, bool isBold = false, int fontSize = defaultFontSize, string jc = defaultJc)
{
    CValue paragraph = api.Call("CreateParagraph");
    paragraph.Call("AddText", text.c_str());
    paragraph.Call("SetBold", isBold);
    if (fontSize != defaultFontSize) {
        paragraph.Call("SetFontSize", fontSize);
    }
    if (jc != defaultJc ) {
        paragraph.Call("SetJc", jc.c_str());
    }
    return paragraph;
}

CValue createRun(CValue api, string text, bool isBold = false, int fontSize = defaultFontSize) {
    CValue run = api.Call("CreateRun");
    run.Call("AddText", text.c_str());
    run.Call("SetBold", isBold);
    if (fontSize != defaultFontSize) {
        run.Call("SetFontSize", fontSize);
    }
    return run;
}

void setNumbering(CValue paragraph, CValue numLvl) {
    paragraph.Call("SetNumbering", numLvl);
}

void setSpacingAfter(CValue paragraph, int spacing) {
    paragraph.Call("SetSpacingAfter", spacing);
}

CValue createConditionsDescParagraph(CValue api, string text) {
    // create paragraph with first line indentation
    CValue paragraph = createParagraph(api, text);
    paragraph.Call("SetIndFirstLine", 400);
    return paragraph;
}

void addParticipantToParagraph(CValue api, CValue paragraph, string pType, string details) {
    paragraph.Call("Push", createRun(api, pType + ": ", true));
    paragraph.Call("Push", createRun(api, details));
}

CValue createNumberedSection(CValue api, string text, CValue numLvl) {
    CValue paragraph = createParagraph(api, text, true);
    setNumbering(paragraph, numLvl);
    setSpacingAfter(paragraph, 50);
    return paragraph;
}

CValue createWorkCondition(CValue api, string title, string text, CValue numLvl, bool setSpacing = false) {
    CValue paragraph = api.Call("CreateParagraph");
    setNumbering(paragraph, numLvl);
    if (setSpacing) {
        setSpacingAfter(paragraph, 20);
    }
    paragraph.Call("SetJc", "left");
    paragraph.Call("Push", createRun(api, title + ": ", true));
    paragraph.Call("Push", createRun(api, text));
    return paragraph;
}

void fillSigner(CValue api, CValue cell, string title) {
    CValue paragraph = cell.Call("GetContent").Call("GetElement", 0);
    paragraph.Call("SetJc", "left");
    paragraph.Call("Push", createRun(api, title, true));

    for (const auto& text : signerData) {
        paragraph.Call("AddLineBreak");
        paragraph.Call("Push", createRun(api, text));
    }
}

// Main function
int main()
{
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/employment_agreement_response.json";
    ifstream fs(jsonPath);
    json data = json::parse(fs);

    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;
    builder.CreateFile(OFFICESTUDIO_FILE_DOCUMENT_OFORM_PDF);

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];
    CValue document = api.Call("GetDocument");

    // DOCUMENT STYLE
    CValue paraPr = document.Call("GetDefaultParaPr");
    paraPr.Call("SetJc", "both");
    CValue textPr = document.Call("GetDefaultTextPr");
    textPr.Call("SetFontSize", 24);
    textPr.Call("SetFontFamily", "Times New Roman");

    // DOCUMENT HEADER
    CValue header = document.Call("GetElement", 0);
    header.Call("AddText", "EMPLOYMENT AGREEMENT");
    header.Call("SetFontSize", 28);
    header.Call("SetBold", true);

    CValue headerDesc = createParagraph(
        api,
        "This Employment Agreement (\"Agreement\") is made and entered into on " + data["date"].get<string>() + " by and between:"
    );
    setSpacingAfter(headerDesc, 50);
    document.Call("Push", headerDesc);

    // PARTICIPANTS OF THE DOCUMENT
    CValue participants = createParagraph(api, "", false, defaultFontSize, "left");
    const json& employer = data["employer"];
    addParticipantToParagraph(
        api,
        participants,
        "Employer",
        employer["name"].get<string>() + ", located at " + employer["address"].get<string>() + "."
    );
    participants.Call("AddLineBreak");
    const json& employee = data["employee"];
    addParticipantToParagraph(
        api,
        participants,
        "Employee",
        employee["name"].get<string>() + ", residing at " + employee["address"].get<string>() + "."
    );
    document.Call("Push", participants);
    document.Call("Push", createParagraph(api, "The parties agree to the following terms and conditions:"));

    // AGREEMENT CONDITIONS
    // Create numbering
//    CValue numbering = document.Call("CreateNumbering", "numbered");
//    CValue numberingLvl = numbering.Call("GetLevel", 0);
//    numberingLvl.Call("SetCustomType", "decimal", "%1.", "left");
//    numberingLvl.Call("SetSuff", "space");

    // Position and duties
//    document.Call("Push", createNumberedSection(api, "POSITION AND DUTIES", numberingLvl));
//    document.Call(
//        "Push",
//        createConditionsDescParagraph(
//            api,
//            "The Employee is hired as " + data["position_and_duties"]["job_title"].get<string>() + ". " +
//            "The Employee shall perform their duties as outlined by the Employer and comply with all applicable policies and guidelines."
//        )
//    );

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_OFORM_PDF, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
