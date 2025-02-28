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
#include <locale>
#include <sstream>

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

void setSpacingAfter(CValue paragraph, int spacing) {
    paragraph.Call("SetSpacingAfter", spacing);
}

void fillHeader(CValue paragraph, string text) {
    paragraph.Call("AddText", text.c_str());
    paragraph.Call("SetFontSize", 28);
    paragraph.Call("SetBold", true);
    setSpacingAfter(paragraph, 50);
}

void setNumbering(CValue paragraph, CValue numLvl) {
    paragraph.Call("SetNumbering", numLvl);
}

CValue createDetailsHeader(CValue api, string text) {
    CValue paragraph = api.Call("CreateParagraph");
    paragraph.Call("AddText", text.c_str());
    paragraph.Call("SetBold", true);
    paragraph.Call("SetItalic", true);
    setSpacingAfter(paragraph, 40);
    return paragraph;
}

void setupRequisitesStyle(CValue paragraph, CValue numLvl, bool setSpacing) {
    if (setSpacing) {
        setSpacingAfter(paragraph, 20);
    }
    if (!numLvl.IsUndefined()) {
        setNumbering(paragraph, numLvl);
    }
}

string formatSum(int value) {
    std::ostringstream oss;
    oss.imbue(std::locale("en_US.UTF-8"));
    oss << "$" << value;
    return oss.str();
}

CValue createRequisitesParagraph(CValue api, string title, string details, CValue numLvl = CValue::CreateUndefined(), bool setSpacing = true, bool setTitleBold = true) {
    CValue paragraph = api.Call("CreateParagraph");
    CValue titleRun = paragraph.Call("AddText", (title + ": ").c_str());
    if (setTitleBold) {
        titleRun.Call("SetBold", true);
    } else {
        titleRun.Call("SetItalic", true);
    }
    CValue detailsRun = paragraph.Call("AddText", details.c_str());
    detailsRun.Call("SetItalic", true);
    setupRequisitesStyle(paragraph, numLvl, setSpacing);
    return paragraph;
}

void setupTableStyle(CValue document, CValue table) {
    // table size
    table.Call("SetWidth", "percent", 100);
    table.Call("Select");
    CValue tableRange = document.Call("GetRangeBySelect");
    CValue tableParagraphs = tableRange.Call("GetAllParagraphs");
    for (int i = 0; i < (int)tableParagraphs.GetLength(); i++) {
        CValue paraPr = tableParagraphs.Get(i).Call("GetParaPr");
        paraPr.Call("SetSpacingBefore", 40);
        paraPr.Call("SetSpacingAfter", 40);
    }

    // table borders
    table.Call("SetTableBorderTop", "single", 4, 0, 0, 0, 0);
    table.Call("SetTableBorderBottom", "single", 4, 0, 0, 0, 0);
    table.Call("SetTableBorderLeft", "single", 4, 0, 0, 0, 0);
    table.Call("SetTableBorderRight", "single", 4, 0, 0, 0, 0);
    table.Call("SetTableBorderInsideV", "single", 4, 0, 0, 0, 0);
    table.Call("SetTableBorderInsideH", "single", 4, 0, 0, 0, 0);
}

CValue getCellContent(CValue cell) {
    return cell.Call("GetContent").Call("GetElement", 0);
}

void fillTableContent(CValue table, json& items) {
    string tableHeaders[] = {"Description", "Quantity", "Unit Price", "Total"};
    string tableFields[] = {"description", "quantity", "unit_price", "total"};

    // fill table header
    const int tableHeadersSize = sizeof(tableHeaders) / sizeof(tableHeaders[0]);
    CValue headerRow = table.Call("GetRow", 0);
    for (int i = 0; i < tableHeadersSize; i++) {
        CValue headerCell = getCellContent(headerRow.Call("GetCell", i));
        headerCell.Call("AddText", tableHeaders[i].c_str());
        headerCell.Call("SetBold", true);
        headerCell.Call("SetJc", "center");
    }

    // fill items
    const int tableFieldsSize = sizeof(tableFields) / sizeof(tableFields[0]);
    for (int i = 0; i < (int)items.size(); i++) {
        CValue row = table.Call("GetRow", i + 1);
        for (int j = 0; j < tableFieldsSize; j++) {
            CValue cell = getCellContent(row.Call("GetCell", j));
            string key = tableFields[j];

            // Handle different field types
            if (key == "unit_price" || key == "total") {
                int value = items[i][key].get<int>();
                cell.Call("AddText", formatSum(value).c_str());
            } else {
                json value = items[i][key];
                string strValue;
                if (value.is_string()) {
                    strValue = value.get<string>();
                } else {
                    strValue = to_string(value.get<int>());
                }
                cell.Call("AddText", strValue.c_str());
            }
        }
    }
}

int main() {
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/commercial_offer_data.json";
    ifstream fs(jsonPath);
    json data = json::parse(fs);

    // Init DocBuilder
    CDocBuilder::Initialize(workDir);
    CDocBuilder builder;
    builder.CreateFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX);

    CContext context = builder.GetContext();
    CValue global = context.GetGlobal();
    CValue api = global["Api"];
    CValue document = api.Call("GetDocument");

    // page margins
    CValue section = document.Call("GetFinalSection");
    section.Call("SetPageMargins", 1440, 1280, 1440, 1280);

    // DOCUMENT STYLE
    CValue paraPr = document.Call("GetDefaultParaPr");
    paraPr.Call("SetSpacingAfter", 100);
    CValue textPr = document.Call("GetDefaultTextPr");
    textPr.Call("SetFontSize", 24);
    textPr.Call("SetFontFamily", "Times New Roman");

    // DOCUMENT HEADER
    CValue header = document.Call("GetElement", 0);
    fillHeader(header, "COMMERCIAL OFFER TEMPLATE");

    // document requisites
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Offer No.", data["offer"]["number"].get<string>())
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Date", data["offer"]["date"].get<string>(), CValue::CreateUndefined(), false)
    );

    // bullet numbering
    CValue bulletNumbering = document.Call("CreateNumbering", "bullet");
    CValue bNumLvl = bulletNumbering.Call("GetLevel", 0);

    // SELLER INFORMATION
    CValue sellerHeader = createDetailsHeader(api, "SELLER INFORMATION");
    document.Call("Push", sellerHeader);

    // seller details
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Company Name", data["seller"]["company_name"].get<string>(), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Address", data["seller"]["address"].get<string>(), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Tax ID (TIN)", data["seller"]["tin"].get<string>(), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Contact Information", "", bNumLvl)
    );

    // contact details
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Phone", data["seller"]["contact"]["phone"].get<string>(), bNumLvl, true, false)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Email", data["seller"]["contact"]["email"].get<string>(), bNumLvl, false, false)
    );

    // BUYER INFORMATION
    CValue buyerHeader = createDetailsHeader(api, "BUYER INFORMATION");
    document.Call("Push", buyerHeader);

    // buyer details
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Company Name", data["buyer"]["company_name"].get<string>(), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Address", data["buyer"]["address"].get<string>(), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Contact Person", data["buyer"]["contact_person"].get<string>(), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Email", data["buyer"]["email"].get<string>(), bNumLvl, false)
    );

    // OFFER DETAILS
    CValue tableHeader = api.Call("CreateParagraph");
    fillHeader(tableHeader, "OFFER DETAILS");
    document.Call("Push", tableHeader);

    // table content
    json offerDetails = data["offer_details"];
    CValue itemsTable = api.Call("CreateTable", 4, (int)offerDetails.size() + 1);
    document.Call("Push", itemsTable);
    setupTableStyle(document, itemsTable);
    fillTableContent(itemsTable, offerDetails);

    // TOTALS
    CValue totals = createDetailsHeader(api, "TOTALS");
    document.Call("Push", totals);
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Subtotal", formatSum(data["totals"]["subtotal"].get<int>()), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Discount", formatSum(data["totals"]["discount"].get<int>()), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Tax (e.g., 20% VAT)", formatSum(data["totals"]["tax"].get<int>()), bNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Total Amount", formatSum(data["totals"]["total"].get<int>()), bNumLvl, false)
    );

    // TERMS AND CONDITIONS
    CValue sellerHeader2 = createDetailsHeader(api, "TERMS AND CONDITIONS");
    document.Call("Push", sellerHeader2);

    // numbering
    CValue numbering = document.Call("CreateNumbering", "numbered");
    CValue dNumLvl = numbering.Call("GetLevel", 0);
    dNumLvl.Call("SetCustomType", "decimal", "%1.", "left");

    document.Call(
        "Push",
        createRequisitesParagraph(api, "Validity Period", data["terms_and_conditions"]["validity_period"].get<string>(), dNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Payment Terms", data["terms_and_conditions"]["payment_terms"].get<string>(), dNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Delivery Terms", data["terms_and_conditions"]["delivery_terms"].get<string>(), dNumLvl)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Additional Notes", data["terms_and_conditions"]["additional_notes"].get<string>(), dNumLvl, false)
    );

    // SIGNATURE
    CValue signHeader = api.Call("CreateParagraph");
    signHeader.Call("AddText", "Signature:");
    signHeader.Call("SetBold", true);
    document.Call("Push", signHeader);

    CValue signDetails = api.Call("CreateParagraph");
    signDetails.Call(
        "AddText",
        (data["seller"]["authorized_person"]["full_name"].get<string>() + ", " +
         data["seller"]["authorized_person"]["position"].get<string>()).c_str()
    );
    signDetails.Call("AddLineBreak");
    signDetails.Call("AddText", data["seller"]["company_name"].get<string>().c_str());
    document.Call("Push", signDetails);

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_DOCX, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
