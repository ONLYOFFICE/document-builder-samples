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
#include <algorithm>

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

// Helper functions
void fillHeader(CValue paragraph, string text) {
    paragraph.Call("AddText", text.c_str());
    paragraph.Call("SetFontSize", 28);
    paragraph.Call("SetBold", true);
}

void setSpacingAfter(CValue paragraph, int spacing) {
    paragraph.Call("SetSpacingAfter", spacing);
}

void setNumbering(CValue paragraph, CValue numLvl) {
    paragraph.Call("SetNumbering", numLvl);
}

CValue createDetailsHeader(CValue api, string text) {
    CValue paragraph = api.Call("CreateParagraph");
    paragraph.Call("AddText", text.c_str());
    paragraph.Call("SetBold", true);
    paragraph.Call("SetItalic", true);
    setSpacingAfter(paragraph, 50);
    return paragraph;
}

void setupRequisitesStyle(CValue paragraph, bool setSpacing, CValue numLvl) {
    if (setSpacing) {
        setSpacingAfter(paragraph, 20);
    }
    if (!numLvl.IsUndefined()) {
        setNumbering(paragraph, numLvl);
    }
}

CValue createRequisitesParagraph(CValue api, string title, string details, CValue numLvl, bool setSpacing = true, bool setTitleBold = true) {
    CValue paragraph = api.Call("CreateParagraph");
    CValue titleRun = paragraph.Call("AddText", (title + ": ").c_str());
    if (setTitleBold) {
        titleRun.Call("SetBold", true);
    } else {
        titleRun.Call("SetItalic", true);
    }
    CValue detailsRun = paragraph.Call("AddText", details.c_str());
    detailsRun.Call("SetItalic", true);
    setupRequisitesStyle(paragraph, setSpacing, numLvl);
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
    }

    // fill items
    json emptyItem;
    for (const auto& field : tableFields) {
        emptyItem[field] = "...";
    }
    items.push_back(emptyItem);

    const int tableFieldsSize = sizeof(tableFields) / sizeof(tableFields[0]);
    for (int i = 0; i < (int)items.size(); i++) {
        CValue row = table.Call("GetRow", i + 1);
        for (int j = 0; j < tableFieldsSize; j++) {
            CValue cell = getCellContent(row.Call("GetCell", j));
            json value = items[i][tableFields[j]];
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

int main() {
    // parse JSON
    string jsonPath = U_TO_UTF8(NSUtils::GetResourcesDirectory()) + "/data/invoice_response.json";
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
    CValue textPr = document.Call("GetDefaultTextPr");
    textPr.Call("SetFontSize", 24);
    textPr.Call("SetFontFamily", "Times New Roman");

    // DOCUMENT HEADER
    CValue header = document.Call("GetElement", 0);
    fillHeader(header, "INVOICE");

    // document requisites
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Invoice No.", data["invoice"]["number"].get<string>(), CValue::CreateUndefined())
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Date", data["invoice"]["date"].get<string>(), CValue::CreateUndefined(), false)
    );

    // bullet numbering
    CValue bulletNumbering = document.Call("CreateNumbering", "bullet");
    CValue numLvl1 = bulletNumbering.Call("GetLevel", 0);

    // SELLER INFORMATION
    CValue sellerHeader = createDetailsHeader(api, "SELLER INFORMATION");
    document.Call("Push", sellerHeader);

    // seller details
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Company Name", data["seller"]["company_name"].get<string>(), numLvl1)
    );
    document.Call(
        "Push", createRequisitesParagraph(api, "Address", data["seller"]["address"].get<string>(), numLvl1)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Tax ID (TIN)", data["seller"]["tin"].get<string>(), numLvl1)
    );
    document.Call("Push", createRequisitesParagraph(api, "Bank Details", "", numLvl1));

    // bank details
    CValue numLvl2 = bulletNumbering.Call("GetLevel", 1);
    numLvl2.Call("SetCustomType", "none", "", "left");
    numLvl2.Call("SetSuff", "space");

    document.Call(
        "Push",
        createRequisitesParagraph(api, "Bank Name", data["seller"]["bank_details"]["bank_name"].get<string>(), numLvl2, true, false)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Account Number", data["seller"]["bank_details"]["account_number"].get<string>(), numLvl2, true, false)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "SWIFT Code", data["seller"]["bank_details"]["swift_code"].get<string>(), numLvl2, false, false)
    );

    // BUYER INFORMATION
    CValue buyerHeader = createDetailsHeader(api, "BUYER INFORMATION");
    document.Call("Push", buyerHeader);

    // buyer details
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Company Name", data["buyer"]["company_name"].get<string>(), numLvl1)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Address", data["buyer"]["address"].get<string>(), numLvl1)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Tax ID (TIN)", data["buyer"]["tin"].get<string>(), numLvl1, false)
    );

    // TABLE OF ITEMS
    CValue tableHeader = api.Call("CreateParagraph");
    fillHeader(tableHeader, "TABLE OF ITEMS");
    document.Call("Push", tableHeader);

    // table content
    json items = data["items"];
    CValue itemsTable = api.Call("CreateTable", 4, (int)items.size() + 2);
    document.Call("Push", itemsTable);
    setupTableStyle(document, itemsTable);
    fillTableContent(itemsTable, items);

    // TOTALS
    CValue totals = createDetailsHeader(api, "TOTALS");
    document.Call("Push", totals);
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Subtotal", "$" + to_string(data["totals"]["subtotal"].get<int>()), numLvl1)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Tax (20% VAT)", "$" + to_string(data["totals"]["tax"].get<int>()), numLvl1)
    );
    document.Call(
        "Push",
        createRequisitesParagraph(api, "Total Amount Due", "$" + to_string(data["totals"]["total_due"].get<int>()), numLvl1, false)
    );

    // SIGNATURE
    CValue signHeader = api.Call("CreateParagraph");
    signHeader.Call("AddText", "Signature:");
    signHeader.Call("SetBold", true);
    document.Call("Push", signHeader);

    CValue signDetails = api.Call("CreateParagraph");
    signDetails.Call(
        "AddText",
        (data["seller"]["authorized_person"].get<string>() + ", " + data["seller"]["position"].get<string>()).c_str()
    );
    signDetails.Call("AddLineBreak");
    signDetails.Call("AddText", data["seller"]["company_name"].get<string>().c_str());
    document.Call("Push", signDetails);

    // Save and close
    builder.SaveFile(OFFICESTUDIO_FILE_DOCUMENT_OFORM_PDF, resultPath);
    builder.CloseFile();
    CDocBuilder::Dispose();
    return 0;
}
