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

import docbuilder.*;

import java.io.FileReader;
import java.text.NumberFormat;
import java.util.Locale;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class Program {
    public static void main(String[] args) throws Exception {
        String resultPath = "result.docx";
        String resourcesDir = "../../resources";

        createCommercialOffer(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createCommercialOffer(String resultPath, String resourcesDir) throws Exception {
        // parse JSON
        String jsonPath = resourcesDir + "/data/commercial_offer_data.json";
        JSONObject data = (JSONObject) new JSONParser().parse(new FileReader(jsonPath));

        // init docbuilder and create new docx file
        int doctype = FileTypes.Document.DOCX;
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");
        CDocBuilderValue document = api.call("GetDocument");

        // DOCUMENT STYLE
        CDocBuilderValue paraPr = document.call("GetDefaultParaPr");
        paraPr.call("SetSpacingAfter", 100);
        CDocBuilderValue textPr = document.call("GetDefaultTextPr");
        textPr.call("SetFontSize", 24);
        textPr.call("SetFontFamily", "Times New Roman");

        // DOCUMENT HEADER
        CDocBuilderValue header = document.call("GetElement", 0);
        fillHeader(header, "COMMERCIAL OFFER TEMPLATE");

        // document requisites
        JSONObject offer = (JSONObject) data.get("offer");
        document.call(
            "Push",
            createRequisitesParagraph(api, "Offer No.", offer.get("number").toString(), null)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Date", offer.get("date").toString(), null, false, true)
        );

        // bullet numbering
        CDocBuilderValue bulletNumbering = document.call("CreateNumbering", "bullet");
        CDocBuilderValue bNumLvl = bulletNumbering.call("GetLevel", 0);

        // SELLER INFORMATION
        CDocBuilderValue sellerHeader = createDetailsHeader(api, "SELLER INFORMATION");
        document.call("Push", sellerHeader);

        // seller details
        JSONObject seller = (JSONObject) data.get("seller");
        document.call(
            "Push",
            createRequisitesParagraph(api, "Company Name", seller.get("company_name").toString(), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Address", seller.get("address").toString(), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Tax ID (TIN)", seller.get("tin").toString(), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Contact Information", "", bNumLvl)
        );

        // contact details
        JSONObject contact = (JSONObject) seller.get("contact");
        document.call(
            "Push",
            createRequisitesParagraph(api, "Phone", contact.get("phone").toString(), bNumLvl, true, false)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Email", contact.get("email").toString(), bNumLvl, false, false)
        );

        // BUYER INFORMATION
        CDocBuilderValue buyerHeader = createDetailsHeader(api, "BUYER INFORMATION");
        document.call("Push", buyerHeader);

        // buyer details
        JSONObject buyer = (JSONObject) data.get("buyer");
        document.call(
            "Push",
            createRequisitesParagraph(api, "Company Name", buyer.get("company_name").toString(), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Address", buyer.get("address").toString(), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Contact Person", buyer.get("contact_person").toString(), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Email", buyer.get("email").toString(), bNumLvl, false, true)
        );

        // OFFER DETAILS
        CDocBuilderValue tableHeader = api.call("CreateParagraph");
        fillHeader(tableHeader, "OFFER DETAILS");
        document.call("Push", tableHeader);

        // table content
        JSONArray offerDetails = (JSONArray) data.get("offer_details");
        CDocBuilderValue itemsTable = api.call("CreateTable", 4, offerDetails.size() + 2);
        document.call("Push", itemsTable);
        setupTableStyle(document, itemsTable);
        fillTableContent(itemsTable, offerDetails);

        // TOTALS
        CDocBuilderValue totals = createDetailsHeader(api, "TOTALS");
        document.call("Push", totals);
        JSONObject totalsData = (JSONObject) data.get("totals");
        document.call(
            "Push",
            createRequisitesParagraph(api, "Subtotal", formatSum(((Long) totalsData.get("subtotal")).intValue()), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Discount", formatSum(((Long) totalsData.get("discount")).intValue()), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Tax (e.g., 20% VAT)", formatSum(((Long) totalsData.get("tax")).intValue()), bNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Total Amount", formatSum(((Long) totalsData.get("total")).intValue()), bNumLvl, false, true)
        );

        // TERMS AND CONDITIONS
        CDocBuilderValue termsHeader = createDetailsHeader(api, "TERMS AND CONDITIONS");
        document.call("Push", termsHeader);

        // numbering
        CDocBuilderValue numbering = document.call("CreateNumbering", "numbered");
        CDocBuilderValue dNumLvl = numbering.call("GetLevel", 0);
        dNumLvl.call("SetCustomType", "decimal", "%1.", "left");

        JSONObject terms = (JSONObject) data.get("terms_and_conditions");
        document.call(
            "Push",
            createRequisitesParagraph(api, "Validity Period", terms.get("validity_period").toString(), dNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Payment Terms", terms.get("payment_terms").toString(), dNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Delivery Terms", terms.get("delivery_terms").toString(), dNumLvl)
        );
        document.call(
            "Push",
            createRequisitesParagraph(api, "Additional Notes", terms.get("additional_notes").toString(), dNumLvl, false, true)
        );

        // SIGNATURE
        CDocBuilderValue signHeader = api.call("CreateParagraph");
        signHeader.call("AddText", "Signature:");
        signHeader.call("SetBold", true);
        document.call("Push", signHeader);

        CDocBuilderValue signDetails = api.call("CreateParagraph");
        JSONObject authorizedPerson = (JSONObject) seller.get("authorized_person");
        signDetails.call(
            "AddText",
            String.format("%s, %s",
                authorizedPerson.get("full_name").toString(),
                authorizedPerson.get("position").toString())
        );
        signDetails.call("AddLineBreak");
        signDetails.call("AddText", seller.get("company_name").toString());
        document.call("Push", signDetails);

        // save and close
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }

    private static void fillHeader(CDocBuilderValue paragraph, String text) {
        paragraph.call("AddText", text);
        paragraph.call("SetFontSize", 28);
        paragraph.call("SetBold", true);
        setSpacingAfter(paragraph, 50);
    }

    private static void setSpacingAfter(CDocBuilderValue paragraph, int spacing) {
        paragraph.call("SetSpacingAfter", spacing);
    }

    private static void setNumbering(CDocBuilderValue paragraph, CDocBuilderValue numLvl) {
        paragraph.call("SetNumbering", numLvl);
    }

    private static CDocBuilderValue createDetailsHeader(CDocBuilderValue api, String text) {
        CDocBuilderValue paragraph = api.call("CreateParagraph");
        paragraph.call("AddText", text);
        paragraph.call("SetBold", true);
        paragraph.call("SetItalic", true);
        setSpacingAfter(paragraph, 40);
        return paragraph;
    }

    private static void setupRequisitesStyle(CDocBuilderValue paragraph, CDocBuilderValue numLvl, boolean setSpacing) {
        if (setSpacing) {
            setSpacingAfter(paragraph, 20);
        }
        if (numLvl != null) {
            setNumbering(paragraph, numLvl);
        }
    }

    private static CDocBuilderValue createRequisitesParagraph(CDocBuilderValue api, String title, String details, CDocBuilderValue numLvl, boolean setSpacing, boolean setTitleBold) {
        CDocBuilderValue paragraph = api.call("CreateParagraph");
        CDocBuilderValue titleRun = paragraph.call("AddText", title + ": ");
        if (setTitleBold) {
            titleRun.call("SetBold", true);
        } else {
            titleRun.call("SetItalic", true);
        }
        CDocBuilderValue detailsRun = paragraph.call("AddText", details);
        detailsRun.call("SetItalic", true);
        setupRequisitesStyle(paragraph, numLvl, setSpacing);
        return paragraph;
    }

    private static CDocBuilderValue createRequisitesParagraph(CDocBuilderValue api, String title, String details, CDocBuilderValue numLvl) {
        return createRequisitesParagraph(api, title, details, numLvl, true, true);
    }

    private static String formatSum(int value) {
        return "$" + NumberFormat.getNumberInstance(Locale.US).format(value);
    }

    private static void setupTableStyle(CDocBuilderValue document, CDocBuilderValue table) {
        // table size
        table.call("SetWidth", "percent", 100);
        table.call("Select");
        CDocBuilderValue tableRange = document.call("GetRangeBySelect");
        CDocBuilderValue tableParagraphs = tableRange.call("GetAllParagraphs");

        for (int i = 0; i < tableParagraphs.getLength(); i++) {
            CDocBuilderValue paraPr = tableParagraphs.get(i).call("GetParaPr");
            paraPr.call("SetSpacingBefore", 40);
            paraPr.call("SetSpacingAfter", 40);
        }

        // table borders
        table.call("SetTableBorderTop", "single", 4, 0, 0, 0, 0);
        table.call("SetTableBorderBottom", "single", 4, 0, 0, 0, 0);
        table.call("SetTableBorderLeft", "single", 4, 0, 0, 0, 0);
        table.call("SetTableBorderRight", "single", 4, 0, 0, 0, 0);
        table.call("SetTableBorderInsideH", "single", 4, 0, 0, 0, 0);
        table.call("SetTableBorderInsideV", "single", 4, 0, 0, 0, 0);
    }

    private static CDocBuilderValue getCellContent(CDocBuilderValue cell) {
        return cell.call("GetContent").call("GetElement", 0);
    }

    private static void fillTableContent(CDocBuilderValue table, JSONArray items) {
        String[] tableHeaders = {"Description", "Quantity", "Unit Price", "Total"};
        String[] tableFields = {"description", "quantity", "unit_price", "total"};

        // fill table header
        CDocBuilderValue headerRow = table.call("GetRow", 0);
        for (int i = 0; i < tableHeaders.length; i++) {
            CDocBuilderValue headerCell = getCellContent(headerRow.call("GetCell", i));
            headerCell.call("AddText", tableHeaders[i]);
            headerCell.call("SetBold", true);
            headerCell.call("SetJc", "center");
        }

        // fill items
        for (int i = 0; i < items.size(); i++) {
            CDocBuilderValue row = table.call("GetRow", i + 1);
            JSONObject item = (JSONObject) items.get(i);

            for (int j = 0; j < tableFields.length; j++) {
                CDocBuilderValue cell = getCellContent(row.call("GetCell", j));
                String fieldName = tableFields[j];

                if (fieldName.equals("unit_price") || fieldName.equals("total")) {
                    cell.call("AddText", formatSum(((Long) item.get(fieldName)).intValue()));
                } else {
                    cell.call("AddText", item.get(fieldName).toString());
                }
            }
        }

        // fill last row with dots
        CDocBuilderValue lastRow = table.call("GetRow", items.size() + 1);
        for (int j = 0; j < tableFields.length; j++) {
            CDocBuilderValue cell = getCellContent(lastRow.call("GetCell", j));
            cell.call("AddText", "...");
        }
    }
}
