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

using docbuilder_net;

using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;

using System.Collections.Generic;
using System.Text.Json;
using System.IO;

namespace Sample
{
    public class CreatingInvoice
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.pdf";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateInvoice(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateInvoice(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string jsonPath = resourcesDir + "/data/invoice_response.json";
            string json = File.ReadAllText(jsonPath);
            JsonData data = JsonSerializer.Deserialize<JsonData>(json);

            // Init DocBuilder
            var doctype = (int)OfficeFileTypes.Document.OFORM_PDF;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

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
            FillHeader(header, "INVOICE");

            // document requisites
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Invoice No.", data.invoice.number, CValue.CreateUndefined())
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Date", data.invoice.date, CValue.CreateUndefined(), false)
            );

            // bullet numbering
            CValue bulletNumbering = document.Call("CreateNumbering", "bullet");
            CValue numLvl1 = bulletNumbering.Call("GetLevel", 0);

            // SELLER INFORMATION
            CValue sellerHeader = CreateDetailsHeader(api, "SELLER INFORMATION");
            document.Call("Push", sellerHeader);

            // seller details
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Company Name", data.seller.company_name, numLvl1)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Address", data.seller.address, numLvl1)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Tax ID (TIN)", data.seller.tin, numLvl1)
            );
            document.Call("Push", CreateRequisitesParagraph(api, "Bank Details", "", numLvl1));

            // bank details
            CValue numLvl2 = bulletNumbering.Call("GetLevel", 1);
            numLvl2.Call("SetCustomType", "none", "", "left");
            numLvl2.Call("SetSuff", "space");

            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Bank Name", data.seller.bank_details.bank_name, numLvl2, true, false)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Account Number", data.seller.bank_details.account_number, numLvl2, true, false)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "SWIFT Code", data.seller.bank_details.swift_code, numLvl2, false, false)
            );

            // BUYER INFORMATION
            CValue buyerHeader = CreateDetailsHeader(api, "BUYER INFORMATION");
            document.Call("Push", buyerHeader);

            // buyer details
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Company Name", data.buyer.company_name, numLvl1)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Address", data.buyer.address, numLvl1)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Tax ID (TIN)", data.buyer.tin, numLvl1, false)
            );

            // TABLE OF ITEMS
            CValue tableHeader = api.Call("CreateParagraph");
            FillHeader(tableHeader, "TABLE OF ITEMS");
            document.Call("Push", tableHeader);

            // table content
            List<Item> items = data.items;
            CValue itemsTable = api.Call("CreateTable", 4, items.Count + 2);
            document.Call("Push", itemsTable);
            SetupTableStyle(document, itemsTable);
            FillTableContent(itemsTable, items);

            // TOTALS
            CValue totals = CreateDetailsHeader(api, "TOTALS");
            document.Call("Push", totals);
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Subtotal", $"${data.totals.subtotal}", numLvl1, true)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Tax (20% VAT)", $"${data.totals.tax}", numLvl1, true)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Total Amount Due", $"${data.totals.total_due}", numLvl1, false)
            );

            // SIGNATURE
            CValue signHeader = api.Call("CreateParagraph");
            signHeader.Call("AddText", "Signature:");
            signHeader.Call("SetBold", true);
            document.Call("Push", signHeader);

            CValue signDetails = api.Call("CreateParagraph");
            signDetails.Call("AddText", $"{data.seller.authorized_person}, {data.seller.position}");
            signDetails.Call("AddLineBreak");
            signDetails.Call("AddText", data.seller.company_name);
            document.Call("Push", signDetails);

            // Save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void FillHeader(CValue paragraph, string text)
        {
            paragraph.Call("AddText", text);
            paragraph.Call("SetFontSize", 28);
            paragraph.Call("SetBold", true);
        }

        public static void SetSpacingAfter(CValue paragraph, int spacing)
        {
            paragraph.Call("SetSpacingAfter", spacing);
        }

        public static void SetNumbering(CValue paragraph, CValue numLvl)
        {
            paragraph.Call("SetNumbering", numLvl);
        }

        public static CValue CreateDetailsHeader(CValue api, string text)
        {
            CValue paragraph = api.Call("CreateParagraph");
            paragraph.Call("AddText", text);
            paragraph.Call("SetBold", true);
            paragraph.Call("SetItalic", true);
            SetSpacingAfter(paragraph, 50);
            return paragraph;
        }

        public static void SetupRequisitesStyle(CValue paragraph, bool setSpacing, CValue numLvl)
        {
            if (setSpacing) {
                SetSpacingAfter(paragraph, 20);
            }
            if (!numLvl.IsUndefined()) {
                SetNumbering(paragraph, numLvl);
            }
        }

        public static CValue CreateRequisitesParagraph(CValue api, string title, string details, CValue numLvl, bool setSpacing = true, bool setTitleBold = true)
        {
            CValue paragraph = api.Call("CreateParagraph");
            CValue titleRun = paragraph.Call("AddText", $"{title}: ");
            if (setTitleBold) {
                titleRun.Call("SetBold", true);
            } else {
                titleRun.Call("SetItalic", true);
            }
            CValue detailsRun = paragraph.Call("AddText", details);
            detailsRun.Call("SetItalic", true);
            SetupRequisitesStyle(paragraph, setSpacing, numLvl);
            return paragraph;
        }

        public static void SetupTableStyle(CValue document, CValue table)
        {
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

        public static CValue GetCellContent(CValue cell)
        {
            return cell.Call("GetContent").Call("GetElement", 0);
        }

        public static void FillTableContent(CValue table, List<Item> items)
        {
            string[] tableHeaders = {"Description", "Quantity", "Unit Price", "Total"};
            string[] tableFields = {"description", "quantity", "unit_price", "total"};

            // fill table header
            CValue headerRow = table.Call("GetRow", 0);
            for (int i = 0; i < tableHeaders.Length; i++) {
                CValue headerCell = GetCellContent(headerRow.Call("GetCell", i));
                headerCell.Call("AddText", tableHeaders[i]);
                headerCell.Call("SetBold", true);
            }

            // fill items
            for (int i = 0; i < items.Count; i++) {
                CValue row = table.Call("GetRow", i + 1);
                for (int j = 0; j < tableFields.Length; j++) {
                    CValue cell = GetCellContent(row.Call("GetCell", j));
                    string value = items[i].GetType().GetProperty(tableFields[j]).GetValue(items[i]).ToString();
                    cell.Call("AddText", value);
                }
            }

            // fill last row with dots
            CValue lastRow = table.Call("GetRow", items.Count + 1);
            for (int j = 0; j < tableFields.Length; j++) {
                CValue cell = GetCellContent(lastRow.Call("GetCell", j));
                cell.Call("AddText", "...");
            }
        }
    }

    // Define classes to represent the JSON structure
    public class JsonData
    {
        public Invoice invoice { get; set; }
        public Seller seller { get; set; }
        public Buyer buyer { get; set; }
        public List<Item> items { get; set; }
        public Totals totals { get; set; }
    }

    public class Invoice
    {
        public string number { get; set; }
        public string date { get; set; }
    }

    public class Seller
    {
        public string company_name { get; set; }
        public string address { get; set; }
        public string tin { get; set; }
        public BankDetails bank_details { get; set; }
        public string authorized_person { get; set; }
        public string position { get; set; }
    }

    public class BankDetails
    {
        public string bank_name { get; set; }
        public string account_number { get; set; }
        public string swift_code { get; set; }
    }

    public class Buyer
    {
        public string company_name { get; set; }
        public string address { get; set; }
        public string tin { get; set; }
    }

    public class Item
    {
        public string description { get; set; }
        public int quantity { get; set; }
        public int unit_price { get; set; }
        public int total { get; set; }
    }

    public class Totals
    {
        public int subtotal { get; set; }
        public int tax { get; set; }
        public int total_due { get; set; }
    }
}
