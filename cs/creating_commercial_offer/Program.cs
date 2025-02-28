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

using docbuilder_net;

using OfficeFileTypes = docbuilder_net.FileTypes;
using CValue = docbuilder_net.CDocBuilderValue;
using CContext = docbuilder_net.CDocBuilderContext;

using System.Collections.Generic;
using System.Text.Json;
using System.IO;
using System.Globalization;

namespace Sample
{
    public class CommercialOffer
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.docx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateCommercialOffer(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateCommercialOffer(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string jsonPath = resourcesDir + "/data/commercial_offer_data.json";
            string json = File.ReadAllText(jsonPath);
            JsonData data = JsonSerializer.Deserialize<JsonData>(json);

            // Init DocBuilder
            var doctype = (int)OfficeFileTypes.Document.DOCX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

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
            FillHeader(header, "COMMERCIAL OFFER TEMPLATE");

            // document requisites
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Offer No.", data.offer.number)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Date", data.offer.date, CValue.CreateUndefined(), false)
            );

            // bullet numbering
            CValue bulletNumbering = document.Call("CreateNumbering", "bullet");
            CValue bNumLvl = bulletNumbering.Call("GetLevel", 0);

            // SELLER INFORMATION
            CValue sellerHeader = CreateDetailsHeader(api, "SELLER INFORMATION");
            document.Call("Push", sellerHeader);

            // seller details
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Company Name", data.seller.company_name, bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Address", data.seller.address, bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Tax ID (TIN)", data.seller.tin, bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Contact Information", "", bNumLvl)
            );

            // contact details
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Phone", data.seller.contact.phone, bNumLvl, true, false)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Email", data.seller.contact.email, bNumLvl, false, false)
            );

            // BUYER INFORMATION
            CValue buyerHeader = CreateDetailsHeader(api, "BUYER INFORMATION");
            document.Call("Push", buyerHeader);

            // buyer details
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Company Name", data.buyer.company_name, bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Address", data.buyer.address, bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Contact Person", data.buyer.contact_person, bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Email", data.buyer.email, bNumLvl, false)
            );

            // OFFER DETAILS
            CValue tableHeader = api.Call("CreateParagraph");
            FillHeader(tableHeader, "OFFER DETAILS");
            document.Call("Push", tableHeader);

            // table content
            List<OfferDetail> offerDetails = data.offer_details;
            CValue itemsTable = api.Call("CreateTable", 4, offerDetails.Count + 1);
            document.Call("Push", itemsTable);
            SetupTableStyle(document, itemsTable);
            FillTableContent(itemsTable, offerDetails);

            // TOTALS
            CValue totals = CreateDetailsHeader(api, "TOTALS");
            document.Call("Push", totals);
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Subtotal", FormatSum(data.totals.subtotal), bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Discount", FormatSum(data.totals.discount), bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Tax (e.g., 20% VAT)", FormatSum(data.totals.tax), bNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Total Amount", FormatSum(data.totals.total), bNumLvl, false)
            );

            // TERMS AND CONDITIONS
            CValue sellerHeader2 = CreateDetailsHeader(api, "TERMS AND CONDITIONS");
            document.Call("Push", sellerHeader2);

            // numbering
            CValue numbering = document.Call("CreateNumbering", "numbered");
            CValue dNumLvl = numbering.Call("GetLevel", 0);
            dNumLvl.Call("SetCustomType", "decimal", "%1.", "left");

            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Validity Period", data.terms_and_conditions.validity_period, dNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Payment Terms", data.terms_and_conditions.payment_terms, dNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Delivery Terms", data.terms_and_conditions.delivery_terms, dNumLvl)
            );
            document.Call(
                "Push",
                CreateRequisitesParagraph(api, "Additional Notes", data.terms_and_conditions.additional_notes, dNumLvl, false)
            );

            // SIGNATURE
            CValue signHeader = api.Call("CreateParagraph");
            signHeader.Call("AddText", "Signature:");
            signHeader.Call("SetBold", true);
            document.Call("Push", signHeader);

            CValue signDetails = api.Call("CreateParagraph");
            signDetails.Call(
                "AddText",
                $"{data.seller.authorized_person.full_name}, {data.seller.authorized_person.position}"
            );
            signDetails.Call("AddLineBreak");
            signDetails.Call("AddText", data.seller.company_name);
            document.Call("Push", signDetails);

            // Save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void SetSpacingAfter(CValue paragraph, int spacing)
        {
            paragraph.Call("SetSpacingAfter", spacing);
        }

        public static void FillHeader(CValue paragraph, string text)
        {
            paragraph.Call("AddText", text);
            paragraph.Call("SetFontSize", 28);
            paragraph.Call("SetBold", true);
            SetSpacingAfter(paragraph, 50);
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
            SetSpacingAfter(paragraph, 40);
            return paragraph;
        }

        public static void SetupRequisitesStyle(CValue paragraph, CValue numLvl, bool setSpacing)
        {
            if (setSpacing)
            {
                SetSpacingAfter(paragraph, 20);
            }
            if (!numLvl.IsUndefined())
            {
                SetNumbering(paragraph, numLvl);
            }
        }

        public static string FormatSum(int value)
        {
            CultureInfo culture = new ("en-US");
            return value.ToString("C0", culture);
        }

        public static CValue CreateRequisitesParagraph(CValue api, string title, string details, CValue numLvl = null, bool setSpacing = true, bool setTitleBold = true)
        {
            if (numLvl == null)
            {
                numLvl = CValue.CreateUndefined();
            }

            CValue paragraph = api.Call("CreateParagraph");
            CValue titleRun = paragraph.Call("AddText", $"{title}: ");
            if (setTitleBold)
            {
                titleRun.Call("SetBold", true);
            }
            else
            {
                titleRun.Call("SetItalic", true);
            }
            CValue detailsRun = paragraph.Call("AddText", details);
            detailsRun.Call("SetItalic", true);
            SetupRequisitesStyle(paragraph, numLvl, setSpacing);
            return paragraph;
        }

        public static void SetupTableStyle(CValue document, CValue table)
        {
            // table size
            table.Call("SetWidth", "percent", 100);
            table.Call("Select");
            CValue tableRange = document.Call("GetRangeBySelect");
            CValue tableParagraphs = tableRange.Call("GetAllParagraphs");
            for (int i = 0; i < (int)tableParagraphs.GetLength(); i++)
            {
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

        public static void FillTableContent(CValue table, List<OfferDetail> items)
        {
            string[] tableHeaders = { "Description", "Quantity", "Unit Price", "Total" };
            string[] tableFields = { "description", "quantity", "unit_price", "total" };

            // fill table header
            CValue headerRow = table.Call("GetRow", 0);
            for (int i = 0; i < tableHeaders.Length; i++)
            {
                CValue headerCell = GetCellContent(headerRow.Call("GetCell", i));
                headerCell.Call("AddText", tableHeaders[i]);
                headerCell.Call("SetBold", true);
                headerCell.Call("SetJc", "center");
            }

            // fill items
            for (int i = 0; i < items.Count; i++)
            {
                CValue row = table.Call("GetRow", i + 1);
                for (int j = 0; j < tableFields.Length; j++)
                {
                    CValue cell = GetCellContent(row.Call("GetCell", j));
                    string key = tableFields[j];

                    // Handle different field types
                    if (key == "unit_price" || key == "total")
                    {
                        int value = (int)items[i].GetType().GetProperty(key).GetValue(items[i]);
                        cell.Call("AddText", FormatSum(value));
                    }
                    else
                    {
                        object value = items[i].GetType().GetProperty(key).GetValue(items[i]);
                        string strValue = value.ToString();
                        cell.Call("AddText", strValue);
                    }
                }
            }
        }
    }

    // Define classes to represent the JSON structure
    public class JsonData
    {
        public Offer offer { get; set; }
        public Seller seller { get; set; }
        public Buyer buyer { get; set; }
        public List<OfferDetail> offer_details { get; set; }
        public Totals totals { get; set; }
        public TermsAndConditions terms_and_conditions { get; set; }
    }

    public class Offer
    {
        public string number { get; set; }
        public string date { get; set; }
    }

    public class Seller
    {
        public string company_name { get; set; }
        public string address { get; set; }
        public string tin { get; set; }
        public Contact contact { get; set; }
        public AuthorizedPerson authorized_person { get; set; }
    }

    public class Contact
    {
        public string phone { get; set; }
        public string email { get; set; }
    }

    public class AuthorizedPerson
    {
        public string full_name { get; set; }
        public string position { get; set; }
    }

    public class Buyer
    {
        public string company_name { get; set; }
        public string address { get; set; }
        public string contact_person { get; set; }
        public string email { get; set; }
    }

    public class OfferDetail
    {
        public string description { get; set; }
        public int quantity { get; set; }
        public int unit_price { get; set; }
        public int total { get; set; }
    }

    public class Totals
    {
        public int subtotal { get; set; }
        public int discount { get; set; }
        public int tax { get; set; }
        public int total { get; set; }
    }

    public class TermsAndConditions
    {
        public string validity_period { get; set; }
        public string payment_terms { get; set; }
        public string delivery_terms { get; set; }
        public string additional_notes { get; set; }
    }
}
