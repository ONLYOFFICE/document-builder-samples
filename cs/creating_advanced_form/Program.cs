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

namespace Sample
{
    public class CreatingAdvancedForm
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.docx";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateAdvancedForm(workDirectory, resultPath);
        }

        public static void CreateAdvancedForm(string workDirectory, string resultPath)
        {
            var doctype = (int)OfficeFileTypes.Document.DOCX;

            // Init DocBuilder
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];

            // Create advanced form
            CValue document = api.Call("GetDocument");
            CValue table = CreateFullWidthTable(api, 1, 2, 255);
            CValue paragraph = GetTableCellParagraph(table, 0, 0);
            AddTextToParagraph(paragraph, "PURCHASE ORDER", 36, true);
            paragraph = GetTableCellParagraph(table, 0, 1);
            AddTextToParagraph(paragraph, "Serial # ", 25, true);

            CValue textForm = api.Call("CreateTextForm");
            SetTextFormProperties(textForm, "Serial", "Enter serial number", false, "Serial", true, 6, 1, false, false);
            AddTextFormToParagraph(paragraph, textForm, 25, "left", true, 255);
            document.Call("Push", table);

            CValue pictureForm = api.Call("CreatePictureForm");
            SetPictureFormProperties(pictureForm, "Photo", "Upload company logo", false, "Photo", "tooBig", false, false, 0, 0);
            paragraph = api.Call("CreateParagraph");
            paragraph.Call("AddElement", pictureForm);
            document.Call("Push", paragraph);

            textForm = api.Call("CreateTextForm");
            SetTextFormProperties(textForm, "Company Name", "Enter company name", false, "Company Name", true, 20, 1, false, false);
            paragraph = api.Call("CreateParagraph");
            AddTextFormToParagraph(paragraph, textForm, 35, "left", false, 255);
            document.Call("Push", paragraph);

            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "Date: ", 25, true);
            textForm = api.Call("CreateTextForm");
            SetTextFormProperties(textForm, "Date", "Date", false, "DD.MM.YYYY", true, 10, 1, false, false);
            AddTextFormToParagraph(paragraph, textForm, 25, "left", true, 255);
            document.Call("Push", paragraph);

            paragraph = api.Call("CreateParagraph");
            AddTextToParagraph(paragraph, "To:", 35, true);
            document.Call("Push", paragraph);

            table = CreateFullWidthTable(api, 1, 1, 200);
            paragraph = GetTableCellParagraph(table, 0, 0);
            textForm = api.Call("CreateTextForm");
            SetTextFormProperties(textForm, "Recipient", "Recipient", false, "Recipient", true, 25, 1, false, false);
            AddTextFormToParagraph(paragraph, textForm, 32, "left", false, 255);
            document.Call("Push", table);

            table = CreateFullWidthTable(api, 10, 2, 200);
            table.Call("GetRow", 0).Call("SetBackgroundColor", 245, 245, 245, false);
            CValue cell = table.Call("GetCell", 0, 0);
            cell.Call("SetWidth", "percent", 30);
            paragraph = GetTableCellParagraph(table, 0, 0);
            AddTextToParagraph(paragraph, "Qty.", 30, true);
            paragraph = GetTableCellParagraph(table, 0, 1);
            AddTextToParagraph(paragraph, "Description", 30, true);

            for (var i = 1; i < 10; i++)
            {
                CValue tempParagraph = GetTableCellParagraph(table, i, 0);
                CValue tempTextForm = api.Call("CreateTextForm");
                SetTextFormProperties(tempTextForm, "Qty" + i, "Qty" + i, false, " ", true, 9, 1, false, false);
                AddTextFormToParagraph(tempParagraph, tempTextForm, 30, "left", false, 255);

                tempParagraph = GetTableCellParagraph(table, i, 1);
                tempTextForm = api.Call("CreateTextForm");
                SetTextFormProperties(tempTextForm, "Description" + i, "Description" + i, false, " ", true, 22, 1, false, false);
                AddTextFormToParagraph(tempParagraph, tempTextForm, 30, "left", false, 255);
            }

            document.Call("Push", table);
            document.Call("RemoveElement", 0);
            document.Call("RemoveElement", 1);

            // Save file and close DocBuilder
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();

            CDocBuilder.Destroy();
        }

        public static CValue CreateFullWidthTable(CValue api, int rows, int cols, int borderColor)
        {
            CValue table = api.Call("CreateTable", cols, rows);
            table.Call("SetWidth", "percent", 100);
            SetTableBorders(table, borderColor);
            return table;
        }

        public static void SetTableBorders(CValue table, int borderColor)
        {
            table.Call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
            table.Call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
        }

        public static CValue GetTableCellParagraph(CValue table, int row, int col)
        {
            return table.Call("GetCell", row, col).Call("GetContent").Call("GetElement", 0);
        }

        public static void AddTextToParagraph(CValue paragraph, string text, int fontSize, bool isBold)
        {
            paragraph.Call("AddText", text);
            paragraph.Call("SetFontSize", fontSize);
            paragraph.Call("SetBold", isBold);
        }

        public static void SetPictureFormProperties(CValue pictureForm, string key, string tip, bool required, string placeholder, string scaleFlag, bool lockAspectRatio, bool respectBorders, int shiftX, int shiftY)
        {
            pictureForm.Call("SetFormKey", key);
            pictureForm.Call("SetTipText", tip);
            pictureForm.Call("SetRequired", required);
            pictureForm.Call("SetPlaceholderText", placeholder);
            pictureForm.Call("SetScaleFlag", scaleFlag);
            pictureForm.Call("SetLockAspectRatio", lockAspectRatio);
            pictureForm.Call("SetRespectBorders", respectBorders);
            pictureForm.Call("SetPicturePosition", shiftX, shiftY);
        }

        public static void SetTextFormProperties(CValue textForm, string key, string tip, bool required, string placeholder, bool comb, int maxCharacters, int cellWidth, bool multiLine, bool autoFit)
        {
            textForm.Call("SetFormKey", key);
            textForm.Call("SetTipText", tip);
            textForm.Call("SetRequired", required);
            textForm.Call("SetPlaceholderText", placeholder);
            textForm.Call("SetComb", comb);
            textForm.Call("SetCharactersLimit", maxCharacters);
            textForm.Call("SetCellWidth", cellWidth);
            textForm.Call("SetCellWidth", multiLine);
            textForm.Call("SetMultiline", autoFit);
        }

        public static void AddTextFormToParagraph(CValue paragraph, CValue textForm, int fontSize, string jc, bool hasBorder, int borderColor)
        {
            if (hasBorder)
            {
                textForm.Call("SetBorderColor", borderColor, borderColor, borderColor);
            }
            paragraph.Call("AddElement", textForm);
            paragraph.Call("SetFontSize", fontSize);
            paragraph.Call("SetJc", jc);
        }
    }
}
