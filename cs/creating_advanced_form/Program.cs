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
using CContextScope = docbuilder_net.CDocBuilderContextScope;

namespace Sample
{
    public class CreatingAdvancedForm
    {
        public static void Main(string[] args)
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
            CDocBuilder oBuilder = new CDocBuilder();
            oBuilder.CreateFile(doctype);

            CContext oContext = oBuilder.GetContext();
            CContextScope oScope = oContext.CreateScope();
            CValue oGlobal = oContext.GetGlobal();
            CValue oApi = oGlobal["Api"];

            // Create advanced form
            CValue oDocument = oApi.Call("GetDocument");
            CValue oTable = createFullWidthTable(oApi, 1, 2, 255);
            CValue oParagraph = getTableCellParagraph(oTable, 0, 0);
            addTextToParagraph(oParagraph, "PURCHASE ORDER", 36, true);
            oParagraph = getTableCellParagraph(oTable, 0, 1);
            addTextToParagraph(oParagraph, "Serial # ", 25, true);

            CValue oTextForm = oApi.Call("CreateTextForm");
            setTextFormProperties(oTextForm, "Serial", "Enter serial number", false, "Serial", true, 6, 1, false, false);
            addTextFormToParagraph(oParagraph, oTextForm, 25, "left", true, 255);
            oDocument.Call("Push", oTable);

            CValue oPictureForm = oApi.Call("CreatePictureForm");
            setPictureFormProperties(oPictureForm, "Photo", "Upload company logo", false, "Photo", "tooBig", false, false, 0, 0);
            oParagraph = oApi.Call("CreateParagraph");
            oParagraph.Call("AddElement", oPictureForm);
            oDocument.Call("Push", oParagraph);

            oTextForm = oApi.Call("CreateTextForm");
            setTextFormProperties(oTextForm, "Company Name", "Enter company name", false, "Company Name", true, 20, 1, false, false);
            oParagraph = oApi.Call("CreateParagraph");
            addTextFormToParagraph(oParagraph, oTextForm, 35, "left", false, 255);
            oDocument.Call("Push", oParagraph);

            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "Date: ", 25, true);
            oTextForm = oApi.Call("CreateTextForm");
            setTextFormProperties(oTextForm, "Date", "Date", false, "DD.MM.YYYY", true, 10, 1, false, false);
            addTextFormToParagraph(oParagraph, oTextForm, 25, "left", true, 255);
            oDocument.Call("Push", oParagraph);

            oParagraph = oApi.Call("CreateParagraph");
            addTextToParagraph(oParagraph, "To:", 35, true);
            oDocument.Call("Push", oParagraph);

            oTable = createFullWidthTable(oApi, 1, 1, 200);
            oParagraph = getTableCellParagraph(oTable, 0, 0);
            oTextForm = oApi.Call("CreateTextForm");
            setTextFormProperties(oTextForm, "Recipient", "Recipient", false, "Recipient", true, 25, 1, false, false);
            addTextFormToParagraph(oParagraph, oTextForm, 32, "left", false, 255);
            oDocument.Call("Push", oTable);

            oTable = createFullWidthTable(oApi, 10, 2, 200);
            oTable.Call("GetRow", 0).Call("SetBackgroundColor", 245, 245, 245, false);
            CValue oCell = oTable.Call("GetCell", 0, 0);
            oCell.Call("SetWidth", "percent", 30);
            oParagraph = getTableCellParagraph(oTable, 0, 0);
            addTextToParagraph(oParagraph, "Qty.", 30, true);
            oParagraph = getTableCellParagraph(oTable, 0, 1);
            addTextToParagraph(oParagraph, "Description", 30, true);

            for (var i = 1; i < 10; i++)
            {
                CValue oTempParagraph = getTableCellParagraph(oTable, i, 0);
                CValue oTempTextForm = oApi.Call("CreateTextForm");
                setTextFormProperties(oTempTextForm, "Qty" + i, "Qty" + i, false, " ", true, 9, 1, false, false);
                addTextFormToParagraph(oTempParagraph, oTempTextForm, 30, "left", false, 255);

                oTempParagraph = getTableCellParagraph(oTable, i, 1);
                oTempTextForm = oApi.Call("CreateTextForm");
                setTextFormProperties(oTempTextForm, "Description" + i, "Description" + i, false, " ", true, 22, 1, false, false);
                addTextFormToParagraph(oTempParagraph, oTempTextForm, 30, "left", false, 255);
            }

            oDocument.Call("Push", oTable);
            oDocument.Call("RemoveElement", 0);
            oDocument.Call("RemoveElement", 1);

            // Save file and close DocBuilder
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();

            CDocBuilder.Destroy();
        }

        public static CValue createFullWidthTable(CValue oApi, int rows, int cols, int borderColor)
        {
            CValue oTable = oApi.Call("CreateTable", cols, rows);
            oTable.Call("SetWidth", "percent", 100);
            setTableBorders(oTable, borderColor);
            return oTable;
        }

        public static void setTableBorders(CValue oTable, int borderColor)
        {
            oTable.Call("SetTableBorderTop", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderBottom", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderLeft", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderRight", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderInsideV", "single", 4, 0, borderColor, borderColor, borderColor);
            oTable.Call("SetTableBorderInsideH", "single", 4, 0, borderColor, borderColor, borderColor);
        }

        public static CValue getTableCellParagraph(CValue oTable, int row, int col)
        {
            return oTable.Call("GetCell", row, col).Call("GetContent").Call("GetElement", 0);
        }

        public static void addTextToParagraph(CValue oParagraph, string text, int fontSize, bool isBold)
        {
            oParagraph.Call("AddText", text);
            oParagraph.Call("SetFontSize", fontSize);
            oParagraph.Call("SetBold", isBold);
        }

        public static void setPictureFormProperties(CValue oPictureForm, string key, string tip, bool required, string placeholder, string scaleFlag, bool lockAspectRatio, bool respectBorders, int shiftX, int shiftY)
        {
            oPictureForm.Call("SetFormKey", key);
            oPictureForm.Call("SetTipText", tip);
            oPictureForm.Call("SetRequired", required);
            oPictureForm.Call("SetPlaceholderText", placeholder);
            oPictureForm.Call("SetScaleFlag", scaleFlag);
            oPictureForm.Call("SetLockAspectRatio", lockAspectRatio);
            oPictureForm.Call("SetRespectBorders", respectBorders);
            oPictureForm.Call("SetPicturePosition", shiftX, shiftY);
        }

        public static void setTextFormProperties(CValue oTextForm, string key, string tip, bool required, string placeholder, bool comb, int maxCharacters, int cellWidth, bool multiLine, bool autoFit)
        {
            oTextForm.Call("SetFormKey", key);
            oTextForm.Call("SetTipText", tip);
            oTextForm.Call("SetRequired", required);
            oTextForm.Call("SetPlaceholderText", placeholder);
            oTextForm.Call("SetComb", comb);
            oTextForm.Call("SetCharactersLimit", maxCharacters);
            oTextForm.Call("SetCellWidth", cellWidth);
            oTextForm.Call("SetCellWidth", multiLine);
            oTextForm.Call("SetMultiline", autoFit);
        }

        public static void addTextFormToParagraph(CValue oParagraph, CValue oTextForm, int fontSize, string jc, bool hasBorder, int borderColor)
        {
            if (hasBorder)
            {
                oTextForm.Call("SetBorderColor", borderColor, borderColor, borderColor);
            }
            oParagraph.Call("AddElement", oTextForm);
            oParagraph.Call("SetFontSize", fontSize);
            oParagraph.Call("SetJc", jc);
        }
    }
}
