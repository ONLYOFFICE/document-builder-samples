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
    public class CreatingBasicForm
    {
        public static void Main(string[] args)
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.docx";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateBasicForm(workDirectory, resultPath);
        }

        public static void CreateBasicForm(string workDirectory, string resultPath)
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

            // Save file and close DocBuilder
            oBuilder.SaveFile(doctype, resultPath);
            oBuilder.CloseFile();

            CDocBuilder.Destroy();
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
    }
}
