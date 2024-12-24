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

namespace Sample
{
    public class CreatingBasicForm
    {
        public static void Main()
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
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];

            // Create basic form
            CValue document = api.Call("GetDocument");
            CValue paragraph = document.Call("GetElement", 0);
            CValue headingStyle = document.Call("GetStyle", "Heading 3");

            paragraph.Call("AddText", "Employee pass card");
            paragraph.Call("SetStyle", headingStyle);
            document.Call("Push", paragraph);

            CValue pictureForm = api.Call("CreatePictureForm");
            SetPictureFormProperties(pictureForm, "Photo", "Upload your photo", false, "Photo", "tooBig", true, false, 50, 50);
            paragraph = api.Call("CreateParagraph");
            paragraph.Call("AddElement", pictureForm);
            document.Call("Push", paragraph);

            CValue textForm = api.Call("CreateTextForm");
            SetTextFormProperties(textForm, "First name", "Enter your first name", false, "First name", true, 13, 3, false, false);
            paragraph = api.Call("CreateParagraph");
            paragraph.Call("AddElement", textForm);
            document.Call("Push", paragraph);

            // Save file and close DocBuilder
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();

            CDocBuilder.Destroy();
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
    }
}
