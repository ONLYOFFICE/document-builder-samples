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

namespace Sample
{
    public class CreatingPresentation
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.pptx";

            IDictionary<string, string> slideImages = new Dictionary<string, string>() {
                { "gun", "https://static.onlyoffice.com/assets/docs/samples/img/presentation_gun.png" },
                { "axe","https://static.onlyoffice.com/assets/docs/samples/img/presentation_axe.png" },
                { "knight", "https://static.onlyoffice.com/assets/docs/samples/img/presentation_knight.png" },
                { "sky","https://static.onlyoffice.com/assets/docs/samples/img/presentation_sky.png" }
            };
        // add Docbuilder dlls in path
        System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreatePresentation(workDirectory, resultPath, slideImages);
        }

        public static void CreatePresentation(string workDirectory, string resultPath, IDictionary<string, string> slideImages)
        {
            var doctype = (int)OfficeFileTypes.Presentation.PPTX;

            // Init DocBuilder
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();

            // Create presentation
            builder.CreateFile(doctype);
            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue presentation = api.Call("GetPresentation");
            presentation.Call("SetSizes", 9144000, 6858000);

            CValue slide = CreateImageSlide(api, presentation, slideImages["gun"]);
            presentation.Call("GetSlideByIndex", 0).Call("Delete");

            CValue shape = api.Call("CreateShape", "rect", 8056800, 3020400, api.Call("CreateNoFill"), api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
            shape.Call("SetPosition", 608400, 1267200);
            CValue content = shape.Call("GetDocContent");
            content.Call("RemoveAllElements");
            AddTextToSlideShape(api, content, "How They", 160, true, "left");
            AddTextToSlideShape(api, content, "Throw Out", 132, false, "left");
            AddTextToSlideShape(api, content, "a Challenge", 132, false, "left");
            slide.Call("AddObject", shape);

            slide = CreateImageSlide(api, presentation, slideImages["axe"]);

            shape = api.Call("CreateShape", "rect", 6904800, 1724400, api.Call("CreateNoFill"), api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
            shape.Call("SetPosition", 1764000, 1191600);
            content = shape.Call("GetDocContent");
            content.Call("RemoveAllElements");
            AddTextToSlideShape(api, content, "American Indians ", 110, true, "right");
            AddTextToSlideShape(api, content, "(XVII century)", 94, false, "right");
            slide.Call("AddObject", shape);

            shape = api.Call("CreateShape", "rect", 4986000, 2419200, api.Call("CreateNoFill"), api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
            shape.Call("SetPosition", 3834000, 3888000);
            content = shape.Call("GetDocContent");
            content.Call("RemoveAllElements");
            AddTextToSlideShape(api, content, "put a tomahawk on the ground in the ", 84, false, "right");
            AddTextToSlideShape(api, content, "rival's camp", 84, false, "right");
            slide.Call("AddObject", shape);

            slide = CreateImageSlide(api, presentation, slideImages["knight"]);

            shape = api.Call("CreateShape", "rect", 6904800, 1724400, api.Call("CreateNoFill"), api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
            shape.Call("SetPosition", 1764000, 1191600);
            content = shape.Call("GetDocContent");
            content.Call("RemoveAllElements");
            AddTextToSlideShape(api, content, "European Knights", 110, true, "right");
            AddTextToSlideShape(api, content, " (XII-XVI centuries)", 94, false, "right");
            slide.Call("AddObject", shape);

            shape = api.Call("CreateShape", "rect", 4986000, 2419200, api.Call("CreateNoFill"), api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
            shape.Call("SetPosition", 3834000, 3888000);
            content = shape.Call("GetDocContent");
            content.Call("RemoveAllElements");
            AddTextToSlideShape(api, content, "threw a glove", 84, false, "right");
            AddTextToSlideShape(api, content, "in the rival's face", 84, false, "right");
            slide.Call("AddObject", shape);

            slide = CreateImageSlide(api, presentation, slideImages["sky"]);

            shape = api.Call("CreateShape", "rect", 7887600, 3063600, api.Call("CreateNoFill"), api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
            shape.Call("SetPosition", 630000, 1357200);
            content = shape.Call("GetDocContent");
            content.Call("RemoveAllElements");
            AddTextToSlideShape(api, content, "OnlyOffice", 176, false, "center");
            AddTextToSlideShape(api, content, "stands for Peace", 132, false, "center");
            slide.Call("AddObject", shape);

            // Save file and close DocBuilder
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static CValue CreateImageSlide(CValue api, CValue presentation, string image_url)
        {
            CValue slide = api.Call("CreateSlide");
            presentation.Call("AddSlide", slide);
            var fill = api.Call("CreateBlipFill", image_url, "stretch");
            slide.Call("SetBackground", fill);
            slide.Call("RemoveAllObjects");
            return slide;
        }

        public static void AddTextToSlideShape(CValue api, CValue content, string text, int fontSize, bool isBold, string js)
        {
            var paragraph = api.Call("CreateParagraph");
            paragraph.Call("SetSpacingBefore", 0);
            paragraph.Call("SetSpacingAfter", 0);
            content.Call("Push", paragraph);
            var run = paragraph.Call("AddText", text);
            run.Call("SetFill", api.Call("CreateSolidFill", api.Call("CreateRGBColor", 0xff, 0xff, 0xff)));
            run.Call("SetFontSize", fontSize);
            run.Call("SetFontFamily", "Georgia");
            run.Call("SetBold", isBold);
            paragraph.Call("SetJc", js);
        }
    }
}
