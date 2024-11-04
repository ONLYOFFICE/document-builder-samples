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

public class Program {
    public static void main(String[] args) {
        String resultPath = "result.docx";

        createBasicForm(resultPath);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createBasicForm(String resultPath) {
        int doctype = FileTypes.Document.DOCX;

        // Initialize builder with docbuilder.jar directory by passing empty string
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");

        // Create basic form
        CDocBuilderValue document = api.call("GetDocument");
        CDocBuilderValue paragraph = document.call("GetElement", 0);
        CDocBuilderValue headingStyle = document.call("GetStyle", "Heading 3");

        paragraph.call("AddText", "Employee pass card");
        paragraph.call("SetStyle", headingStyle);
        document.call("Push", paragraph);

        CDocBuilderValue pictureForm = api.call("CreatePictureForm");
        setPictureFormProperties(pictureForm, "Photo", "Upload your photo", false, "Photo", "tooBig", true, false, 50, 50);
        paragraph = api.call("CreateParagraph");
        paragraph.call("AddElement", pictureForm);
        document.call("Push", paragraph);

        CDocBuilderValue textForm = api.call("CreateTextForm");
        setTextFormProperties(textForm, "First name", "Enter your first name", false, "First name", true, 13, 3, false, false);
        paragraph = api.call("CreateParagraph");
        paragraph.call("AddElement", textForm);
        document.call("Push", paragraph);

        // Save file and close DocBuilder
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }

    public static void setPictureFormProperties(CDocBuilderValue pictureForm, String key, String tip, boolean required, String placeholder, String scaleFlag, boolean lockAspectRatio, boolean respectBorders, int shiftX, int shiftY) {
        pictureForm.call("SetFormKey", key);
        pictureForm.call("SetTipText", tip);
        pictureForm.call("SetRequired", required);
        pictureForm.call("SetPlaceholderText", placeholder);
        pictureForm.call("SetScaleFlag", scaleFlag);
        pictureForm.call("SetLockAspectRatio", lockAspectRatio);
        pictureForm.call("SetRespectBorders", respectBorders);
        pictureForm.call("SetPicturePosition", shiftX, shiftY);
    }

    public static void setTextFormProperties(CDocBuilderValue textForm, String key, String tip, boolean required, String placeholder, boolean comb, int maxCharacters, int cellWidth, boolean multiLine, boolean autoFit) {
        textForm.call("SetFormKey", key);
        textForm.call("SetTipText", tip);
        textForm.call("SetRequired", required);
        textForm.call("SetPlaceholderText", placeholder);
        textForm.call("SetComb", comb);
        textForm.call("SetCharactersLimit", maxCharacters);
        textForm.call("SetCellWidth", cellWidth);
        textForm.call("SetCellWidth", multiLine);
        textForm.call("SetMultiline", autoFit);
    }
}
