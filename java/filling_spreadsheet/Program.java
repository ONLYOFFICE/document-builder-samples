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

import docbuilder.*;

public class Program {
    public static void main(String[] args) {
        String resultPath = "result.xlsx";
        // Fill table data
        Object[][] data = {
            { "Id", "Product", "Price", "Available" },
            { 1001, "Item A", 12.2, true },
            { 1002, "Item B", 18.8, true },
            { 1003, "Item C", 70.1, false },
            { 1004, "Item D", 60.6, true },
            { 1005, "Item E", 32.6, true },
            { 1006, "Item F", 28.3, false },
            { 1007, "Item G", 11.1, false },
            { 1008, "Item H", 41.4, true }
        };

        fillSpreadsheet(resultPath, data);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void fillSpreadsheet(String resultPath, Object[][] data) {
        int doctype = FileTypes.Spreadsheet.XLSX;

        // Initialize builder with docbuilder.jar directory by passing empty String
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");

        // Get current worksheet
        CDocBuilderValue worksheet = api.call("GetActiveSheet");

        // First cell in the range (A1) is equal to (0,0)
        CDocBuilderValue startCell = worksheet.call("GetRangeByNumber", 0, 0);
        // Last cell in the range is equal to array length -1
        CDocBuilderValue endCell = worksheet.call("GetRangeByNumber", data.length - 1, data[0].length - 1);
        worksheet.call("GetRange", startCell, endCell).call("SetValue", data);

        // Save file and close DocBuilder
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }
}
