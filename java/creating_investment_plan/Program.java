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

import java.io.FileReader;

import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class Program {
    public static void main(String[] args) throws Exception {
        String resultPath = "result.xlsx";
        String resourcesDir = "../../resources";

        createInvestmentPlan(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createInvestmentPlan(String resultPath, String resourcesDir) throws Exception {
        // parse JSON
        String jsonPath = resourcesDir + "/data/investment_data.json";
        JSONObject data = (JSONObject)new JSONParser().parse(new FileReader(jsonPath));

        // init docbuilder and create new xlsx file
        int doctype = FileTypes.Spreadsheet.XLSX;
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");
        CDocBuilderValue worksheet = api.call("GetActiveSheet");

        // initialize financial data from JSON
        int initAmount = (int)(long)data.get("initial_amount");
        double rate = (double)data.get("return_rate");
        int term = (int)(long)data.get("term");

        // fill years
        CDocBuilderValue startCell = worksheet.call("GetRangeByNumber", 1, 0);
        CDocBuilderValue endCell = worksheet.call("GetRangeByNumber", term + 1, 0);
        String[] years = new String[term + 1];
        for (int year = 0; year <= term; year++) {
            years[year] = Integer.toString(year);
        }
        worksheet.call("GetRange", startCell, endCell).call("SetValue", CreateColumnData(years));

        // fill initial amount
        worksheet.call("GetRangeByNumber", 1, 1).call("SetValue", initAmount);
        // fill remaining cells
        startCell = worksheet.call("GetRangeByNumber", 2, 1);
        endCell = worksheet.call("GetRangeByNumber", term + 1, 1);
        String[] amounts = new String[term];
        for (int year = 0; year < term; year++) {
            amounts[year] = String.format("=$B$2*POWER((1+%s),A%d)", Double.toString(rate).replace(',', '.'), year + 3);
        }
        worksheet.call("GetRange", startCell, endCell).call("SetValue", CreateColumnData(amounts));

        // create chart
        CDocBuilderValue chart = worksheet.call("AddChart", String.format("Sheet1!$A$1:$B$%d", term + 2), false, "lineNormal", 2, 135.38 * 36000, 81.28 * 36000);
        chart.call("SetPosition", 3, 0, 2, 0);
        chart.call("SetTitle", "Capital Growth Over Time", 22);
        CDocBuilderValue color = api.call("CreateRGBColor", 134, 134, 134);
        CDocBuilderValue fill = api.call("CreateSolidFill", color);
        CDocBuilderValue stroke = api.call("CreateStroke", 1, fill);
        chart.call("SetMinorVerticalGridlines", stroke);
        chart.call("SetMajorHorizontalGridlines", stroke);
        // fill table headers
        worksheet.call("GetRangeByNumber", 0, 0).call("SetValue", "Year");
        worksheet.call("GetRangeByNumber", 0, 1).call("SetValue", "Amount");

        // save and close
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }

    public static CDocBuilderValue CreateColumnData(String[] data) {
        CDocBuilderValue columnData = CDocBuilderValue.createArray(data.length);
        for (int i = 0; i < data.length; i++) {
            CDocBuilderValue row = CDocBuilderValue.createArray(1);
            row.set(0, data[i]);
            columnData.set(i, row);
        }
        return columnData;
    }
}
