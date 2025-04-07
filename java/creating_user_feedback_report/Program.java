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

import java.io.FileReader;
import java.util.LinkedHashMap;
import java.util.ArrayList;
import java.util.Locale;
import java.util.Map;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class Program {
    static private CDocBuilderValue colorBlack;
    static private CDocBuilderValue colorOrange;
    static private CDocBuilderValue colorGrey;
    static private CDocBuilderValue colorBlue;

    public static void main(String[] args) throws Exception {
        String resultPath = "result.xlsx";
        String resourcesDir = "../../resources";

        createUserFeedbackReport(resultPath, resourcesDir);

        // Need to explicitly call System.gc() to free up resources
        System.gc();
    }

    public static void createUserFeedbackReport(String resultPath, String resourcesDir) throws Exception {
        // Parse JSON
        String jsonPath = resourcesDir + "/data/user_feedback_data.json";
        JSONArray data = (JSONArray) new JSONParser().parse(new FileReader(jsonPath));
        // Set default locale to US for correct table value display and chart plotting
        Locale.setDefault(Locale.US);

        // Init docbuilder and create xlsx file
        int doctype = FileTypes.Spreadsheet.XLSX;
        CDocBuilder.initialize("");
        CDocBuilder builder = new CDocBuilder();
        builder.createFile(doctype);

        CDocBuilderContext context = builder.getContext();
        CDocBuilderValue global = context.getGlobal();
        CDocBuilderValue api = global.get("Api");

        // Set main colors
        colorBlack = api.call("CreateColorFromRGB", 0, 0, 0);
        colorOrange = api.call("CreateColorFromRGB", 237, 125, 49);
        colorGrey = api.call("CreateRGBColor", 128, 128, 128);
        colorBlue = api.call("CreateRGBColor", 91, 155, 213);

        // Get current worksheet
        CDocBuilderValue worksheet1 = api.call("GetActiveSheet");

        // Create worksheet with average values
        worksheet1.call("SetName", "Average");
        int table1RowsCount = fillAverageSheet(worksheet1, data);

        // Create worksheet with comments and personal ratings
        api.call("AddSheet", "Comments");
        CDocBuilderValue worksheet2 = api.call("GetActiveSheet");
        int table2RowsCount = fillPersonalRatingsAndComments(worksheet2, data);

        // Create worksheet with charts
        api.call("AddSheet", "Charts");
        CDocBuilderValue worksheet3 = api.call("GetActiveSheet");
        createColumnChart(worksheet3, "Average!$A$2:$B$" + table1RowsCount, "Average ratings");
        createLineChart(api, worksheet3, data, "Dynamics of the average ratings");
        createPieChart(api, worksheet3, "Comments!$D$1:$D$" + table2RowsCount, "Shares of reviews");

        // Set first worksheet active
        worksheet1.call("SetActive");

        // Save and close
        builder.saveFile(doctype, resultPath);
        builder.closeFile();

        CDocBuilder.dispose();
    }

    private static int getSum(ArrayList<Integer> values){
        int sum = 0;
        for (int value : values) {
            sum += value;
        }
        return sum;
    }

    private static void setTableStyle(CDocBuilderValue range) {
        range.call("SetRowHeight", 24);
        range.call("SetAlignVertical", "center");

        String lineStyle = "Thin";
        range.call("SetBorders", "Top", lineStyle, colorBlack);
        range.call("SetBorders", "Left", lineStyle, colorBlack);
        range.call("SetBorders", "Right", lineStyle, colorBlack);
        range.call("SetBorders", "Bottom", lineStyle, colorBlack);
        range.call("SetBorders", "InsideHorizontal", lineStyle, colorBlack);
        range.call("SetBorders", "InsideVertical", lineStyle, colorBlack);
    }

    private static int fillAverageSheet(CDocBuilderValue worksheet, JSONArray feedbackData) {
        // Count detailed statistics for each question
        Map<String, ArrayList<Integer>> result = new LinkedHashMap<>();
        for (int i = 0; i < feedbackData.size(); i++) {
            JSONObject userFeedback = (JSONObject) feedbackData.get(i);
            JSONArray feedback = (JSONArray) userFeedback.get("feedback");
            for (int j = 0; j < feedback.size(); j++) {
                JSONObject feedbackItem = (JSONObject) feedback.get(j);
                String question = feedbackItem.get("question").toString();
                int rating = ((Long) ((JSONObject) feedbackItem.get("answer")).get("rating")).intValue();

                result.putIfAbsent(question, new ArrayList<>());
                result.get(question).add(rating);
            }
        }

        // Convert results
        String[] tableHeaders = {"Question", "Average Rating", "Number of Responses"};
        String[][] averageValues = new String[result.size() + 1][tableHeaders.length];
        averageValues[0] = tableHeaders;
        int index = 1;
        for (Map.Entry<String, ArrayList<Integer>> entry : result.entrySet()) {
            ArrayList<Integer> values = entry.getValue();
            int sum = getSum(values);
            averageValues[index] = new String[]{
                (String) entry.getKey(),
                String.format("%.1f", (double) sum / values.size()),
                String.valueOf(values.size())
            };
            index++;
        }

        int colsCount = tableHeaders.length - 1;
        CDocBuilderValue startCell = worksheet.call("GetRangeByNumber", 0, 0);
        CDocBuilderValue endCell = worksheet.call("GetRangeByNumber", averageValues.length - 1, colsCount);

        CDocBuilderValue averageRange = worksheet.call("GetRange", startCell, endCell);
        setTableStyle(averageRange);
        worksheet.call("GetRange", worksheet.call("GetRangeByNumber", 1, 1), endCell).call("SetAlignHorizontal", "center");

        CDocBuilderValue headerRow = worksheet.call("GetRange", startCell, worksheet.call("GetRangeByNumber", 0, colsCount));
        headerRow.call("SetBold", true);

        averageRange.call("SetValue", averageValues);
        averageRange.call("AutoFit", false, true);

        return averageValues.length;
    }

    private static int fillPersonalRatingsAndComments(CDocBuilderValue worksheet, JSONArray feedbackData) {
        String[][] tableHeaders = {{"Date", "Question", "Comment", "Rating", "Average User Rating"}};
        int colsCount = tableHeaders[0].length - 1;
        CDocBuilderValue startCell = worksheet.call("GetRangeByNumber", 0, 0);
        CDocBuilderValue headerRow = worksheet.call("GetRange", startCell, worksheet.call("GetRangeByNumber", 0, colsCount));
        headerRow.call("SetValue", tableHeaders);
        headerRow.call("SetBold", true);

        int rowsCount = 1;
        for (int i = 0; i < feedbackData.size(); i++) {
            JSONObject record = (JSONObject) feedbackData.get(i);
            JSONArray feedback = (JSONArray) record.get("feedback");

            int[] ratings = new int[2];
            String[][] userFeedback = new String[feedback.size()][tableHeaders.length];
            for (int j = 0; j < feedback.size(); j++) {
                JSONObject item = (JSONObject) feedback.get(j);
                JSONObject answer = (JSONObject) item.get("answer");
                userFeedback[j] = new String[]{
                    item.get("question").toString(),
                    answer.get("comment").toString(),
                    answer.get("rating").toString()
                };
                ratings[0] += ((Long) ((JSONObject) answer).get("rating")).intValue();
                ratings[1]++;
            }

            int userRowsCount = userFeedback.length - 1;
            double totalRating = (double) ratings[0] / ratings[1];

            // Fill date
            CDocBuilderValue dateCell = worksheet.call(
                "GetRange",
                worksheet.call("GetRangeByNumber", rowsCount, 0),
                worksheet.call("GetRangeByNumber", rowsCount + userRowsCount, 0)
            );
            dateCell.call("Merge", false);
            dateCell.call("SetValue", record.get("date").toString());

            // Fill ratings
            CDocBuilderValue userRange = worksheet.call(
                "GetRange",
                worksheet.call("GetRangeByNumber", rowsCount, 1),
                worksheet.call("GetRangeByNumber", rowsCount + userRowsCount, colsCount - 1)
            );
            userRange.call("SetValue", userFeedback);

            // Count average rating
            CDocBuilderValue ratingCell = worksheet.call(
                "GetRange",
                worksheet.call("GetRangeByNumber", rowsCount, colsCount),
                worksheet.call("GetRangeByNumber", rowsCount + userRowsCount, colsCount)
            );
            ratingCell.call("Merge", false);
            ratingCell.call("SetValue", String.format("%.1f", totalRating));

            // If rating <= 2, highlight it
            if (totalRating <= 2) {
                worksheet.call(
                    "GetRange",
                    worksheet.call("GetRangeByNumber", rowsCount, 0),
                    worksheet.call("GetRangeByNumber", rowsCount + userRowsCount, colsCount)
                ).call("SetFillColor", colorOrange);
            }

            // Update rows count
            rowsCount += userFeedback.length;
        }

        // Format table
        rowsCount -= 1;
        CDocBuilderValue resultRange = worksheet.call(
            "GetRange",
            startCell,
            worksheet.call("GetRangeByNumber", rowsCount, colsCount)
        );
        setTableStyle(resultRange);
        worksheet.call(
            "GetRange",
            worksheet.call("GetRangeByNumber", 1, colsCount - 1),
            worksheet.call("GetRangeByNumber", rowsCount, colsCount)
        ).call("SetAlignHorizontal", "center");
        resultRange.call("AutoFit", false, true);

        return rowsCount + 1;
    }

    private static void createColumnChart(CDocBuilderValue worksheet, String dataRange, String title) {
        CDocBuilderValue chart = worksheet.call("AddChart", dataRange, false, "bar", 2, 135.38 * 36000, 81.28 * 36000);
        chart.call("SetPosition", 0, 0, 0, 0);
        chart.call("SetTitle", title, 16);
    }

    private static void createLineChart(CDocBuilderValue api, CDocBuilderValue worksheet, JSONArray feedbackData, String title) {
        // Count average statistics for each date
        Map<String, ArrayList<Integer>> result = new LinkedHashMap<>();
        for (int i = 0; i < feedbackData.size(); i++) {
            JSONObject item = (JSONObject) feedbackData.get(i);
            JSONArray feedback = (JSONArray) item.get("feedback");
            String date = item.get("date").toString();

            for (int j = 0; j < feedback.size(); j++) {
                JSONObject feedbackItem = (JSONObject) feedback.get(j);
                int rating = ((Long) ((JSONObject) feedbackItem.get("answer")).get("rating")).intValue();

                result.putIfAbsent(date, new ArrayList<>());
                result.get(date).add(rating);
            }
        }

        // Convert results
        String[] headers = {"Date", "Rating"};
        String[][] avgValues = new String[result.size() + 1][headers.length];
        avgValues[0] = headers;
        int index = 1;
        for (Map.Entry<String, ArrayList<Integer>> entry : result.entrySet()) {
            ArrayList<Integer> values = entry.getValue();
            int sum = getSum(values);
            avgValues[index] = new String[]{
                (String) entry.getKey(),
                String.format("%.1f", (double) sum / values.size())
            };
            index++;
        }

        String dataRange = "$E$1:$F$" + avgValues.length;
        worksheet.call("GetRange", dataRange).call("SetValue", avgValues);

        CDocBuilderValue chart = worksheet.call("AddChart", "Charts!" + dataRange, false, "scatter", 2, 135.38 * 36000, 81.28 * 36000);
        chart.call("SetPosition", 0, 0, 18, 0);

        chart.call("SetSeriesFill", colorBlue, 0, false);
        CDocBuilderValue stroke = api.call(
            "CreateStroke",
            0.5 * 36000,
            api.call("CreateSolidFill", colorGrey)
        );
        chart.call("SetSeriesOutLine", stroke, 0, false);
        chart.call("SetTitle", title, 16);
        chart.call("SetMajorHorizontalGridlines", api.call("CreateStroke", 0, api.call("CreateNoFill")));
    }

    private static void createPieChart(CDocBuilderValue api, CDocBuilderValue worksheet, String dataRange, String title) {
        String[][] chartData = {
            {"Negative", "Neutral", "Positive"},
            {
                "=COUNTIF(" + dataRange + ", \"<=2\")",
                "=COUNTIF(" + dataRange + ", \"=3\")",
                "=COUNTIF(" + dataRange + ", \">=4\")"
            }
        };
        worksheet.call("GetRange", "$A$1:$C$2").call("SetValue", chartData);

        CDocBuilderValue chart = worksheet.call("AddChart", "Charts!$A$1:$C$2", true, "pie", 2, 135.38 * 36000, 81.28 * 36000);
        chart.call("SetPosition", 9, 0, 0, 0);
        chart.call("SetTitle", title, 16);

        chart.call("SetDataPointFill", api.call("CreateSolidFill", api.call("CreateRGBColor", 237, 125, 49)), 0, 0);
        chart.call("SetDataPointFill", api.call("CreateSolidFill", colorGrey), 0, 1);
        chart.call("SetDataPointFill", api.call("CreateSolidFill", colorBlue), 0, 2);

        CDocBuilderValue stroke = api.call(
            "CreateStroke",
            0.5 * 36000,
            api.call("CreateSolidFill", api.call("CreateRGBColor", 255, 255, 255))
        );
        chart.call("SetSeriesOutLine", stroke, 0, false);
    }
}
