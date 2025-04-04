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
using System.Linq;

namespace Sample
{
    public class FeedbackReport
    {
        private static CValue color_black;
        private static CValue color_orange;
        private static CValue color_grey;
        private static CValue color_blue;

        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.xlsx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateFeedbackReport(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateFeedbackReport(string workDirectory, string resultPath, string resourcesDir)
        {
            // parse JSON
            string jsonPath = resourcesDir + "/data/user_feedback_data.json";
            string json = File.ReadAllText(jsonPath);
            JsonData data = JsonSerializer.Deserialize<JsonData>(json);
            // Set default locale to US for correct table value display and chart plotting
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            // Init DocBuilder
            var doctype = (int)OfficeFileTypes.Spreadsheet.XLSX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];

            // Set main colors
            color_black = api.Call("CreateColorFromRGB", 0, 0, 0);
            color_orange = api.Call("CreateColorFromRGB", 237, 125, 49);
            color_grey = api.Call("CreateRGBColor", 128, 128, 128);
            color_blue = api.Call("CreateRGBColor", 91, 155, 213);

            // Get current worksheet
            CValue worksheet1 = api.Call("GetActiveSheet");

            // Create worksheet with average values
            worksheet1.Call("SetName", "Average");
            int table1RowsCount = FillAverageSheet(worksheet1, data);

            // Create worksheet with comments and personal ratings
            api.Call("AddSheet", "Comments");
            CValue worksheet2 = api.Call("GetActiveSheet");
            int table2RowsCount = FillPersonalRatingsAndComments(worksheet2, data);

            // Create worksheet with charts
            api.Call("AddSheet", "Charts");
            CValue worksheet3 = api.Call("GetActiveSheet");

            CreateColumnChart(worksheet3, $"Average!$A$2:$B${table1RowsCount}", "Average ratings");
            CreateLineChart(api, worksheet3, data, "Dynamics of the average ratings");
            CreatePieChart(api, worksheet3, $"Comments!$D$1:$D${table2RowsCount}", "Shares of reviews");

            // Set first worksheet active
            worksheet1.Call("SetActive");

            // Save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        // Helper functions
        public static void SetTableStyle(CValue range)
        {
            range.Call("SetRowHeight", 24);
            range.Call("SetAlignVertical", "center");

            string lineStyle = "Thin";
            range.Call("SetBorders", "Top", lineStyle, color_black);
            range.Call("SetBorders", "Left", lineStyle, color_black);
            range.Call("SetBorders", "Right", lineStyle, color_black);
            range.Call("SetBorders", "Bottom", lineStyle, color_black);
            range.Call("SetBorders", "InsideHorizontal", lineStyle, color_black);
            range.Call("SetBorders", "InsideVertical", lineStyle, color_black);
        }

        public static int FillAverageSheet(CValue worksheet, JsonData feedbackData)
        {
            Dictionary<string, List<int>> result = new();
            List<string> questionOrder = new();

            foreach (UserFeedback record in feedbackData) {
                foreach (FeedbackItem item in record.feedback) {
                    if (!result.ContainsKey(item.question)) {
                        questionOrder.Add(item.question);
                        result[item.question] = new List<int> { item.answer.rating };
                    }
                    else {
                        result[item.question].Add(item.answer.rating);
                    }
                }
            }

            CValue[] averageValues = new CValue[questionOrder.Count + 1];
            averageValues[0] = new CValue[] { "Question", "Average Rating", "Number of Responses" };
            for (int i = 0; i < questionOrder.Count; i++) {
                List<int> ratings = result[questionOrder[i]];
                double average = (double) ratings.Sum() / ratings.Count;
                averageValues[i + 1] = new CValue[] { questionOrder[i], $"{average:F2}", ratings.Count.ToString() };
            }

            int colsCount = (int)averageValues[0].GetLength() - 1;
            int rowsCount = (int)averageValues.Length;
            CValue startСell = worksheet.Call("GetRangeByNumber", 0, 0);
            CValue endCell = worksheet.Call("GetRangeByNumber", rowsCount - 1, colsCount);

            CValue averageRange = worksheet.Call("GetRange", startСell, endCell);
            SetTableStyle(averageRange);
            worksheet.Call(
                "GetRange",
                worksheet.Call("GetRangeByNumber", 1, 1),
                endCell
            ).Call("SetAlignHorizontal", "center");

            CValue headerRow = worksheet.Call(
                "GetRange",
                startСell,
                worksheet.Call("GetRangeByNumber", 0, colsCount)
            );
            headerRow.Call("SetBold", true);

            averageRange.Call("SetValue", averageValues);
            averageRange.Call("AutoFit", false, true);

            return rowsCount;
        }

        public static int FillPersonalRatingsAndComments(CValue worksheet, JsonData feedbackData)
        {
            CValue[] headerValues = new CValue[1];
            headerValues[0] = new CValue[] { "Date", "Question", "Comment", "Rating", "Average User Rating" };
            int colsCount = (int)headerValues[0].GetLength() - 1;
            CValue startСell = worksheet.Call("GetRangeByNumber", 0, 0);
            CValue headerRow = worksheet.Call(
                "GetRange",
                startСell,
                worksheet.Call("GetRangeByNumber", 0, colsCount)
            );

            headerRow.Call("SetValue", headerValues);
            headerRow.Call("SetBold", true);

            int rowsCount = 1;
            foreach (UserFeedback record in feedbackData) {
                // Count and fill user feedback
                double avgRating = 0;

                int feedbackSize = record.feedback.Count;
                CValue[] userFeedback = new CValue[feedbackSize];
                int i = 0;
                foreach (FeedbackItem item in record.feedback) {
                    userFeedback[i] = new CValue[] { item.question, item.answer.comment, item.answer.rating.ToString() };
                    avgRating += item.answer.rating;
                    i++;
                }

                int userRowsCount = feedbackSize - 1;
                // Fill date
                CValue dateCell = worksheet.Call(
                    "GetRange",
                    worksheet.Call("GetRangeByNumber", rowsCount, 0),
                    worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, 0)
                );
                dateCell.Call("Merge", false);
                dateCell.Call("SetValue", record.date);

                // Fill ratings
                CValue userRange = worksheet.Call(
                    "GetRange",
                    worksheet.Call("GetRangeByNumber", rowsCount, 1),
                    worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, colsCount - 1)
                );
                userRange.Call("SetValue", userFeedback);

                // Count average rating
                avgRating = avgRating / feedbackSize;
                CValue ratingCell = worksheet
                    .Call(
                        "GetRange",
                        worksheet.Call("GetRangeByNumber", rowsCount, colsCount),
                        worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, colsCount)
                    );
                ratingCell.Call("Merge", false);
                ratingCell.Call("SetValue", $"{avgRating:F2}");

                // If rating <= 2, highlight it
                if (avgRating <= 2) {
                    worksheet.Call(
                        "GetRange",
                        worksheet.Call("GetRangeByNumber", rowsCount, 0),
                        worksheet.Call("GetRangeByNumber", rowsCount + userRowsCount, colsCount)
                    ).Call("SetFillColor", color_orange);
                }

                // Update rows count
                rowsCount += feedbackSize;
            }

            // Format table
            rowsCount -= 1;
            CValue resultRange = worksheet.Call(
                "GetRange",
                startСell,
                worksheet.Call("GetRangeByNumber", rowsCount, colsCount)
            );
            SetTableStyle(resultRange);
            worksheet.Call(
                "GetRange",
                worksheet.Call("GetRangeByNumber", 1, colsCount - 1),
                worksheet.Call("GetRangeByNumber", rowsCount, colsCount)
            ).Call("SetAlignHorizontal", "center");
            resultRange.Call("AutoFit", false, true);

            return rowsCount + 1;
        }

        public static void CreateColumnChart(CValue worksheet, string dataRange, string title)
        {
            CValue chart = worksheet.Call("AddChart", dataRange, false, "bar", 2, 135.38 * 36000, 81.28 * 36000);
            chart.Call("SetPosition", 0, 0, 0, 0);
            chart.Call("SetTitle", title, 16);
        }

        public static void CreateLineChart(CValue api, CValue worksheet, JsonData feedbackData, string title)
        {
            Dictionary<string, List<int>> result = new();
            List<string> dateOrder = new();

            foreach (UserFeedback record in feedbackData) {
                if (!result.ContainsKey(record.date)) {
                    dateOrder.Add(record.date);
                    result[record.date] = new List<int>();
                }
                foreach (FeedbackItem item in record.feedback) {
                    result[record.date].Add(item.answer.rating);
                }
            }

            CValue[] averageDayRating = new CValue[dateOrder.Count + 1];
            averageDayRating[0] = new CValue[] { "Date", "Rating" };
            for (int i = 0; i < dateOrder.Count; i++) {
                List<int> ratings = result[dateOrder[i]];
                double average = (double) ratings.Sum() / ratings.Count;
                averageDayRating[i + 1] = new CValue[] { dateOrder[i], $"{average:F2}" };
            }

            string dataRange = $"$E$1:$F${averageDayRating.Length}";
            worksheet.Call("GetRange", dataRange).Call("SetValue", averageDayRating);
            CValue chart = worksheet.Call("AddChart", $"Charts!{dataRange}", false, "scatter", 2, 135.38 * 36000, 81.28 * 36000);
            chart.Call("SetPosition", 0, 0, 18, 0);
            chart.Call("SetSeriesFill", color_blue, 0, false);

            CValue stroke = api.Call(
                "CreateStroke",
                0.5 * 36000,
                api.Call("CreateSolidFill", color_grey)
            );
            chart.Call("SetSeriesOutLine", stroke, 0, false);
            chart.Call("SetTitle", title, 16);
            chart.Call("SetMajorHorizontalGridlines", api.Call("CreateStroke", 0, api.Call("CreateNoFill")));
        }

        public static void CreatePieChart(CValue api, CValue worksheet, string dataRange, string title)
        {
            CValue[] pieChartData = new CValue[2] {
                new CValue[] { "Negative", "Neutral", "Positive" },
                new CValue[] {
                    $"=COUNTIF({dataRange}, \"<=2\")",
                    $"=COUNTIF({dataRange}, \"=3\")",
                    $"=COUNTIF({dataRange}, \">=4\")"
                }
            };
            worksheet.Call("GetRange", "$A$1:$C$2").Call("SetValue", pieChartData);

            CValue chart = worksheet.Call("AddChart", "Charts!$A$1:$C$2", true, "pie", 2, 135.38 * 36000, 81.28 * 36000);
            chart.Call("SetPosition", 9, 0, 0, 0);
            chart.Call("SetTitle", title, 16);
            chart.Call("SetDataPointFill", api.Call("CreateSolidFill", api.Call("CreateRGBColor", 237, 125, 49)), 0, 0);
            chart.Call("SetDataPointFill", api.Call("CreateSolidFill", color_grey), 0, 1);
            chart.Call("SetDataPointFill", api.Call("CreateSolidFill", color_blue), 0, 2);

            CValue stroke = api.Call(
                "CreateStroke",
                0.5 * 36000,
                api.Call("CreateSolidFill", api.Call("CreateRGBColor", 255, 255, 255))
            );
            chart.Call("SetSeriesOutLine", stroke, 0, false);
        }
    }

    // Define classes to represent the JSON structure
    public class JsonData : List<UserFeedback>
    {
    }

    public class UserFeedback
    {
        public string user_id { get; set; }
        public string date { get; set; }
        public List<FeedbackItem> feedback { get; set; }
    }

    public class FeedbackItem
    {
        public string question { get; set; }
        public Answer answer { get; set; }
    }

    public class Answer
    {
        public int rating { get; set; }
        public string comment { get; set; }
    }
}
