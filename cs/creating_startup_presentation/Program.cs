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

using System;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;

namespace Sample
{
    public class CreatingPresentation
    {
        public static void Main()
        {
            string workDirectory = Constants.BUILDER_DIR;
            string resultPath = "../../../result.pptx";
            string resourcesDir = "../../../../../../resources";

            // add Docbuilder dlls in path
            System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + workDirectory);

            CreateStartupPresentation(workDirectory, resultPath, resourcesDir);
        }

        public static void CreateStartupPresentation(string workDirectory, string resultPath, string resourcesDir)
        {
            // init docbuilder and create new pptx file
            var doctype = (int)OfficeFileTypes.Presentation.PPTX;
            CDocBuilder.Initialize(workDirectory);
            CDocBuilder builder = new();
            builder.CreateFile(doctype);

            CContext context = builder.GetContext();
            CValue global = context.GetGlobal();
            CValue api = global["Api"];
            CValue presentation = api.Call("GetPresentation");

            // init colors
            CValue backgroundFill = api.Call("CreateSolidFill", api.Call("CreateRGBColor", 255, 255, 255));
            CValue textFill = api.Call("CreateSolidFill", api.Call("CreateRGBColor", 80, 80, 80));
            CValue textSpecialFill = api.Call("CreateSolidFill", api.Call("CreateRGBColor", 15, 102, 7));
            CValue textAltFill = api.Call("CreateSolidFill", api.Call("CreateRGBColor", 230, 69, 69));
            CValue chartGridFill = api.Call("CreateSolidFill", api.Call("CreateRGBColor", 134, 134, 134));
            CValue master = presentation.Call("GetMaster", 0);
            CValue colorScheme = master.Call("GetTheme").Call("GetColorScheme");
            colorScheme.Call("ChangeColor", 0, api.Call("CreateRGBColor", 15, 102, 7));

            // TITLE slide
            CValue slide = presentation.Call("GetSlideByIndex", 0);
            slide.Call("SetBackground", backgroundFill);
            CValue paragraph = slide.Call("GetAllShapes")[0].Call("GetContent").Call("GetElement", 0);
            AddTextToParagraph(api, paragraph, "GreenVibe Solutions", 120, textSpecialFill, true, "center", "Arial Black");
            paragraph = slide.Call("GetAllShapes")[1].Call("GetContent").Call("GetElement", 0);
            AddTextToParagraph(api, paragraph, "12.12.2024", 48, textFill, false, "center");

            // MARKET OVERVIEW slide
            // parse JSON, obtained as Statista API response
            string jsonPath = resourcesDir + "/data/statista_api_response.json";
            string json = File.ReadAllText(jsonPath);
            MarketOverviewData dataMarket = JsonSerializer.Deserialize<MarketOverviewData>(json);
            // create new slide
            slide = AddNewSlide(api, backgroundFill);
            // title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            AddTextToParagraph(api, paragraph, "Market Overview", 72, textFill, false, "center");
            // market size
            Tuple<string, string> marketSize = SeparateValueAndUnit(dataMarket.market.size);
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.58);
            AddTextToParagraph(api, paragraph, "Market size:", 48, textFill, false, "center");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.97);
            AddTextToParagraph(api, paragraph, marketSize.Item1, 144, textSpecialFill, false, "center", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 3.06);
            AddTextToParagraph(api, paragraph, marketSize.Item2, 48, textFill, false, "center");
            // growth rate
            Tuple<string, string> marketGrowth = SeparateValueAndUnit(dataMarket.market.growth_rate);
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 7, 1.58);
            AddTextToParagraph(api, paragraph, "Growth rate:", 48, textFill, false, "center");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 7, 1.97);
            AddTextToParagraph(api, paragraph, marketGrowth.Item1, 144, textSpecialFill, false, "center", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 7, 3.06);
            AddTextToParagraph(api, paragraph, marketGrowth.Item2, 48, textFill, false, "center");
            // trends
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 3.75);
            AddTextToParagraph(api, paragraph, "Trends:", 48, textFill, false, "center");
            paragraph = AddParagraphToSlide(api, slide, 0.93, 2.92, 1.57, 4.31);
            AddTextToParagraph(api, paragraph, MakeBulletString('>', dataMarket.market.trends.Count), 72, textSpecialFill, false, "left", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 9.21, 2.92, 2.1, 4.31);
            string trendsText = "";
            foreach (string trend in dataMarket.market.trends)
            {
                trendsText += trend + "\n";
            }
            AddTextToParagraph(api, paragraph, trendsText, 72, textSpecialFill, false, "center", "Arial Black");

            // COMPETITORS OVERVIEW section
            // parse JSON, obtained as Statista API response
            jsonPath = resourcesDir + "/data/crunchbase_api_response.json";
            json = File.ReadAllText(jsonPath);
            CompetitorsOverviewData dataCompetitors = JsonSerializer.Deserialize<CompetitorsOverviewData>(json);
            // create new slide
            slide = AddNewSlide(api, backgroundFill);
            // title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            AddTextToParagraph(api, paragraph, "Competitors Overview", 72, textFill, false, "center");
            // chart header
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
            AddTextToParagraph(api, paragraph, "Market shares", 48, textFill, false, "center");
            // get chart data
            double othersShare = 100.0;
            int competitorsCount = dataCompetitors.competitors.Count;
            CValue[] shares = new CValue[competitorsCount + 1];
            CValue[] competitors = new CValue[competitorsCount + 1];
            for (int i = 0; i < competitorsCount; i++)
            {
                CompetitorData competitor = dataCompetitors.competitors[i];
                competitors[i] = competitor.name;
                string share = competitor.market_share;
                // remove last percent symbol
                share = share.Remove(share.Length - 1);
                othersShare -= double.Parse(share);
                shares[i] = share;
            }
            shares[competitorsCount] = othersShare.ToString().Replace(',', '.');
            competitors[competitorsCount] = "Others";
            // create a chart
            CValue chartData = context.CreateArray(1);
            chartData[0] = shares;
            CValue chart = api.Call("CreateChart", "pie", chartData, context.CreateArray(0), competitors);
            SetChartSizes(chart, 6.51, 5.9, 4.18, 1.49);
            chart.Call("SetLegendFontSize", 14);
            chart.Call("SetLegendPos", "right");
            slide.Call("AddObject", chart);

            // create slide for every competitor with brief info
            foreach (CompetitorData competitor in dataCompetitors.competitors)
            {
                // create new slide
                slide = AddNewSlide(api, backgroundFill);
                // title
                paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
                AddTextToParagraph(api, paragraph, "Competitors Overview", 72, textFill, false, "center");
                // header
                paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
                AddTextToParagraph(api, paragraph, competitor.name, 64, textFill, false, "center");
                // recent funding
                paragraph = AddParagraphToSlide(api, slide, 3.13, 0.8, 1.07, 2.65);
                AddTextToParagraph(api, paragraph, "Recent funding:", 48, textFill);
                paragraph = AddParagraphToSlide(api, slide, 8.9, 0.8, 4.19, 2.52);
                AddTextToParagraph(api, paragraph, competitor.recent_funding, 96, textSpecialFill, false, "left", "Arial Black");
                // main products
                paragraph = AddParagraphToSlide(api, slide, 3.13, 0.8, 1.07, 3.72);
                AddTextToParagraph(api, paragraph, "Main products:", 48, textFill);
                paragraph = AddParagraphToSlide(api, slide, 0.93, 3.53, 4.19, 3.72);
                AddTextToParagraph(api, paragraph, MakeBulletString('>', competitor.products.Count), 72, textSpecialFill, false, "left", "Arial Black");
                paragraph = AddParagraphToSlide(api, slide, 7.97, 3.53, 5.12, 3.72);
                string productsText = "";
                foreach (string product in competitor.products)
                {
                    productsText += product + '\n';
                }
                AddTextToParagraph(api, paragraph, productsText, 72, textSpecialFill, false, "left", "Arial Black");
            }

            // TARGET AUDIENCE section
            // parse JSON, obtained as Social Media Insights API response
            jsonPath = resourcesDir + "/data/smi_api_response.json";
            json = File.ReadAllText(jsonPath);
            AudienceOverviewData dataAudience = JsonSerializer.Deserialize<AudienceOverviewData>(json);
            // create new slide
            slide = AddNewSlide(api, backgroundFill);
            // title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            AddTextToParagraph(api, paragraph, "Target Audience", 72, textFill, false, "center");

            // demographics
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.33);
            AddTextToParagraph(api, paragraph, "Demographics:", 48, textFill, false, "center");
            // age range
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.97);
            AddTextToParagraph(api, paragraph, dataAudience.demographics.age_range, 128, textSpecialFill, false, "center", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 2.95);
            AddTextToParagraph(api, paragraph, "age range", 40, textFill, false, "center");
            // location
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 3.68);
            AddTextToParagraph(api, paragraph, dataAudience.demographics.location, 72, textSpecialFill, false, "center", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 4.27);
            AddTextToParagraph(api, paragraph, "location", 40, textFill, false, "center");
            // income level
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 5.28);
            AddTextToParagraph(api, paragraph, dataAudience.demographics.income_level, 56, textSpecialFill, false, "center", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 5.83);
            AddTextToParagraph(api, paragraph, "income level", 40, textFill, false, "center");

            // social trends
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 6, 1.33);
            AddTextToParagraph(api, paragraph, "Social trends:", 48, textFill, false, "center");
            // positive feedback
            paragraph = AddParagraphToSlide(api, slide, 0.63, 2.42, 7, 2.06);
            AddTextToParagraph(api, paragraph, MakeBulletString('+', dataAudience.social_trends.positive_feedback.Count), 52, textSpecialFill, false, "left", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.56, 2.42, 7.67, 2.06);
            string positiveFeedback = "";
            foreach (string feedback in dataAudience.social_trends.positive_feedback)
            {
                positiveFeedback += feedback + '\n';
            }
            AddTextToParagraph(api, paragraph, positiveFeedback, 52, textSpecialFill, false, "left", "Arial Black");
            // negative feedback
            paragraph = AddParagraphToSlide(api, slide, 0.63, 2.42, 7, 4.55);
            AddTextToParagraph(api, paragraph, MakeBulletString('-', dataAudience.social_trends.negative_feedback.Count), 52, textAltFill, false, "left", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.56, 2.42, 7.67, 4.55);
            string negativeFeedback = "";
            foreach (string feedback in dataAudience.social_trends.negative_feedback)
            {
                negativeFeedback += feedback + '\n';
            }
            AddTextToParagraph(api, paragraph, negativeFeedback, 52, textAltFill, false, "left", "Arial Black");

            // SEARCH TRENDS section
            // parse JSON, obtained as Google Trends API response
            jsonPath = resourcesDir + "/data/google_trends_api_response.json";
            json = File.ReadAllText(jsonPath);
            TrendsOverviewData dataTrends = JsonSerializer.Deserialize<TrendsOverviewData>(json);
            // create new slide
            slide = AddNewSlide(api, backgroundFill);
            // title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            AddTextToParagraph(api, paragraph, "Search Trends", 72, textFill, false, "center");
            // add every trend on the slide
            double offsetY = 1.43;
            foreach (SearchTrendData trend in dataTrends.search_trends)
            {
                paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, offsetY);
                AddTextToParagraph(api, paragraph, trend.topic, 96, textSpecialFill, false, "center", "Arial Black");
                paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, offsetY + 0.8);
                AddTextToParagraph(api, paragraph, trend.growth, 40, textFill, false, "center");
                offsetY += 1.25;
            }

            // FINANCIAL MODEL section
            // parse JSON, obtained from financial system
            jsonPath = resourcesDir + "/data/financial_model_data.json";
            json = File.ReadAllText(jsonPath);
            FinancialModelData dataFinancial = JsonSerializer.Deserialize<FinancialModelData>(json);
            // create new slide
            slide = AddNewSlide(api, backgroundFill);
            // title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            AddTextToParagraph(api, paragraph, "Financial Model", 72, textFill, false, "center");
            // chart title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
            AddTextToParagraph(api, paragraph, "Profit forecast", 48, textFill, false, "center");
            // chart
            string[] chartKeys = { "revenue", "cost_of_goods_sold", "gross_profit", "operating_expenses", "net_profit" };
            var profitForecast = dataFinancial.profit_forecast;
            CValue chartYears = context.CreateArray(profitForecast.Count);
            for (int i = 0; i < profitForecast.Count; i++)
            {
                chartYears[i] = profitForecast[i].year;
            }
            chartData = context.CreateArray(chartKeys.Length);
            for (int i = 0; i < chartKeys.Length; i++)
            {
                chartData[i] = context.CreateArray(profitForecast.Count);
                for (int j = 0; j < profitForecast.Count; j++)
                {
                    chartData[i][j] = (string)profitForecast[j].GetType().GetProperty(chartKeys[i]).GetValue(profitForecast[j]);
                }
            }
            CValue chartNames = new CValue[] { "Revenue", "Cost of goods sold", "Gross profit", "Operating expenses", "Net profit" };
            chart = api.Call("CreateChart", "lineNormal", chartData, chartNames, chartYears);
            SetChartSizes(chart, 10.06, 5.06, 1.67, 2);
            string moneyUnit = SeparateValueAndUnit(profitForecast[0].revenue).Item2;
            string verAxisTitle = "Amount (" + moneyUnit + ")";
            chart.Call("SetVerAxisTitle", verAxisTitle, 14, false);
            chart.Call("SetHorAxisTitle", "Year", 14, false);
            chart.Call("SetLegendFontSize", 14);
            CValue stroke = api.Call("CreateStroke", 1, chartGridFill);
            chart.Call("SetMinorVerticalGridlines", stroke);
            slide.Call("AddObject", chart);

            // break even analysis
            // create new slide
            slide = AddNewSlide(api, backgroundFill);
            // title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            AddTextToParagraph(api, paragraph, "Financial Model", 72, textFill, false, "center");
            // chart title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
            AddTextToParagraph(api, paragraph, "Break even analysis", 48, textFill, false, "center");
            // chart data
            Tuple<string, string> fixedCostsWithUnit = SeparateValueAndUnit(dataFinancial.break_even_analysis.fixed_costs);
            double fixedCosts = double.Parse(fixedCostsWithUnit.Item1);
            moneyUnit = fixedCostsWithUnit.Item2;
            double sellingPricePerUnit = double.Parse(SeparateValueAndUnit(dataFinancial.break_even_analysis.selling_price_per_unit).Item1);
            double variableCostPerUnit = double.Parse(SeparateValueAndUnit(dataFinancial.break_even_analysis.variable_cost_per_unit).Item1);
            int breakEvenPoint = dataFinancial.break_even_analysis.break_even_point;
            int step = breakEvenPoint / 4;
            CValue chartUnits = context.CreateArray(9);
            CValue chartRevenue = context.CreateArray(9);
            CValue chartTotalCosts = context.CreateArray(9);
            for (int i = 0; i < 9; i++)
            {
                int currUnits = i * step;
                chartUnits[i] = currUnits;
                chartRevenue[i] = currUnits * sellingPricePerUnit;
                chartTotalCosts[i] = fixedCosts + currUnits * variableCostPerUnit;
            }
            chartData = new CValue[] { chartRevenue, chartTotalCosts };
            chartNames = new CValue[] { "Revenue", "Total costs" };
            // create chart
            chart = api.Call("CreateChart", "lineNormal", chartData, chartNames, chartUnits);
            SetChartSizes(chart, 9.17, 5.06, 0.31, 2);
            verAxisTitle = "Amount (" + moneyUnit + ")";
            chart.Call("SetVerAxisTitle", verAxisTitle, 14, false);
            chart.Call("SetHorAxisTitle", "Units sold", 14, false);
            chart.Call("SetLegendFontSize", 14);
            chart.Call("SetMinorVerticalGridlines", stroke);
            chart.Call("SetLegendPos", "bottom");
            slide.Call("AddObject", chart);
            // break even point
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 3.11);
            AddTextToParagraph(api, paragraph, "Break even point:", 48, textFill, false, "center");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 3.51);
            AddTextToParagraph(api, paragraph, breakEvenPoint.ToString(), 128, textSpecialFill, false, "center", "Arial Black");
            paragraph = AddParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 4.38);
            AddTextToParagraph(api, paragraph, "units", 40, textFill, false, "center");

            // growth rates
            // create new slide
            slide = AddNewSlide(api, backgroundFill);
            // title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4);
            AddTextToParagraph(api, paragraph, "Financial Model", 72, textFill, false, "center");
            // chart title
            paragraph = AddParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2);
            AddTextToParagraph(api, paragraph, "Growth rates", 48, textFill, false, "center");
            // chart
            var growthRates = dataFinancial.growth_rates;
            chartYears = context.CreateArray(growthRates.Count);
            CValue chartGrowth = context.CreateArray(growthRates.Count);
            for (int i = 0; i < growthRates.Count; i++)
            {
                chartYears[i] = growthRates[i].year;
                chartGrowth[i] = growthRates[i].growth;
            }
            chartData = new CValue[] { chartGrowth };
            chart = api.Call("CreateChart", "lineNormal", chartData, context.CreateArray(0), chartYears);
            SetChartSizes(chart, 10.06, 5.06, 1.67, 2);
            chart.Call("SetVerAxisTitle", "Growth (%)", 14, false);
            chart.Call("SetHorAxisTitle", "Year", 14, false);
            chart.Call("SetMinorVerticalGridlines", stroke);
            slide.Call("AddObject", chart);

            // save and close
            builder.SaveFile(doctype, resultPath);
            builder.CloseFile();
            CDocBuilder.Destroy();
        }

        public static void AddTextToParagraph(CValue api, CValue paragraph, string text, int fontSize, CValue fill, bool isBold = false, string jc = "left", string fontFamily = "Arial")
        {
            CValue run = api.Call("CreateRun");
            run.Call("AddText", text);
            run.Call("SetFontSize", fontSize);
            run.Call("SetBold", isBold);
            run.Call("SetFill", fill);
            run.Call("SetFontFamily", fontFamily);
            paragraph.Call("AddElement", run);
            paragraph.Call("SetJc", jc);
        }

        public static CValue AddNewSlide(CValue api, CValue fill)
        {
            CValue slide = api.Call("CreateSlide");
            CValue presentation = api.Call("GetPresentation");
            presentation.Call("AddSlide", slide);
            slide.Call("SetBackground", fill);
            slide.Call("RemoveAllObjects");
            return slide;
        }

        public const int em_in_inch = 914400;
        // width, height, pos_x and pos_y are set in INCHES
        public static CValue AddParagraphToSlide(CValue api, CValue slide, double width, double height, double pos_x, double pos_y)
        {
            CValue shape = api.Call("CreateShape", "rect", width * em_in_inch, height * em_in_inch);
            shape.Call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
            CValue paragraph = shape.Call("GetDocContent").Call("GetElement", 0);
            slide.Call("AddObject", shape);
            return paragraph;
        }

        public static void SetChartSizes(CValue chart, double width, double height, double pos_x, double pos_y)
        {
            chart.Call("SetSize", width * em_in_inch, height * em_in_inch);
            chart.Call("SetPosition", pos_x * em_in_inch, pos_y * em_in_inch);
        }

        public static Tuple<string, string> SeparateValueAndUnit(string data)
        {
            int spacePos = data.IndexOf(' ');
            return Tuple.Create(data.Substring(0, spacePos), data[(spacePos + 1)..]);
        }

        public static string MakeBulletString(char bullet, int repeats)
        {
            string result = "";
            for (int i = 0; i < repeats; i++)
            {
                result += bullet;
                result += '\n';
            }
            return result;
        }
    }

    // Define classes to represent the JSON structure
    public class MarketOverviewData
    {
        public MarketData market { get; set; }
}

    public class MarketData
    {
        public string size { get; set; }
        public string growth_rate { get; set; }
        public List<string> trends { get; set; }
    }

    public class CompetitorsOverviewData
    {
        public List<CompetitorData> competitors { get; set; }
    }

    public class CompetitorData
    {
        public string name { get; set; }
        public string market_share { get; set; }
        public string recent_funding { get; set; }
        public List<string> products { get; set; }
    }

    public class AudienceOverviewData
    {
        public SocialTrendsData social_trends { get; set; }
        public DemographicsData demographics { get; set; }
    }

    public class SocialTrendsData
    {
        public List<string> positive_feedback { get; set; }
        public List<string> negative_feedback { get; set; }
    }

    public class DemographicsData
    {
        public string age_range { get; set; }
        public string location { get; set; }
        public string income_level { get; set; }
    }

    public class TrendsOverviewData
    {
        public List<SearchTrendData> search_trends { get; set; }
    }

    public class SearchTrendData
    {
        public string topic { get; set; }
        public string growth { get; set; }
    }

    public class FinancialModelData
    {
        public List<YearForecastData> profit_forecast { get; set; }
        public BreakEvenAnalysisData break_even_analysis { get; set; }
        public List<YearGrowthData> growth_rates { get; set; }
    }

    public class YearForecastData
    {
        public string year { get; set; }
        public string revenue { get; set; }
        public string cost_of_goods_sold { get; set; }
        public string gross_profit { get; set; }
        public string operating_expenses { get; set; }
        public string net_profit { get; set; }
    }

    public class BreakEvenAnalysisData
    {
        public string fixed_costs { get; set; }
        public string selling_price_per_unit { get; set; }
        public string variable_cost_per_unit { get; set; }
        public int break_even_point { get; set; }
    }

    public class YearGrowthData
    {
        public string year { get; set; }
        public string growth { get; set; }
    }
}
