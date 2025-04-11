'''
(c) Copyright Ascensio System SIA 2025

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
'''

# import docbuilder
import sys
sys.path.append('../../out/python')
import constants
sys.path.append(constants.BUILDER_DIR)
import docbuilder
# import other standard libraries
import json
import os

def addTextToParagraph(api, paragraph, text, font_size, fill, is_bold=False, jc='left', font_family='Arial'):
    run = api.Call('CreateRun')
    run.Call('AddText', text)
    run.Call('SetFontSize', font_size)
    run.Call('SetBold', is_bold)
    run.Call('SetFill', fill)
    run.Call('SetFontFamily', font_family)
    paragraph.Call('AddElement', run)
    paragraph.Call('SetJc', jc)

def addNewSlide(api, fill):
    slide = api.Call('CreateSlide')
    presentation = api.Call('GetPresentation')
    presentation.Call('AddSlide', slide)
    slide.Call('SetBackground', fill)
    slide.Call('RemoveAllObjects')
    return slide

em_in_inch = 914400
# width, height, pos_x and pos_y are set in INCHES
def addParagraphToSlide(api, slide, width, height, pos_x, pos_y):
    shape = api.Call('CreateShape', 'rect', width * em_in_inch, height * em_in_inch)
    shape.Call('SetPosition', pos_x * em_in_inch, pos_y * em_in_inch)
    paragraph = shape.Call('GetDocContent').Call('GetElement', 0)
    slide.Call('AddObject', shape)
    return paragraph

def setChartSizes(chart, width, height, pos_x, pos_y):
    chart.Call('SetSize', width * em_in_inch, height * em_in_inch)
    chart.Call('SetPosition', pos_x * em_in_inch, pos_y * em_in_inch)

def separateValueAndUnit(data):
    space_pos = data.find(' ')
    return (data[:space_pos], data[space_pos + 1:])

if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')

    # init docbuilder and create new pptx file
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(docbuilder.FileTypes.Presentation.PPTX)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    presentation = api.Call('GetPresentation')

    # init colors
    background_fill = api.Call('CreateSolidFill', api.Call('CreateRGBColor', 255, 255, 255))
    text_fill = api.Call('CreateSolidFill', api.Call('CreateRGBColor', 80, 80, 80))
    text_special_fill = api.Call('CreateSolidFill', api.Call('CreateRGBColor', 15, 102, 7))
    text_alt_fill = api.Call('CreateSolidFill', api.Call('CreateRGBColor', 230, 69, 69))
    chart_grid_fill = api.Call('CreateSolidFill', api.Call('CreateRGBColor', 134, 134, 134))
    master = presentation.Call('GetMaster', 0)
    color_scheme = master.Call('GetTheme').Call('GetColorScheme')
    color_scheme.Call('ChangeColor', 0, api.Call('CreateRGBColor', 15, 102, 7))

    # TITLE slide
    slide = presentation.Call('GetSlideByIndex', 0)
    slide.Call('SetBackground', background_fill)
    paragraph = slide.Call('GetAllShapes')[0].Call('GetContent').Call('GetElement', 0)
    addTextToParagraph(api, paragraph, 'GreenVibe Solutions', 120, text_special_fill, True, 'center', 'Arial Black')
    paragraph = slide.Call('GetAllShapes')[1].Call('GetContent').Call('GetElement', 0)
    addTextToParagraph(api, paragraph, '12.12.2024', 48, text_fill, jc='center')

    # MARKET OVERVIEW slide
    # parse JSON, obtained as Statista API response
    with open(os.path.join(resources_dir, 'data/statista_api_response.json'), 'r') as file_json:
        data = json.load(file_json)
    # create new slide
    slide = addNewSlide(api, background_fill)
    # title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
    addTextToParagraph(api, paragraph, 'Market Overview', 72, text_fill, jc='center')
    # market size
    market_size, market_size_unit = separateValueAndUnit(str(data['market']['size']))
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.58)
    addTextToParagraph(api, paragraph, 'Market size:', 48, text_fill, jc='center')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.97)
    addTextToParagraph(api, paragraph, market_size, 144, text_special_fill, jc='center', font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 3.06)
    addTextToParagraph(api, paragraph, market_size_unit, 48, text_fill, jc='center')
    # growth rate
    growth_rate, frequency = separateValueAndUnit(str(data['market']['growth_rate']))
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 7, 1.58)
    addTextToParagraph(api, paragraph, 'Growth rate:', 48, text_fill, jc='center')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 7, 1.97)
    addTextToParagraph(api, paragraph, growth_rate, 144, text_special_fill, jc='center', font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 7, 3.06)
    addTextToParagraph(api, paragraph, frequency, 48, text_fill, jc='center')
    # trends
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 3.75)
    addTextToParagraph(api, paragraph, 'Trends:', 48, text_fill, jc='center')
    paragraph = addParagraphToSlide(api, slide, 0.93, 2.92, 1.57, 4.31)
    addTextToParagraph(api, paragraph, '>\n' * len(data['market']['trends']), 72, text_special_fill, font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 9.21, 2.92, 2.1, 4.31)
    trends_text = ''
    for trend in data['market']['trends']:
        trends_text += str(trend) + '\n'
    addTextToParagraph(api, paragraph, trends_text, 72, text_special_fill, jc='center', font_family='Arial Black')

    # COMPETITORS OVERVIEW section
    # parse JSON, obtained as Statista API response
    with open(os.path.join(resources_dir, 'data/crunchbase_api_response.json'), 'r') as file_json:
        data = json.load(file_json)
    # create new slide
    slide = addNewSlide(api, background_fill)
    # title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
    addTextToParagraph(api, paragraph, 'Competitors Overview', 72, text_fill, jc='center')
    # chart header
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2)
    addTextToParagraph(api, paragraph, 'Market shares', 48, text_fill, jc='center')
    # get chart data
    others_share = 100
    shares = []
    competitors = []
    for competitor in data['competitors']:
        competitors.append(str(competitor['name']))
        share = int(competitor['market_share'][:-1])
        others_share -= share
        shares.append(share)
    shares.append(others_share)
    competitors.append('Others')
    # create a chart
    chart = api.Call('CreateChart', 'pie', [shares], [], competitors)
    setChartSizes(chart, 6.51, 5.9, 4.18, 1.49)
    chart.Call('SetLegendFontSize', 14)
    chart.Call('SetLegendPos', 'right')
    slide.Call('AddObject', chart)

    # create slide for every competitor with brief info
    for competitor in data['competitors']:
        # create new slide
        slide = addNewSlide(api, background_fill)
        # title
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
        addTextToParagraph(api, paragraph, 'Competitors Overview', 72, text_fill, jc='center')
        # header
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2)
        addTextToParagraph(api, paragraph, str(competitor['name']), 64, text_fill, jc='center')
        # recent funding
        paragraph = addParagraphToSlide(api, slide, 3.13, 0.8, 1.07, 2.65)
        addTextToParagraph(api, paragraph, 'Recent funding:', 48, text_fill)
        paragraph = addParagraphToSlide(api, slide, 8.9, 0.8, 4.19, 2.52)
        addTextToParagraph(api, paragraph, str(competitor['recent_funding']), 96, text_special_fill, font_family='Arial Black')
        # main products
        paragraph = addParagraphToSlide(api, slide, 3.13, 0.8, 1.07, 3.72)
        addTextToParagraph(api, paragraph, 'Main products:', 48, text_fill)
        paragraph = addParagraphToSlide(api, slide, 0.93, 3.53, 4.19, 3.72)
        addTextToParagraph(api, paragraph, '>\n' * len(competitor['products']), 72, text_special_fill, font_family='Arial Black')
        paragraph = addParagraphToSlide(api, slide, 7.97, 3.53, 5.12, 3.72)
        products_text = ''
        for product in competitor['products']:
            products_text += str(product) + '\n'
        addTextToParagraph(api, paragraph, products_text, 72, text_special_fill, font_family='Arial Black')

    # TARGET AUDIENCE section
    # parse JSON, obtained as Social Media Insights API response
    with open(os.path.join(resources_dir, 'data/smi_api_response.json'), 'r') as file_json:
        data = json.load(file_json)
    # create new slide
    slide = addNewSlide(api, background_fill)
    # title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
    addTextToParagraph(api, paragraph, 'Target Audience', 72, text_fill, jc='center')

    # demographics
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.33)
    addTextToParagraph(api, paragraph, 'Demographics:', 48, text_fill, jc='center')
    # age range
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 1.97)
    addTextToParagraph(api, paragraph, str(data['demographics']['age_range']), 128, text_special_fill, jc='center', font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 2.95)
    addTextToParagraph(api, paragraph, 'age range', 40, text_fill, jc='center')
    # location
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 3.68)
    addTextToParagraph(api, paragraph, str(data['demographics']['location']), 72, text_special_fill, jc='center', font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 4.27)
    addTextToParagraph(api, paragraph, 'location', 40, text_fill, jc='center')
    # income level
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 5.28)
    addTextToParagraph(api, paragraph, str(data['demographics']['income_level']), 56, text_special_fill, jc='center', font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 0.8, 5.83)
    addTextToParagraph(api, paragraph, 'income level', 40, text_fill, jc='center')

    # social trends
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 6, 1.33)
    addTextToParagraph(api, paragraph, 'Social trends:', 48, text_fill, jc='center')
    # positive feedback
    paragraph = addParagraphToSlide(api, slide, 0.63, 2.42, 7, 2.06)
    addTextToParagraph(api, paragraph, '+\n' * len(data['social_trends']['positive_feedback']), 52, text_special_fill, font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.56, 2.42, 7.67, 2.06)
    positive_feedback = ''
    for feedback in data['social_trends']['positive_feedback']:
        positive_feedback += str(feedback) + '\n'
    addTextToParagraph(api, paragraph, positive_feedback, 52, text_special_fill, font_family='Arial Black')
    # negative feedback
    paragraph = addParagraphToSlide(api, slide, 0.63, 2.42, 7, 4.55)
    addTextToParagraph(api, paragraph, '-\n' * len(data['social_trends']['negative_feedback']), 52, text_alt_fill, font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.56, 2.42, 7.67, 4.55)
    negative_feedback = ''
    for feedback in data['social_trends']['negative_feedback']:
        negative_feedback += str(feedback) + '\n'
    addTextToParagraph(api, paragraph, negative_feedback, 52, text_alt_fill, font_family='Arial Black')

    # SEARCH TRENDS section
    # parse JSON, obtained as Google Trends API response
    with open(os.path.join(resources_dir, 'data/google_trends_api_response.json'), 'r') as file_json:
        data = json.load(file_json)
    # create new slide
    slide = addNewSlide(api, background_fill)
    # title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
    addTextToParagraph(api, paragraph, 'Search Trends', 72, text_fill, jc='center')
    # add every trend on the slide
    offset_y = 1.43
    for trend in data['search_trends']:
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, offset_y)
        addTextToParagraph(api, paragraph, str(trend['topic']), 96, text_special_fill, jc='center', font_family='Arial Black')
        paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, offset_y + 0.8)
        addTextToParagraph(api, paragraph, str(trend['growth']), 40, text_fill, jc='center')
        offset_y += 1.25

    # FINANCIAL MODEL section
    # parse JSON, obtained from financial system
    with open(os.path.join(resources_dir, 'data/financial_model_data.json'), 'r') as file_json:
        data = json.load(file_json)
    # create new slide
    slide = addNewSlide(api, background_fill)
    # title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
    addTextToParagraph(api, paragraph, 'Financial Model', 72, text_fill, jc='center')
    # chart title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2)
    addTextToParagraph(api, paragraph, 'Profit forecast', 48, text_fill, jc='center')
    # chart
    chart_names = ['revenue', 'cost_of_goods_sold', 'gross_profit', 'operating_expenses', 'net_profit']
    chart_years = [str(entry['year']) for entry in data['profit_forecast']]
    chart_data = [[str(entry[key]) for entry in data['profit_forecast']] for key in chart_names]
    chart = api.Call('CreateChart', 'lineNormal', chart_data, ['Revenue', 'Cost of goods sold', 'Gross profit', 'Operating expenses', 'Net profit'], chart_years)
    setChartSizes(chart, 10.06, 5.06, 1.67, 2)
    money_unit = separateValueAndUnit(str(data['profit_forecast'][0]['revenue']))[1]
    chart.Call('SetVerAxisTitle', 'Amount (%s)' % money_unit, 14, False)
    chart.Call('SetHorAxisTitle', 'Year', 14, False)
    chart.Call('SetLegendFontSize', 14)
    stroke = api.Call('CreateStroke', 1, chart_grid_fill)
    chart.Call('SetMinorVerticalGridlines', stroke)
    slide.Call('AddObject', chart)

    # break even analysis
    # create new slide
    slide = addNewSlide(api, background_fill)
    # title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
    addTextToParagraph(api, paragraph, 'Financial Model', 72, text_fill, jc='center')
    # chart title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2)
    addTextToParagraph(api, paragraph, 'Break even analysis', 48, text_fill, jc='center')
    # chart
    fixed_costs_str, money_unit = separateValueAndUnit(str(data['break_even_analysis']['fixed_costs']))
    fixed_costs = float(fixed_costs_str)
    selling_price_per_unit = float(separateValueAndUnit(data['break_even_analysis']['selling_price_per_unit'])[0])
    variable_cost_per_unit = float(separateValueAndUnit(data['break_even_analysis']['variable_cost_per_unit'])[0])
    break_even_point = data['break_even_analysis']['break_even_point']
    step = break_even_point / 4
    chart_units = range(0, int(break_even_point * 2 + step), int(step))
    chart_revenue = [units * selling_price_per_unit for units in chart_units]
    chart_total_costs = [fixed_costs + units * variable_cost_per_unit for units in chart_units]
    # create chart
    chart = api.Call('CreateChart', 'lineNormal', [chart_revenue, chart_total_costs], ['Revenue', 'Total costs'], list(chart_units))
    setChartSizes(chart, 9.17, 5.06, 0.31, 2)
    chart.Call('SetVerAxisTitle', 'Amount (%s)' % money_unit, 14, False)
    chart.Call('SetHorAxisTitle', 'Units sold', 14, False)
    chart.Call('SetLegendFontSize', 14)
    chart.Call('SetMinorVerticalGridlines', stroke)
    chart.Call('SetLegendPos', 'bottom')
    slide.Call('AddObject', chart)
    # break even point
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 3.11)
    addTextToParagraph(api, paragraph, 'Break even point:', 48, text_fill, jc='center')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 3.51)
    addTextToParagraph(api, paragraph, str(break_even_point), 128, text_special_fill, jc='center', font_family='Arial Black')
    paragraph = addParagraphToSlide(api, slide, 5.62, 0.8, 8.4, 4.38)
    addTextToParagraph(api, paragraph, 'units', 40, text_fill, jc='center')

    # growth rates
    # create new slide
    slide = addNewSlide(api, background_fill)
    # title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 0.4)
    addTextToParagraph(api, paragraph, 'Financial Model', 72, text_fill, jc='center')
    # chart title
    paragraph = addParagraphToSlide(api, slide, 11.8, 0.8, 0.8, 1.2)
    addTextToParagraph(api, paragraph, 'Growth rates', 48, text_fill, jc='center')
    # chart
    chart_years = [str(entry['year']) for entry in data['growth_rates']]
    chart_growth = [str(entry['growth']) for entry in data['growth_rates']]
    chart = api.Call('CreateChart', 'lineNormal', [chart_growth], [], chart_years)
    setChartSizes(chart, 10.06, 5.06, 1.67, 2)
    chart.Call('SetVerAxisTitle', 'Growth (%)', 14, False)
    chart.Call('SetHorAxisTitle', 'Year', 14, False)
    stroke = api.Call('CreateStroke', 1, chart_grid_fill)
    chart.Call('SetMinorVerticalGridlines', stroke)
    slide.Call('AddObject', chart)

    # save and close
    result_path = os.getcwd() + '/result.pptx'
    builder.SaveFile(docbuilder.FileTypes.Presentation.PPTX, result_path)
    builder.CloseFile()
