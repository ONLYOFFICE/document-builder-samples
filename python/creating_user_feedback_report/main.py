'''
(c) Copyright Ascensio System SIA 2024

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


def set_table_style(range):
    range.Call('SetRowHeight', 24)
    range.Call('SetAlignVertical', 'center')

    line_style = 'Thin'
    range.Call('SetBorders', 'Top', line_style, color_black)
    range.Call('SetBorders', 'Left', line_style, color_black)
    range.Call('SetBorders', 'Right', line_style, color_black)
    range.Call('SetBorders', 'Bottom', line_style, color_black)
    range.Call('SetBorders', 'InsideHorizontal', line_style, color_black)
    range.Call('SetBorders', 'InsideVertical', line_style, color_black)


def fill_average_sheet(worksheet, feedback_data):
    average_values = [['Question', 'Average Rating', 'Number of Responses']]
    result = {}
    for record in feedback_data:
        for item in record['feedback']:
            if item['question'] not in result:
                result[item['question']] = [item['answer']['rating']]
            else:
                result[item['question']].append(item['answer']['rating'])
    for key, value in result.items():
        average_values.append([key, round(sum(value) / len(value), 1), len(value)])

    cols_count = len(average_values[0]) - 1
    start_cell = worksheet.Call('GetRangeByNumber', 0, 0)
    end_cell = worksheet.Call('GetRangeByNumber', len(average_values) - 1, cols_count)

    average_range = worksheet.Call('GetRange', start_cell, end_cell)
    set_table_style(average_range)
    worksheet.Call('GetRange', worksheet.Call('GetRangeByNumber', 1, 1), end_cell).Call('SetAlignHorizontal', 'center')

    header_row = worksheet.Call('GetRange', start_cell, worksheet.Call('GetRangeByNumber', 0, cols_count))
    header_row.Call('SetBold', True)

    average_range.Call('SetValue', average_values)
    average_range.Call('AutoFit', False, True)

    return len(average_values)


def fill_personal_ratings_and_comments(worksheet, feedback_data):
    header_values = [['Date', 'Question', 'Comment', 'Rating', 'Average User Rating']]
    cols_count = len(header_values[0]) - 1
    start_cell = worksheet.Call('GetRangeByNumber', 0, 0)

    header_row = worksheet.Call('GetRange', start_cell, worksheet.Call('GetRangeByNumber', 0, cols_count))
    header_row.Call('SetValue', header_values)
    header_row.Call('SetBold', True)

    rows_count = 1
    for record in feedback_data:
        # Count and fill user feedback
        user_feedback = []
        avg_rating = 0
        for item in record['feedback']:
            user_feedback.append([item['question'], item['answer']['comment'], item['answer']['rating']])
            avg_rating += item['answer']['rating']

        user_rows_count = len(user_feedback) - 1
        avg_rating = round(avg_rating / len(user_feedback), 1)

        # Fill date
        date_cell = worksheet.Call(
            'GetRange',
            worksheet.Call('GetRangeByNumber', rows_count, 0),
            worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, 0),
        )
        date_cell.Call('Merge', False)
        date_cell.Call('SetValue', record['date'])

        # Fill ratings
        user_range = worksheet.Call(
            'GetRange',
            worksheet.Call('GetRangeByNumber', rows_count, 1),
            worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, cols_count - 1),
        )
        user_range.Call('SetValue', user_feedback)

        # Count average rating
        rating_cell = worksheet.Call(
            'GetRange',
            worksheet.Call('GetRangeByNumber', rows_count, cols_count),
            worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, cols_count),
        )
        rating_cell.Call('Merge', False)
        rating_cell.Call('SetValue', avg_rating)

        # If rating <= 2, highlight it
        if avg_rating <= 2:
            worksheet.Call(
                'GetRange',
                worksheet.Call('GetRangeByNumber', rows_count, 0),
                worksheet.Call('GetRangeByNumber', rows_count + user_rows_count, cols_count),
            ).Call('SetFillColor', color_orange)

        # Update rows count
        rows_count += len(user_feedback)

    # Format table
    rows_count -= 1
    result_range = worksheet.Call(
        'GetRange',
        start_cell,
        worksheet.Call('GetRangeByNumber', rows_count, cols_count),
    )
    set_table_style(result_range)
    worksheet.Call(
        'GetRange',
        worksheet.Call('GetRangeByNumber', 1, cols_count - 1),
        worksheet.Call('GetRangeByNumber', rows_count, cols_count),
    ).Call('SetAlignHorizontal', 'center')
    result_range.Call('AutoFit', False, True)

    return rows_count + 1


def create_column_chart(worksheet, data_range, title):
    chart = worksheet.Call('AddChart', data_range, False, 'bar', 2, 135.38 * 36000, 81.28 * 36000)
    chart.Call('SetPosition', 0, 0, 0, 0)
    chart.Call('SetTitle', title, 16)


def create_line_chart(api, worksheet, feedback_data, title):
    average_day_rating = [['Date', 'Rating']]
    result = {}
    for record in feedback_data:
        if record['date'] not in result.keys():
            result[record['date']] = []
        for item in record['feedback']:
            result[record['date']].append(item['answer']['rating'])
    for key, value in result.items():
        average_day_rating.append([key, round(sum(value) / len(value), 1)])

    data_range = f'$E$1:$F${len(average_day_rating)}'
    worksheet.Call('GetRange', data_range).Call('SetValue', average_day_rating)
    chart = worksheet.Call('AddChart', f'Charts!{data_range}', False, 'scatter', 2, 135.38 * 36000, 81.28 * 36000)
    chart.Call('SetPosition', 0, 0, 18, 0)
    chart.Call('SetSeriesFill', color_blue, 0, False)
    stroke = api.Call(
        'CreateStroke',
        0.5 * 36000,
        api.Call('CreateSolidFill', color_grey),
    )
    chart.Call('SetSeriesOutLine', stroke, 0, False)
    chart.Call('SetTitle', title, 16)
    chart.Call('SetMajorHorizontalGridlines', api.Call('CreateStroke', 0, api.Call('CreateNoFill')))


def create_pie_chart(api, worksheet, data_range, title):
    worksheet.Call('GetRange', '$A$1:$C$2').Call(
        'SetValue',
        [
            ['Negative', 'Neutral', 'Positive'],
            [f'=COUNTIF({data_range}, "<=2")', f'=COUNTIF({data_range}, "=3")', f'=COUNTIF({data_range}, ">=4")'],
        ],
    )
    chart = worksheet.Call('AddChart', 'Charts!$A$1:$C$2', True, 'pie', 2, 135.38 * 36000, 81.28 * 36000)
    chart.Call('SetPosition', 9, 0, 0, 0)
    chart.Call('SetTitle', title, 16)
    chart.Call('SetDataPointFill', api.Call('CreateSolidFill', api.Call('CreateRGBColor', 237, 125, 49)), 0, 0)
    chart.Call('SetDataPointFill', api.Call('CreateSolidFill', color_grey), 0, 1)
    chart.Call('SetDataPointFill', api.Call('CreateSolidFill', color_blue), 0, 2)
    stroke = api.Call(
        'CreateStroke',
        0.5 * 36000,
        api.Call('CreateSolidFill', api.Call('CreateRGBColor', 255, 255, 255)),
    )
    chart.Call('SetSeriesOutLine', stroke, 0, False)


if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/user_feedback_data.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new pdf file
    doctype = docbuilder.FileTypes.Spreadsheet.XLSX
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(doctype)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']

    # Set main colors
    color_black = api.Call('CreateColorFromRGB', 0, 0, 0)
    color_orange = api.Call('CreateColorFromRGB', 237, 125, 49)
    color_grey = api.Call('CreateRGBColor', 128, 128, 128)
    color_blue = api.Call('CreateRGBColor', 91, 155, 213)

    # Get current worksheet
    worksheet1 = api.Call('GetActiveSheet')

    # Create worksheet with average values
    worksheet1.Call('SetName', 'Average')
    table1_rows_count = fill_average_sheet(worksheet1, data)

    # Create worksheet with comments and personal ratings
    api.Call('AddSheet', 'Comments')
    worksheet2 = api.Call('GetActiveSheet')
    table2_rows_count = fill_personal_ratings_and_comments(worksheet2, data)

    # Create worksheet with charts
    api.Call('AddSheet', 'Charts')
    worksheet3 = api.Call('GetActiveSheet')
    create_column_chart(worksheet3, f'Average!$A$2:$B${table1_rows_count}', 'Average ratings')
    create_line_chart(api, worksheet3, data, 'Dynamics of the average ratings')
    create_pie_chart(api, worksheet3, f'Comments!$D$1:$D${table2_rows_count}', 'Shares of reviews')

    # Set first worksheet active
    worksheet1.Call('SetActive')

    # save and close
    result_path = os.getcwd() + '/result.xlsx'
    builder.SaveFile(doctype, result_path)
    builder.CloseFile()
