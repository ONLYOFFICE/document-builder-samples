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


def fill_header(paragraph, text):
    paragraph.Call('AddText', text)
    paragraph.Call('SetFontSize', 28)
    paragraph.Call('SetBold', True)
    set_spacing_after(paragraph, 50)


def set_spacing_after(paragraph, spacing):
    paragraph.Call('SetSpacingAfter', spacing)


def set_numbering(paragraph, num_lvl):
    paragraph.Call('SetNumbering', num_lvl)


def create_details_header(api, text):
    paragraph = api.Call('CreateParagraph')
    paragraph.Call('AddText', text)
    paragraph.Call('SetBold', True)
    paragraph.Call('SetItalic', True)
    set_spacing_after(paragraph, 40)
    return paragraph


def setup_requisites_style(paragraph, num_lvl, set_spacing):
    if set_spacing:
        set_spacing_after(paragraph, 20)
    if num_lvl:
        set_numbering(paragraph, num_lvl)


def create_requisites_paragraph(api, title, details, num_lvl=None, set_spacing=True, set_title_bold=True):
    paragraph = api.Call('CreateParagraph')
    title_run = paragraph.Call('AddText', f'{title}: ')
    if set_title_bold:
        title_run.Call('SetBold', True)
    else:
        title_run.Call('SetItalic', True)
    details_run = paragraph.Call('AddText', details)
    details_run.Call('SetItalic', True)
    setup_requisites_style(paragraph, num_lvl, set_spacing)
    return paragraph


def setup_table_style(document, table):
    # table size
    table.Call('SetWidth', 'percent', 100)
    table.Call('Select')
    table_range = document.Call('GetRangeBySelect')
    table_paragraphs = table_range.Call('GetAllParagraphs')
    for i in range(table_paragraphs.GetLength()):
        para_pr = table_paragraphs.Get(i).Call('GetParaPr')
        para_pr.Call('SetSpacingBefore', 40)
        para_pr.Call('SetSpacingAfter', 40)

    # table borders
    params = ("single", 4, 0, 0, 0, 0)
    table.Call('SetTableBorderTop', *params)
    table.Call('SetTableBorderBottom', *params)
    table.Call('SetTableBorderLeft', *params)
    table.Call('SetTableBorderRight', *params)
    table.Call('SetTableBorderInsideH', *params)
    table.Call('SetTableBorderInsideV', *params)


def get_cell_content(cell):
    return cell.Call('GetContent').Call('GetElement', 0)


def format_sum(value):
    return f'${value:,}'


def fill_table_content(table, items):
    table_headers = ['Description', 'Quantity', 'Unit Price', 'Total']
    table_fields = ['description', 'quantity', 'unit_price', 'total']

    # fill table header
    header_row = table.Call('GetRow', 0)
    for i, field_name in enumerate(table_headers):
        header_cell = get_cell_content(header_row.Call('GetCell', i))
        header_cell.Call('AddText', field_name)
        header_cell.Call('SetBold', True)
        header_cell.Call('SetJc', 'center')

    # fill items
    for i in range(len(items)):
        row = table.Call('GetRow', i + 1)
        for j, key in enumerate(table_fields):
            cell = get_cell_content(row.Call('GetCell', j))
            match key:
                case 'unit_price' | 'total':
                    cell.Call('AddText', format_sum(items[i][key]))
                case _:
                    cell.Call('AddText', str(items[i][key]))


if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/commercial_offer_data.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new docx file
    doctype = docbuilder.FileTypes.Document.DOCX
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(doctype)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    document = api.Call('GetDocument')

    # page margins
    section = document.Call('GetFinalSection')
    section.Call('SetPageMargins', 1440, 1280, 1440, 1280)

    # DOCUMENT STYLE
    para_pr = document.Call('GetDefaultParaPr')
    para_pr.Call('SetSpacingAfter', 100)
    text_pr = document.Call('GetDefaultTextPr')
    text_pr.Call('SetFontSize', 24)
    text_pr.Call('SetFontFamily', 'Times New Roman')

    # DOCUMENT HEADER
    header = document.Call('GetElement', 0)
    fill_header(header, 'COMMERCIAL OFFER TEMPLATE')

    # document requisites
    document.Call('Push', create_requisites_paragraph(api, 'Offer No.', data["offer"]["number"]))
    document.Call('Push', create_requisites_paragraph(api, 'Date', data["offer"]["date"], set_spacing=False))

    # bullet numbering
    bullet_numbering = document.Call('CreateNumbering', 'bullet')
    b_num_lvl = bullet_numbering.Call('GetLevel', 0)

    # SELLER INFORMATION
    seller_header = create_details_header(api, 'SELLER INFORMATION')
    document.Call('Push', seller_header)

    # seller details
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Company Name', data["seller"]["company_name"], num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Address', data["seller"]["address"], num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Tax ID (TIN)', data["seller"]["tin"], num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Contact Information', '', num_lvl=b_num_lvl),
    )

    # contact details
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Phone',
            data["seller"]["contact"]["phone"],
            num_lvl=b_num_lvl,
            set_title_bold=False,
        ),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Email', data["seller"]["contact"]["email"],
            num_lvl=b_num_lvl,
            set_spacing=False,
            set_title_bold=False,
        ),
    )

    # BUYER INFORMATION
    buyer_header = create_details_header(api, 'BUYER INFORMATION')
    document.Call('Push', buyer_header)

    # buyer details
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Company Name', data["buyer"]["company_name"], num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Address', data["buyer"]["address"], num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Contact Person', data["buyer"]["contact_person"], num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Email', data["buyer"]["email"], num_lvl=b_num_lvl, set_spacing=False),
    )

    # OFFER DETAILS
    table_header = api.Call('CreateParagraph')
    fill_header(table_header, 'OFFER DETAILS')
    document.Call('Push', table_header)

    # table content
    items_table = api.Call('CreateTable', 4, len(data["offer_details"]) + 1)
    document.Call('Push', items_table)
    setup_table_style(document, items_table)
    fill_table_content(items_table, data["offer_details"])

    # TOTALS
    totals = create_details_header(api, 'TOTALS')
    document.Call('Push', totals)
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Subtotal', format_sum(data["totals"]["subtotal"]), num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Discount', format_sum(data["totals"]["discount"]), num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api, 'Tax (e.g., 20% VAT)', format_sum(data["totals"]["tax"]), num_lvl=b_num_lvl),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(api,'Total Amount', format_sum(data["totals"]["total"]), num_lvl=b_num_lvl, set_spacing=False),
    )

    # TERMS AND CONDITIONS
    seller_header = create_details_header(api, 'SELLER INFORMATION')
    document.Call('Push', seller_header)

    # numbering
    numbering = document.Call('CreateNumbering', 'numbered')
    d_num_lvl = numbering.Call('GetLevel', 0)
    d_num_lvl.Call('SetCustomType', 'decimal', '%1.', 'left')

    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Validity Period',
            data["terms_and_conditions"]["validity_period"],
            num_lvl=d_num_lvl,
        ),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Payment Terms',
            data["terms_and_conditions"]["payment_terms"],
            num_lvl=d_num_lvl,
        ),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Delivery Terms',
            data["terms_and_conditions"]["delivery_terms"],
            num_lvl=d_num_lvl,
        ),
    )
    document.Call(
        'Push',
        create_requisites_paragraph(
            api,
            'Additional Notes',
            data["terms_and_conditions"]["additional_notes"],
            num_lvl=d_num_lvl,
            set_spacing=False,
        ),
    )

    # SIGNATURE
    sign_header = api.Call('CreateParagraph')
    sign_header.Call('AddText', 'Signature:')
    sign_header.Call('SetBold', True)
    document.Call('Push', sign_header)

    sign_details = api.Call('CreateParagraph')
    sign_details.Call(
        'AddText',
        f'{data["seller"]["authorized_person"]["full_name"]}, {data["seller"]["authorized_person"]["position"]}',
    )
    sign_details.Call('AddLineBreak')
    sign_details.Call('AddText', data["seller"]["company_name"])
    document.Call('Push', sign_details)

    # save and close
    result_path = os.getcwd() + '/result.docx'
    builder.SaveFile(doctype, result_path)
    builder.CloseFile()
