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


def create_paragraph(api, text, is_bold=False, font_size=None, jc=None):
    paragraph = api.Call('CreateParagraph')
    paragraph.Call('AddText', text)
    paragraph.Call('SetBold', is_bold)
    if font_size:
        paragraph.Call('SetFontSize', font_size)
    if jc:
        paragraph.Call('SetJc', jc)
    return paragraph


def create_run(api, text, is_bold=False, font_size=None):
    run = api.Call('CreateRun')
    run.Call('AddText', text)
    run.Call('SetBold', is_bold)
    if font_size:
        run.Call('SetFontSize', font_size)
    return run


def set_numbering(paragraph, num_lvl):
    paragraph.Call('SetNumbering', num_lvl)

def set_spacing_after(paragraph, spacing):
    paragraph.Call('SetSpacingAfter', spacing)


def create_conditions_desc_paragraph(api, text):
    # create paragraph with first line indentation
    paragraph = create_paragraph(api, text)
    paragraph.Call('SetIndFirstLine', 400)
    return paragraph


def add_participant_to_paragraph(api, paragraph, p_type, details):
    paragraph.Call('Push', create_run(api, f'{p_type}: ', is_bold=True))
    paragraph.Call('Push', create_run(api, details))


def create_numbered_section(api, text, num_lvl):
    paragraph = create_paragraph(api, text, is_bold=True)
    set_numbering(paragraph, num_lvl)
    set_spacing_after(paragraph, 50)
    return paragraph


def create_work_condition(api, title, text, num_lvl, set_spacing=True):
    paragraph = api.Call('CreateParagraph')
    set_numbering(paragraph, num_lvl)
    if set_spacing:
        set_spacing_after(paragraph, 20)
    paragraph.Call('SetJc', 'left')
    paragraph.Call('Push', create_run(api, f'{title}: ', is_bold=True))
    paragraph.Call('Push', create_run(api, text))
    return paragraph


def fill_signer(api, cell, title):
    paragraph = cell.Call('GetContent').Call('GetElement', 0)
    paragraph.Call('SetJc', 'left')
    paragraph.Call('Push', create_run(api, title, is_bold=True))
    for text in [
        'Name: __________________________',
        'Signature: _______________________',
        'Date: ___________________________'
    ]:
        paragraph.Call('AddLineBreak')
        paragraph.Call('Push', create_run(api, text))


if __name__ == '__main__':
    resources_dir = os.path.normpath('../../resources')
    # parse JSON
    with open(os.path.join(resources_dir, 'data/employment_agreement_data.json'), 'r') as file_json:
        data = json.load(file_json)

    # init docbuilder and create new docx file
    doctype = docbuilder.FileTypes.Document.OFORM_PDF
    builder = docbuilder.CDocBuilder()
    builder.CreateFile(doctype)

    context = builder.GetContext()
    global_obj = context.GetGlobal()
    api = global_obj['Api']
    document = api.Call('GetDocument')

    # DOCUMENT STYLE
    para_pr = document.Call('GetDefaultParaPr')
    para_pr.Call('SetJc', 'both')
    text_pr = document.Call('GetDefaultTextPr')
    text_pr.Call('SetFontSize', 24)
    text_pr.Call('SetFontFamily', 'Times New Roman')

    # DOCUMENT HEADER
    header = document.Call('GetElement', 0)
    header.Call('AddText', 'EMPLOYMENT AGREEMENT')
    header.Call('SetFontSize', 28)
    header.Call('SetBold', True)

    header_desc = create_paragraph(
        api,
        f'This Employment Agreement ("Agreement") is made and entered into on {data["date"]} by and between:',
    )
    set_spacing_after(header_desc, 50)
    document.Call('Push', header_desc)

    # PARTICIPANTS OF THE DOCUMENT
    participants = create_paragraph(api, '', jc='left')
    add_participant_to_paragraph(
        api,
        participants,
        'Employer',
        f'{data["employer"]["name"]}, located at {data["employer"]["address"]}.',
    )
    participants.Call('AddLineBreak')
    add_participant_to_paragraph(
        api,
        participants,
        'Employee',
        f'{data["employee"]["full_name"]}, residing at {data["employee"]["address"]}.',
    )
    document.Call('Push', participants)
    document.Call('Push', create_paragraph(api, 'The parties agree to the following terms and conditions:'))

    # AGREEMENT CONDITIONS
    # create numbering
    numbering = document.Call('CreateNumbering', 'numbered')
    numbering_lvl = numbering.Call('GetLevel', 0)
    numbering_lvl.Call('SetCustomType', 'decimal', '%1.', 'left')
    numbering_lvl.Call('SetSuff', 'space')

    # position and duties
    document.Call('Push', create_numbered_section(api, 'POSITION AND DUTIES', numbering_lvl))
    document.Call(
        'Push',
        create_conditions_desc_paragraph(
            api,
            f'The Employee is hired as {data["position_and_duties"]["job_title"]}. The Employee shall perform '
            'their duties as outlined by the Employer and comply with all applicable policies and guidelines.',
        ),
    )

    # compensation
    document.Call('Push', create_numbered_section(api, 'COMPENSATION', numbering_lvl))
    document.Call(
        'Push',
        create_conditions_desc_paragraph(
            api,
            f'The Employee will receive a salary of {data["compensation"]["salary"]} '
            f'{data["compensation"]["currency"]} {data["compensation"]["frequency"]} ({data["compensation"]["type"]}), '
            "payable in accordance with the Employer's payroll schedule and subject to lawful deductions.",
        ),
    )

    # probationary period
    document.Call('Push', create_numbered_section(api, 'PROBATIONARY PERIOD', numbering_lvl))
    document.Call(
        'Push',
        create_conditions_desc_paragraph(
            api,
            f'The Employee will serve a probationary period of {data["probationary_period"]["duration"]}. '
            'During this period, the Employer may terminate this Agreement with '
            f"{data['probationary_period']['terminate']} days' notice if performance is deemed unsatisfactory.",
        ),
    )

    # work conditions
    document.Call('Push', create_numbered_section(api, 'WORK CONDITIONS', numbering_lvl))
    conditions_text = create_conditions_desc_paragraph(
        api,
        "The following terms apply to the Employee's working conditions:",
    )
    set_spacing_after(conditions_text, 50)
    document.Call('Push', conditions_text)

    # create bullet numbering
    bullet_numbering = document.Call('CreateNumbering', 'bullet')
    bul_num_lvl = bullet_numbering.Call('GetLevel', 0)

    document.Call(
        'Push',
        create_work_condition(api, 'Working Hours', data["work_conditions"]["working_hours"], bul_num_lvl),
    )
    document.Call(
        'Push',
        create_work_condition(api, 'Work Schedule', data["work_conditions"]["work_schedule"], bul_num_lvl),
    )
    document.Call(
        'Push',
        create_work_condition(api, 'Benefits', ', '.join(data["work_conditions"]["benefits"]), bul_num_lvl),
    )
    document.Call(
        'Push',
        create_work_condition(
            api,
            'Other terms',
            ', '.join(data["work_conditions"]["other_terms"]),
            bul_num_lvl,
            set_spacing=False,
        ),
    )

    # TERMINATION
    document.Call('Push', create_numbered_section(api, 'TERMINATION', numbering_lvl))
    document.Call(
        'Push',
        create_conditions_desc_paragraph(
            api,
            f'Either party may terminate this Agreement by providing {data["termination"]["notice_period"]} '
            'written notice. The Employer reserves the right to terminate employment immediately for cause, '
            'including but not limited to misconduct or breach of Agreement.',
        ),
    )

    # GOVERNING LAW
    document.Call('Push', create_numbered_section(api, 'GOVERNING LAW', numbering_lvl))
    document.Call(
        'Push',
        create_conditions_desc_paragraph(
            api,
            f'This Agreement is governed by the laws of {data["governing_law"]["jurisdiction"]}, '
            'and any disputes arising under this Agreement will be resolved in accordance with these laws.',
        ),
    )

    # ENTIRE AGREEMENT
    document.Call('Push', create_numbered_section(api, 'ENTIRE AGREEMENT', numbering_lvl))
    document.Call(
        'Push',
        create_conditions_desc_paragraph(
            api,
            'This document constitutes the entire Agreement between the parties and supersedes all prior '
            'agreements. Any amendments must be made in writing and signed by both parties.',
        ),
    )

    # SIGNATURES
    # create table
    table = api.Call('CreateTable', 2, 2)
    # set table properties
    table.Call('SetWidth', 'percent', 100)
    # fill table
    table_title = table.Call('GetRow', 0)
    title_paragraph = table_title.Call('MergeCells').Call('GetContent').Call('GetElement', 0)
    title_paragraph.Call('Push', create_run(api, 'SIGNATURES', is_bold=True, font_size=24))
    fill_signer(api, table.Call('GetCell', 1, 0), 'Employer')
    fill_signer(api, table.Call('GetCell', 1, 1), 'Employee')
    document.Call('Push', table)

    # save and close
    result_path = os.getcwd() + '/result.pdf'
    builder.SaveFile(doctype, result_path)
    builder.CloseFile()
