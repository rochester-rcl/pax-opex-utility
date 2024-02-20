import PySimpleGUI as sg
from preservica_pre_ingest import create_container, file_hash_list, folder_ds_files, representation_preservation_access, create_pax, pax_metadata, ao_opex_metadata, opex_container_metadata
from preservica_post_ingest import test_connection, move_opex_aspace, move_aspace_trash, quality_control
from documentation import pre_ingest_documentation, post_ingest_documentation, utilities_documentation, about
from utilities import folder_desc_size
from icon import icon

vsc_theme = {'BACKGROUND': '#252525',
            'TEXT': '#7edcf0',
            'INPUT': '#181818',
            'TEXT_INPUT': '#5dd495',
            'SCROLL': '#252525',
            'BUTTON': ('#dcdca1', '#181818'),
            'PROGRESS': ('#dcdca1', '#181818'),
            'BORDER': 1, 'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 0}

alt_background = '#181818'
init_color = '#5dd495'
update_color = '#7edcf0'
summary_color = '#dcdca1'
conclusion_color = '#e4a177'
button_font = 'Courier 14'
program_icon = icon()
sg.set_options(font='Courier 12')
sg.theme_add_new('VSC Theme', vsc_theme)
sg.theme('VSC Theme')
sg.user_settings_filename(path='.')

preing_projfol_frame = sg.Frame('Project Folder', [
        [sg.Input(focus=True, expand_x=True), sg.FolderBrowse(key='-PROJFOLDER-')],
        [sg.Push(), sg.Text('Generate File/Hash Manifest?'), sg.Checkbox('', default=True, key='-MANIFEST-'), sg.Combo(['', 'MD5', 'SHA-1'], default_value=sg.user_settings_get_entry('algorithm', default=''), key='-ALG-', readonly=True), sg.Push()]
        ], expand_x=True, pad=5)

preing_workord_frame = sg.Frame('Work Order Spreadsheet Information', [
        [sg.Input(expand_x=True), sg.FileBrowse(file_types=(('Work Order', '*.xlsx'),), key='-WORKORDER-')],
        [sg.Text('Worksheet Name', size=14, justification='right'), sg.Input(size=41, default_text='digitization_work_order_report', key='-WORKSHEET-')],
        [sg.Text('Max Row', size=14, justification='right'), sg.Input(size=5, key='-MAXROW-')]
        ], expand_x=True, pad=5)

preing_fileext_frame = sg.Frame('Preservation/Access File Extensions', [
        [sg.Text('Preservation', size=12, justification='right'), sg.Input(size=5, default_text ='.tif', key='-PRES-'), sg.Text('Access', size=12, justification='right'), sg.Input(size=5, default_text='.pdf', key='-ACC-'), sg.Push()]
        ], expand_x=True, pad=5)

preing_options_frame = sg.Frame('Options', [
        [sg.Text('Format Dates', size=18, justification='right'), sg.Checkbox('', default=sg.user_settings_get_entry('format dates', default=True), key='-DATES-')],
        [sg.Text('Filename Delimiter'), sg.Input(size=1, default_text=sg.user_settings_get_entry('file delimiter', default='-'), key='-DELIMITER-')]
        ], expand_x=True, pad=5)

tab1 = sg.Tab('Pre-Ingest', [
    [preing_projfol_frame],
    [preing_workord_frame],
    [preing_fileext_frame],
    [preing_options_frame],
    [sg.Push(), sg.Button('Create PAX Objects and OPEX Metadata', pad=(10, 5)), sg.Push()]
    ], key='-TAB_1-')

posting_pres_frame = sg.Frame('Preservica Administrator Credentials', [
        [sg.Text('Username', size=10, justification='right'), sg.Input(default_text=sg.user_settings_get_entry('username', default=''), key='-USERNAME-')],
        [sg.Text('Password', size=10, justification='right'), sg.Input(default_text=sg.user_settings_get_entry('password', default=''), password_char='*', key='-PASSWORD-')],
        [sg.Text('Tenancy', size=10, justification='right'), sg.Input(default_text=sg.user_settings_get_entry('tenancy', default=''), key='-TENANCY-')],
        [sg.Text('Server', size=10, justification='right'), sg.Input(default_text=sg.user_settings_get_entry('server', default='us.preservica.com'), key='-SERVER-')],
        [sg.Text('Using 2FA?', size=10, justification='right'), sg.Checkbox('', default=sg.user_settings_get_entry('twofa_cb', default=False), key='-2FACB-')],
        [sg.Text('2FA Token', size=10, justification='right'), sg.Input(default_text=sg.user_settings_get_entry('twofa_value', default=''), password_char='*', key='-2FA-')],
        [sg.Push(), sg.Button('Test Connection', pad=(10,5)), sg.Push()]
        ], pad=5, expand_x=True)

posting_refid_frame = sg.Frame('Preservica Folder Ref IDs', [
        [sg.Text('OPEX Folder Ref', size=19, justification='right'), sg.Input(size=(36,1), default_text=sg.user_settings_get_entry('opex', default=''), key='-OPEX-')],
        [sg.Text('ASpace Folder Ref', size=19, justification='right'), sg.Input(size=(36,1), default_text=sg.user_settings_get_entry('aspace', default=''), key='-ASPACE-')],
        [sg.Text('Trash Folder Ref', size=19, justification='right'), sg.Input(size=(36,1), default_text=sg.user_settings_get_entry('trash', default=''), key='-TRASH-')],
        [sg.Push(), sg.Button('Move From OPEX to ASpace Link', pad=(10, 5)), sg.Push()], 
        [sg.Push(), sg.Button('Move From ASpace Link to Trash', pad=(10, 5)), sg.Push()]
        ], pad=5)

posting_qc_frame = sg.Frame('Quality Control', [
        [sg.Input(size=48, expand_x=True), sg.FileBrowse(file_types=(('File/Hash Manifest', '*.csv'),), key='-QC-')],
        [sg.Push(), sg.Button('Quality Control', pad=(10, 5)), sg.Push()]
        ], expand_x=True, pad=5)

tab2 = sg.Tab('Post-Ingest', [
    [posting_pres_frame],
    [posting_refid_frame],
    [posting_qc_frame]
    ], key='-TAB_2-')

reports_frame = sg.Frame('Reports', [
        [sg.Text('Output Storage Size of Given Folder')],
        [sg.Text('Folder Ref:'), sg.Input(size=(36,1), key='-FOLDER_REPORT-')],
        [sg.Button('Generate Storage Report on Folder', pad=(10, 5))]
        ], expand_x=True, pad=5)

tab3 = sg.Tab('Utilities', [
    [reports_frame]
    ], key='-TAB_3-')

ops_tabgroup = sg.TabGroup([
[tab1, tab2, tab3]
], pad=5, key='-TAB_GROUP-')

ops_col = sg.Column([
    [ops_tabgroup]
], vertical_alignment='t')

settings_frame = sg.Frame('Settings', [
        [sg.Button('Save', font=(button_font), pad=5), sg.Button('Display', font=(button_font), pad=5)]
        ], pad=5)

output_col = sg.Column([
[sg.Output(key='-OUTPUT-', size = (60, 27), background_color=alt_background, expand_x=True, expand_y=True, pad=5)],
[settings_frame, sg.Button('Help', font=(button_font), pad=5), sg.Button('Clear', font=(button_font), pad=5), sg.Button('About', font=(button_font), pad=5), sg.Push(), sg.Button('Exit', font=(button_font), pad=5), sg.Sizegrip()]
], expand_x=True, expand_y=True, vertical_alignment='t')

layout = [[ops_col, output_col]]

window = sg.Window('OPEX/PAX Utility', layout, finalize=True, resizable=True, icon=program_icon)
window.set_min_size(window.size)
mline = window['-OUTPUT-']

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    proj_path = values['-PROJFOLDER-']
    manifest = values['-MANIFEST-']
    alg_choice = values['-ALG-']
    work_order = values['-WORKORDER-']
    work_sheet = values['-WORKSHEET-']
    max_row = values['-MAXROW-']
    rep_pres = values['-PRES-']
    rep_acc = values['-ACC-']
    format_dates = values['-DATES-']
    filename_delimiter = values['-DELIMITER-']
    username = values['-USERNAME-']
    password = values['-PASSWORD-']
    tenancy = values['-TENANCY-']
    server = values['-SERVER-']
    twofactorcb = values['-2FACB-']
    twofactorkey = values['-2FA-']
    opex_folder = values['-OPEX-']
    aspace_folder = values['-ASPACE-']
    trash_folder = values['-TRASH-']
    qual_control = values['-QC-']
    folder_report = values['-FOLDER_REPORT-']
    if event == 'Save':
        sg.user_settings_set_entry('generate manifest', values['-MANIFEST-'])
        sg.user_settings_set_entry('algorithm', values['-ALG-'])
        sg.user_settings_set_entry('format dates', values['-DATES-'])
        sg.user_settings_set_entry('file delimiter', values['-DELIMITER-'])
        sg.user_settings_set_entry('username', values['-USERNAME-'])
        sg.user_settings_set_entry('password', values['-PASSWORD-'])
        sg.user_settings_set_entry('tenancy', values['-TENANCY-'])
        sg.user_settings_set_entry('server', values['-SERVER-'])    
        sg.user_settings_set_entry('twofa_cb', values['-2FACB-'])
        sg.user_settings_set_entry('twofa_value', values['-2FA-'])
        sg.user_settings_set_entry('opex', values['-OPEX-'])
        sg.user_settings_set_entry('aspace', values['-ASPACE-'])
        sg.user_settings_set_entry('trash', values['-TRASH-'])
    if event == 'Display':
        mline.update('')
        count = 0
        for key in sg.user_settings():
            if (count % 2) == 0:
                mline.update('{}: {}\n'.format(key, sg.user_settings()[key]), append=True, text_color_for_value=update_color)
                count += 1
            else:
                mline.update('{}: {}\n'.format(key, sg.user_settings()[key]), append=True, text_color_for_value=update_color, background_color_for_value=alt_background)
                count += 1
    if event == 'Create PAX Objects and OPEX Metadata':
        mline.update('')
        create_container(window, mline, init_color, summary_color, proj_path, work_order)
        if manifest == True:
            file_hash_list(window, mline, alt_background, init_color, update_color, summary_color, proj_path, alg_choice)
        folder_ds_files(window, mline, alt_background, init_color, update_color, summary_color, proj_path, filename_delimiter)
        representation_preservation_access(window, mline, alt_background, init_color, update_color, summary_color, proj_path, rep_pres, rep_acc)
        create_pax(window, mline, alt_background, init_color, update_color, summary_color, proj_path)
        pax_metadata(window, mline, alt_background, init_color, update_color, summary_color, proj_path, work_order, work_sheet, max_row, format_dates)
        ao_opex_metadata(window, mline, alt_background, init_color, update_color, summary_color, proj_path, work_order, work_sheet, max_row)
        opex_container_metadata(window, mline, init_color, update_color, proj_path)
        mline.update('****PAX OBJECTS AND OPEX METADATA CREATED****\n', append=True, text_color_for_value=conclusion_color)
        window.refresh()
    if event == 'Test Connection':
        test_connection(window, mline, alt_background, init_color, update_color, summary_color, username, password, tenancy, server, twofactorcb, twofactorkey)
    if event == 'Move From OPEX to ASpace Link':
        move_opex_aspace(window, mline, alt_background, init_color, update_color, summary_color, username, password, tenancy, server, twofactorcb, twofactorkey, opex_folder, aspace_folder)
    if event == 'Move From ASpace Link to Trash':
        move_aspace_trash(window, mline, alt_background, init_color, update_color, summary_color, username, password, tenancy, server, twofactorcb, twofactorkey, aspace_folder, trash_folder)
    if event == 'Quality Control':
        quality_control(window, mline, alt_background, init_color, update_color, summary_color, username, password, tenancy, server, twofactorcb, twofactorkey, qual_control, work_order, work_sheet, max_row)
    if event == 'Generate Storage Report on Folder':
        folder_desc_size(window, mline, alt_background, init_color, update_color, summary_color, username, password, tenancy, server, twofactorcb, twofactorkey, folder_report)
    if event == 'Help':
        if values['-TAB_GROUP-'] == '-TAB_1-':
            mline.update('')
            pre_ingest_documentation(window, mline, init_color, update_color)
            mline.set_vscroll_position(0)
        if values['-TAB_GROUP-'] == '-TAB_2-':
            mline.update('')
            post_ingest_documentation(window, mline, init_color, update_color, summary_color)
            mline.set_vscroll_position(0)
        if values['-TAB_GROUP-'] == '-TAB_3-':
            mline.update('')
            utilities_documentation(window, mline, init_color, update_color)
            mline.set_vscroll_position(0)
    if event == 'Clear':
        mline.update('')
    if event == 'About':
        mline.update('')
        about(window, mline, update_color)
window.close()
