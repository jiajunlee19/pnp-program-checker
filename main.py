try:
    import time
    import os
    import sys
    import pandas as pd
    import xml.etree.ElementTree as ETree
    import numpy as np
    import getpass
    # import xlwings as xw
    from utils.logger import logger_init
    from utils.Common_Functions_64 import removeExtraDelimiter, ExpandSeries, digit_to_nondigit, split_into_rows, extract_num_from_end, string_remove_duplicate, flatten

except ImportError as IE:
    print(f"Import Error: {str(IE)}")
    time.sleep(5)


def init():
    '''init'''

    # Get path_main and transform into absolute path (so it works for onedrive path too)
    sharepoint_online_path = f"https://microncorp-my.sharepoint.com/personal/{getpass.getuser()}_micron.com/Documents/"
    sharepoint_local_path = f"C:\\Users\\{getpass.getuser()}\\OneDrive - Micron Technology, Inc\\"
    path_main = os.path.dirname(os.path.realpath(sys.argv[0]))
    path_main = path_main.replace(sharepoint_online_path, sharepoint_local_path)
    path_main = path_main.replace("/", "\\")

    # Define working folder paths
    path_590 = f"{path_main}\\BOM_590"
    path_MCTO = f"{path_main}\\MCTO"
    path_program = f"{path_main}\\PNP_PROGRAM"
    filename_checker = "CHECKER.xlsx"
    path_checker = f"{path_main}\\{filename_checker}"

    # Init logger
    try:
        df_settings = pd.read_excel(path_checker, sheet_name='settings')
        loglevel = list(df_settings['LOG_LEVEL'])[0]
        loglevel_error = False 
    except (ValueError, KeyError):
        loglevel_error = True
        loglevel = 'INFO'

    if loglevel_error:
        log.warning('LOG_LEVEL is not defined in settings, setting to INFO...')

    log = logger_init('PNP_PROGRAM_CHECKER.log', f"{path_main}\\Log", 'w', loglevel)
    log.info(f"Running main.py in {path_main} with loglevel = {loglevel}")

    input_columns = ['BOM', 'MCTO', 'PV','PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2']
    output_columns = input_columns + ['COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?', 'CHEKCER']
    log.info(f"input_columns = {input_columns}")
    log.info(f"output_columns = {output_columns}")

    return log, path_main, path_590, path_MCTO, path_program, path_checker, input_columns, output_columns


def main(log, path_main, path_590, path_MCTO, path_program, path_checker, input_columns, output_columns):
    '''main'''

    log.info(f"BOM_590 folder = {path_590}")
    log.info(f"MCTO folder = {path_MCTO}")
    log.info(f"program folder = {path_program}")
    log.info(f"Checker file = {path_checker}")

    try:
        df_settings = pd.read_excel(path_checker, sheet_name='settings')
        SAP_SOURCE = list(df_settings['SAP_SOURCE'])[0]
    except (ValueError, KeyError):
        log.warning('SAP_SOURCE is not defined in settings, setting to manual...')
        SAP_SOURCE = 'manual'

    log.info(f"SAP_SOURCE = {SAP_SOURCE}")

    exclude_comp_prefix = ('590', '550', '540', '542', '561', '562', 'ECN')

    # Read main excel workbook
    log.info('Reading checker file...')
    # try:
    #     # the excel file is not opened
    #     workbook = xw.Book(path_checker)
    # except:
    #     # the excel file is opened
    #     workbook = xw.Book(filename_checker)
    # inputSheet = workbook.sheets['PNP_PROGRAM_CHECKER_INPUT'].used_range.value

    # Create df_input for input sheet
    log.info('Creating dataframe for input sheet...')
    # df_input = pd.DataFrame(inputSheet)
    df_input = pd.read_excel(path_checker, sheet_name='CHECKER')

    # log.debug('Transposing df_input, keep 1st 4 rows and retransposing...')
    # df_input = df_input.T.head(4).T

    # df_input.columns = df_input.iloc[0]
    # log.debug('Stripping off 1st header row...')
    # df_input = df_input[1:]

    log.debug('Trimming all input columns...')
    for input_column in input_columns:
        df_input[input_column] = df_input[input_column].astype(str)
        df_input[input_column] = df_input[input_column].str.strip().str.upper().str.lstrip('0')
    df_input = df_input.replace([' '], ['']).replace(['NAN'], ['']).replace([''], [np.NaN], regex=True)
    log.debug(f"\n{df_input.head(5).to_string(index=False)}")

    log.debug('Dropping null rows...')
    df_input.dropna(how='any', subset=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1'], inplace=True)
    log.debug(f"\n{df_input.head(5).to_string(index=False)}")

    if len(df_input) < 1:
        raise ConnectionAbortedError ('There is no input to be processed, force exiting application...')

    log.debug('Dropping duplicates...')    
    df_input.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2'], keep='first', inplace=True)
    log.debug(f"\n{df_input.head(5).to_string(index=False)}")

    log.debug('Removing duplicates of selected 590, MCTO, program...')
    selected_590 = set(df_input['BOM'])
    selected_MCTO = set(df_input['MCTO'])
    try:
        selected_program = (set(df_input['PNP_PROGRAM_SIDE1']).union(set(df_input['PNP_PROGRAM_SIDE2']))).remove('')
    except KeyError:
        selected_program = (set(df_input['PNP_PROGRAM_SIDE1']).union(set(df_input['PNP_PROGRAM_SIDE2'])))

    # log.debug('Hardcoding selected files...')
    # selected_590 = {'590-624664'}
    # selected_MCTO = {'705043'}
    # selected_program = {'3440CB-PD0-M5-IT', '3440CB-SD0-M5-IT'}

    log.info(f"Selected_590 = {selected_590}")
    log.info(f"Selected_MCTO = {selected_MCTO}")
    log.info(f"Selected_program = {selected_program}")

    bom_columns = ['BOM', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR']
    mcto_columns = ['MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR']

    if SAP_SOURCE == 'manual':

        # Scan for the selected files only to save resources
        log.info('Scanning files_590 and files_MCTO...')
        scan_files_590 = os.scandir(path_590)
        scan_files_MCTO = os.scandir(path_MCTO)

        file_590 = {f.path for f in scan_files_590 if f.name[-4:].lower() == '.csv' and any (matcher in f.name for matcher in selected_590)}
        file_MCTO = {f.path for f in scan_files_MCTO if f.name[-4:].lower() == '.csv' and any (matcher in f.name for matcher in selected_MCTO)}

        log.info(f"Matched file_590 = {file_590}")
        log.info(f"Matched file_MCTO = {file_MCTO}")

        # Continue only if at least one 590,MCTO file is found
        if len(file_590) < 1 and len(file_590) < 1:
            df_checker_all = pd.DataFrame(columns=output_columns)
            df_checker_all.to_excel(f"{path_main}\\SCRIPT_OUTPUT.xksx", sheet_name='OUTPUT', index=False)
            raise ConnectionAbortedError ('There is no selected 590 or MCTO file found, force exiting application...')

        # Combine all files_590
        log.info(f"Starting to read {str(len(file_590))} BOM_590 files...")
        df_590 = pd.DataFrame()
        for file in file_590:
            log.info(f"Reading: {file}...")
            df = pd.read_csv(file, sep='\t', skiprows=9, usecols=[1,3,5,10], skip_blank_lines=True, skipinitialspace=True, on_bad_lines='warn')
            df.columns = df.columns.str.strip()
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Renaming columns...')
            df = df.rename(columns={'Object no.':'COMPONENT', 'Quantity':'QUANTITY', 'Material Description':'COMPDESC', 'Reference Designator':'DESIGNATOR'})
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Create BOM column with component starting with 590...')
            df['BOM'] = np.where(df['COMPONENT'].str.startswith('590'), df['COMPONENT'], np.NaN)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Trimming all bom columns...')
            df = df[bom_columns]
            for input_column in bom_columns:
                df[input_column] = df[input_column].astype(str)
                df[input_column] = df[input_column].str.strip().str.upper().str.lstrip('0')
            df = df.replace([' '], ['']).replace(['NAN'], ['']).replace([''], [np.NaN], regex=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Front-filling BOM...')
            df['BOM'].ffill(inplace=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Removing rows with null designator...')
            df = df[~df.DESIGNATOR.isnull()]
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Front-filling...')
            df.ffill(inplace=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug(f"Removing rows with comp_prefix = {exclude_comp_prefix}...")
            df = df[~df.COMPONENT.str.startswith(exclude_comp_prefix)]
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Removing rows with comp_prefox = 511 and compdesc contains TH AE or THAE...')
            df = df[~(df.COMPONENT.str.startswith('511') & (df.COMPDESC.str.contains('TH AE') | df.COMPDESC.str.contains('THAE')))]
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Converting quantity string into int...')
            df['QUANTITY'] = df['QUANTITY'].replace([','], ['.'], regex=True)
            df['QUANTITY'] = df['QUANTITY'].str.split('.').str[0].str.strip()
            df['QUANTITY'] = df['QUANTITY'].fillna('0').astype(int)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Grouping designator...')
            df = df.groupby(['BOM', 'COMPONENT', 'COMPDESC', 'QUANTITY'])['DESIGNATOR'].apply(','.join).reset_index()
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Removing extra delimiter from designator...')
            df['DESIGNATOR'] = df['DESIGNATOR'].apply(removeExtraDelimiter)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Expanding designator series...')
            df['DESIGNATOR'] = df['DESIGNATOR'].apply(ExpandSeries).str.replace(' ', '')
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            if df['DESIGNATOR'].str.contains('-').any():
                log.warning(f"Designators are not expanded, skipping {file}...")
                continue

            log.debug('Dropping duplicates...')
            df.drop_duplicates(subset=['BOM', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR'], keep='last', inplace=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug(f"Concating {str(len(df))} rows into df_590...")
            df_590 = pd.concat([df_590, df], ignore_index=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")
        log.info(f"Total of {str(len(df_590))} rows detected in BOM_590 files.")

        # Combine all MCTO files
        log.info(f"Starting to read {str(len(file_MCTO))} MCTO files...")
        df_MCTO = pd.DataFrame()
        for file in file_MCTO:
            filename = file.rsplit('\\', 1)[-1]
            filename_without_ext = filename.rsplit('.', 1)[0]
            try:
                PV = filename_without_ext.rsplit('_', 1)[1]
            except IndexError:
                PV = '1'

            log.info(f"Reading: {file}...")
            df = pd.read_csv(file, sep='\t', skiprows=9, usecols=[1,4,6,11], skip_blank_lines=True, skipinitialspace=True, on_bad_lines='warn')
            df.columns = df.columns.str.strip()
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Renaming columns...')
            df = df.rename(columns={'Object no.':'COMPONENT', 'Quantity':'QUANTITY', 'Material Description':'COMPDESC', 'Reference Designator':'DESIGNATOR'})
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Create MCTO column with component not null and compdesc, quantity, designator are null...')
            df['MCTO'] = np.where(~(df['COMPONENT'].isna()) & (df['COMPDESC'].isnull()) & (df['QUANTITY'].isnull()) & (df['DESIGNATOR'].isnull()), df['COMPONENT'], np.NaN)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Adding PV columns...')
            df['PV'] = PV
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Trimming all mcto columns...')
            df = df[mcto_columns]
            for input_column in mcto_columns:
                df[input_column] = df[input_column].astype(str)
                df[input_column] = df[input_column].str.strip().str.upper().str.lstrip('0')
            df = df.replace([' '], ['']).replace(['NAN'], ['']).replace([''], [np.NaN], regex=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Front-filling MCTO...')
            df['MCTO'].ffill(inplace=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Removing rows with null designator...')
            df = df[~df.DESIGNATOR.isnull()]
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Front-filling...')
            df.ffill(inplace=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug(f"Removing rows with comp_prefix = {exclude_comp_prefix}...")
            df = df[~df.COMPONENT.str.startswith(exclude_comp_prefix)]
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Converting quantity string into int...')
            df['QUANTITY'] = df['QUANTITY'].replace([','], ['.'], regex=True)
            df['QUANTITY'] = df['QUANTITY'].str.split('.').str[0].str.strip()
            df['QUANTITY'] = df['QUANTITY'].fillna('0').astype(int)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Grouping designator...')
            df = df.groupby(['MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY'])['DESIGNATOR'].apply(','.join).reset_index()
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Removing extra delimiter from designator...')
            df['DESIGNATOR'] = df['DESIGNATOR'].apply(removeExtraDelimiter)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug('Expanding designator series...')
            df['DESIGNATOR'] = df['DESIGNATOR'].apply(ExpandSeries).str.replace(' ', '')
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            if df['DESIGNATOR'].str.contains('-').any():
                log.warning(f"Designators are not expanded, skipping {file}...")
                continue

            log.debug('Dropping duplicates...')
            df.drop_duplicates(subset=['MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR'], keep='last', inplace=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")

            log.debug(f"Concating {str(len(df))} rows into df_MCTO...")
            df_MCTO = pd.concat([df_MCTO, df], ignore_index=True)
            log.debug(f"\n{df.head(5).to_string(index=False)}")
        log.info(f"Total of {str(len(df_MCTO))} rows detected in MCTO files.")

    else:
        log.info('Loading SAP database...')
        from settings import DB_TYPE
        from db_connection import connect_db
        connection, conn_response = connect_db(db_type=DB_TYPE)

        if connection is None:
            raise ConnectionAbortedError(conn_response)
        
        log.info(conn_response)
        
        log.info('Running query_BOM_590...')
        query_BOM_590 = '''
                    SELECT bh.SAP_mat_no as BOM, 
                        bi.compnt_no as COMPONENT, bi.compnt_desc as COMPDESC, 
                        CASE
                            WHEN CAST(bi.compnt_qty as INT) >= 1000 THEN CAST(CAST(bi.compnt_qty as INT)/1000 as INT)
                            ELSE CAST(bi.compnt_qty as INT)
                        END as QUANTITY, 
                        bi.item_text as DESIGNATOR
                    FROM [SAP_PP].[dbo].[SAP_BOM_item] bi
                    inner join [SAP_PP].[dbo].[SAP_BOM_header] bh on bi.BOM_no = bh.BOM_no
                    inner join [SAP_PP].[dbo].[material_master] mm on bh.SAP_mat_no = mm.material_no
                    where mm.material_group_code in ('590') and bh.SAP_mat_no in ({0});
                '''
        query_BOM_590 = query_BOM_590.format(','.join('?' * len(selected_590)))
        params_BOM_590 = tuple(flatten(selected_590))
        try:
            df_590 = pd.read_sql(sql=query_BOM_590, con=connection, params=params_BOM_590)
        except Exception:
            raise ConnectionAbortedError ('Failed to run query_BOM_590, force exiting application...')
        log.info(f"Total of {str(len(df_590))} rows detected in query_BOM_590.")

        log.debug('Trimming all bom columns...')
        df_590 = df_590[bom_columns]
        for input_column in bom_columns:
            df_590[input_column] = df_590[input_column].astype(str)
            df_590[input_column] = df_590[input_column].str.strip().str.upper().str.lstrip('0')
        df_590 = df_590.replace([' '], ['']).replace(['NAN'], ['']).replace([''], [np.NaN], regex=True)
        log.debug(f"\n{df_590.head(5).to_string(index=False)}")

        log.debug('Dropping null rows...')
        df_590.dropna(how='any', subset=bom_columns, inplace=True)
        log.debug(f"\n{df_590.head(5).to_string(index=False)}")

        log.debug(f"Removing rows with comp_prefix = {exclude_comp_prefix}...")
        df_590 = df_590[~df_590.COMPONENT.str.startswith(exclude_comp_prefix)]
        log.debug(f"\n{df_590.head(5).to_string(index=False)}")

        log.debug('Removing rows with comp_prefix = 511 and compdesc contains TH AE or THAE...')
        df_590 = df_590[~(df_590.COMPONENT.str.startswith('511') & (df_590.COMPDESC.str.contains('TH AE') | df_590.COMPDESC.str.contains('THAE')))]
        log.debug(f"\n{df_590.head(5).to_string(index=False)}")

        log.debug('Expanding designator series...')
        df_590['DESIGNATOR'] = df_590['DESIGNATOR'].apply(ExpandSeries).str.replace(' ', '')

        if df_590['DESIGNATOR'].str.contains('-').any():
            raise ConnectionAbortedError ('Designators are not expanded, force exiting application...')

        log.debug(f"\n{df_590.head(5).to_string(index=False)}")
        log.info(f"Total of {str(len(df_590))} rows detected in df_590.")

        log.info('Running query_MCTO...')
        query_MCTO = '''
                    SELECT REPLACE(bh.SAP_mat_no, '000000000000', '') as MCTO, bi.alt_BOM_type as PV,
                        bi.compnt_no as COMPONENT, bi.compnt_desc as COMPDESC, 
                        CASE
                            WHEN CAST(bi.compnt_qty as INT) >= 1000 THEN CAST(CAST(bi.compnt_qty as INT)/1000 as INT)
                            ELSE CAST(bi.compnt_qty as INT)
                        END as QUANTITY, 
                        bi.item_text as DESIGNATOR
                    FROM [SAP_PP].[dbo].[SAP_BOM_item] bi
                    inner join [SAP_PP].[dbo].[SAP_BOM_header] bh on bi.BOM_no = bh.BOM_no and bi.alt_BOM_type = bh.alt_BOM_type
                    inner join [SAP_PP].[dbo].[material_master] mm on bh.SAP_mat_no = mm.material_no
                    where mm.material_group_code in ('MCTO', '002', 'MFGPN') and REPLACE(bh.SAP_mat_no, '000000000000', '') in ({0});
                '''
        query_MCTO = query_MCTO.format(','.join('?' * len(selected_MCTO)))
        params_MCTO = tuple(flatten(selected_MCTO))
        try:
            df_MCTO = pd.read_sql(sql=query_MCTO, con=connection, params=params_MCTO)
        except Exception:
            raise ConnectionAbortedError ('Failed to run query_MCTO, force exiting application...')
        log.info(f"Total of {str(len(df_MCTO))} rows detected in query_MCTO.")

        log.debug('Trimming all mcto columns...')
        df_MCTO = df_MCTO[mcto_columns]
        for input_column in mcto_columns:
            df_MCTO[input_column] = df_MCTO[input_column].astype(str)
            df_MCTO[input_column] = df_MCTO[input_column].str.strip().str.upper().str.lstrip('0')
        df_MCTO = df_MCTO.replace([' '], ['']).replace(['NAN'], ['']).replace([''], [np.NaN], regex=True)
        log.debug(f"\n{df_MCTO.head(5).to_string(index=False)}")

        log.debug('Dropping null rows...')
        df_MCTO.dropna(how='any', subset=mcto_columns, inplace=True)
        log.debug(f"\n{df_MCTO.head(5).to_string(index=False)}")

        log.debug(f"Removing rows with comp_prefix = {exclude_comp_prefix}...")
        df_MCTO = df_MCTO[~df_MCTO.COMPONENT.str.startswith(exclude_comp_prefix)]
        log.debug(f"\n{df_MCTO.head(5).to_string(index=False)}")

        log.debug('Expanding designator series...')
        df_MCTO['DESIGNATOR'] = df_MCTO['DESIGNATOR'].apply(ExpandSeries).str.replace(' ', '')

        if df_MCTO['DESIGNATOR'].str.contains('-').any():
            raise ConnectionAbortedError ('Designators are not expanded, force exiting application...')
        
        log.debug(f"\n{df_MCTO.head(5).to_string(index=False)}")
        log.info(f"Total of {str(len(df_MCTO))} rows detected in query_MCTO.")

        if connection is not None:
            connection.close()

    # Recursively call scandir inclusive of subfolders for filename matching
    def scan_dir_file(path):
        for f in os.scandir(path):
            if f.is_file() and (f.name[-3:].lower() == '.pp' or f.name[-4:].lower() == '.pp7') and any (matcher in f.name for matcher in selected_program):
                yield f.path
            elif f.is_dir():
                yield from scan_dir_file(f.path)
    file_program = {f for f in scan_dir_file(path_program)}


    log.info(f"Matched file_program = {file_program}")

    # Continue only if at least one program file is found
    if len(file_program) < 1:
        df_checker_all = pd.DataFrame(columns=output_columns)
        df_checker_all.to_excel(f"{path_main}\\SCRIPT_OUTPUT.xksx", sheet_name='OUTPUT', index=False)
        raise ConnectionAbortedError ('There is no selected program file found, force exiting application...')

    # Combining all program files
    log.info(f"Starting to read {str(len(file_program))} program files...")
    all_items = []
    all_feeder_items = []
    all_action_items = []
    for file in file_program:
        log.info(f"Reading: {file}...")

        filename = file.rsplit('\\', 1)[-1]
        log.debug(f"filename = {filename}")

        filename_without_ext = filename.rsplit('.', 1)[0]
        log.debug(f"filename_without_ext = {filename_without_ext}")

        file_ext = filename.rsplit('.', 1)[-1]
        log.debug(f"file_ext = {file_ext}")

        xmldata = file

        store_items = []
        store_feeder_items = []
        store_action_items = []

        if file_ext.lower() == 'pp7':
            ETree.register_namespace('', 'http://api.assembleon.com/pp7/v1')
            prstree = ETree.parse(xmldata)
            root = prstree.getroot()
            pp_url = '{http://api.assembleon.com/pp7/v1}'
            log.debug(f"Setting pp_url to {pp_url}...")
            for GeneralInfo in root.iter(f"{pp_url}General"):
                sMachine = '3'
                sCycleTime_Raw = GeneralInfo.attrib.get('cycleTime')
                sCycleTime = sCycleTime_Raw[:-3] + '.' + sCycleTime_Raw[-3:]

            for BoardInfo in root.iter(f"{pp_url}Board"):
                sProgramName = BoardInfo.attrib.get('id')
                for ComponentInfo in BoardInfo.iter(f"{pp_url}Component"):
                    sPartNumber = ComponentInfo.attrib.get('partNumber')
                    sREFDES = ComponentInfo.attrib.get('refDes')
                    sBoardNumber = ComponentInfo.attrib.get('circuitNumber')

                    store_items = [sProgramName, sMachine, sPartNumber, sREFDES, sBoardNumber, sCycleTime]
                    all_items.append(store_items)

            for FeedSectionInfo in root.iter(f"{pp_url}FeedSection"):
                sSectionNumber = FeedSectionInfo.attrib.get('number')
                sTrolleyType = FeedSectionInfo.attrib.get('type')
                for FeederInfo in FeedSectionInfo.iter(f"{pp_url}Feeder"):
                    sFeederNumber = FeederInfo.attrib.get('slotNumber')
                    sFeederType = FeederInfo.attrib.get('type')
                    for LaneInfo in FeederInfo.iter(f"{pp_url}FeederLane"):
                        sLaneNumber = LaneInfo.attrib.get('number')
                        sPartNumber = LaneInfo.attrib.get('partNumber')
                        sShape = LaneInfo.attrib.get('shapeId')
                        store_feeder_items = [sProgramName, sMachine, sSectionNumber, sTrolleyType, sFeederNumber, sFeederType, sLaneNumber, sPartNumber, sShape]
                        all_feeder_items.append(store_feeder_items)

            # One feeder has 1 robot with 2 heads
            for PickInfo in root.iter(f"{pp_url}Pick"):
                sSectionNumber = PickInfo.attrib.get('feedSectionNumber')
                sRobotNumber = PickInfo.attrib.get('robotNumber')
                sHeadNumber = PickInfo.attrib.get('headNumber')
                sREFDES = PickInfo.attrib.get('refDes')
                sCircuitNumber = PickInfo.attrib.get('circuitNumber')
                sFeederNumber = PickInfo.attrib.get('feederSlotNumber')
                sLaneNumber = PickInfo.attrib.get('feederLaneNumber')
                store_action_items = [sProgramName, sMachine, sSectionNumber, sRobotNumber, sHeadNumber, sREFDES, sCircuitNumber, sFeederNumber, sLaneNumber]
                all_action_items.append(store_action_items)

        else:
            ETree.register_namespace('', 'http://api.assembleon.com/pp/v2')
            prstree = ETree.parse(xmldata)
            root = prstree.getroot()
            pp_url = '{http://api.assembleon.com/pp/v2}'
            log.debug(f"Setting pp_url to {pp_url}...")
            for GeneralInfo in root.iter(f"{pp_url}General"):
                sMachine = GeneralInfo.attrib.get('positionInLine')
                sCycleTime_Raw = GeneralInfo.attrib.get('cycleTime')
                sCycleTime = sCycleTime_Raw[:-3] + '.' + sCycleTime_Raw[-3:]

            for BoardInfo in root.iter(f"{pp_url}Board"):
                sProgramName = BoardInfo.attrib.get('id')
                # for ComponentInfo in BoardInfo.iter(f"{pp_url}Component"):
                #     sPartNumber = ComponentInfo.attrib.get('partNumber')
                #     sREFDES = ComponentInfo.attrib.get('refDes')
                #     sBoardNumber = ComponentInfo.attrib.get('circuitNumber')

                #     store_items = [sProgramName, sMachine, sPartNumber, sREFDES, sBoardNumber, sCycleTime]
                #     all_items.append(store_items)

            for SectionInfo in root.iter(f"{pp_url}Section"):
                sSectionNumber = SectionInfo.attrib.get('number')
                for TrolleyInfo in SectionInfo.iter(f"{pp_url}Trolley"):
                    sTrolleyType = TrolleyInfo.attrib.get('type')
                    for FeederInfo in TrolleyInfo.iter(f"{pp_url}Feeder"):
                        sFeederNumber = FeederInfo.attrib.get('number')
                        sFeederType = FeederInfo.attrib.get('type')
                        for LaneInfo in FeederInfo.iter(f"{pp_url}Lane"):
                            sLaneNumber = LaneInfo.attrib.get('number')
                            sPartNumber = LaneInfo.attrib.get('partNumber')
                            sShape = LaneInfo.attrib.get('shapeId')
                            store_feeder_items = [sProgramName, sMachine, sSectionNumber, sTrolleyType, sFeederNumber, sFeederType, sLaneNumber, sPartNumber, sShape]
                            all_feeder_items.append(store_feeder_items)

            # Each section has 4 robots, total 5 sections with 20 robots, each with 1 head
            sHeadNumber = '1'
            robots_per_section = 4
            section_number = 1
            log.debug(f"Assuming each section has {str(int(robots_per_section))} robots, each with {sHeadNumber} head...")
            for a, ActionInfo in enumerate(root.iter(f"{pp_url}Actions")):
                sSectionNumber = str(int(section_number))
                if (a+1) % robots_per_section == 0:
                    section_number += 1
                sRobotNumber = ActionInfo.attrib.get('robotNumber')
                for IndexInfo in ActionInfo.iter(f"{pp_url}Index"):
                    for PickInfo in IndexInfo.iter(f"{pp_url}Pick"):
                        sREFDES = PickInfo.attrib.get('refDes')
                        sCircuitNumber = PickInfo.attrib.get('circuitNumber')
                        sFeederNumber = PickInfo.attrib.get('feederNumber')
                        sLaneNumber = PickInfo.attrib.get('laneNumber')
                        store_action_items = [sProgramName, sMachine, sSectionNumber, sRobotNumber, sHeadNumber, sREFDES, sCircuitNumber, sFeederNumber, sLaneNumber]
                        all_action_items.append(store_action_items)

    log.debug('Concating all_feeder_items into df_feeder...')
    df_feeder = pd.DataFrame(all_feeder_items, columns=['PROGRAM_NAME', 'MACHINE', 'SECTION_NUMBER', 'TROLLEY_TYPE', 'FEEDER_NUMBER', 'FEEDER_TYPE', 'LANE_NUMBER', 'COMPONENT', 'SHAPE'])
    log.debug(f"\n{df_feeder.head(5).to_string(index=False)}")

    log.debug('Concating all_action_items into df_action...')
    df_action = pd.DataFrame(all_action_items, columns=['PROGRAM_NAME', 'MACHINE', 'SECTION_NUMBER', 'ROBOT_NUMBER', 'HEAD_NUMBER', 'DESIGNATOR', 'BOARD_NUMBER', 'FEEDER_NUMBER', 'LANE_NUMBER'])
    log.debug(f"\n{df_action.head(5).to_string(index=False)}")

    log.debug('Inner joining df_feeder into df_action...')
    df_feeder_action = df_action.merge(df_feeder, how='inner', left_on=['PROGRAM_NAME', 'MACHINE', 'SECTION_NUMBER', 'FEEDER_NUMBER', 'LANE_NUMBER'], right_on=['PROGRAM_NAME', 'MACHINE', 'SECTION_NUMBER', 'FEEDER_NUMBER', 'LANE_NUMBER'])
    df_feeder_action = df_feeder_action[['PROGRAM_NAME', 'MACHINE', 'COMPONENT', 'DESIGNATOR', 'BOARD_NUMBER', 'SHAPE', 'SECTION_NUMBER', 'FEEDER_NUMBER', 'LANE_NUMBER', 'ROBOT_NUMBER', 'HEAD_NUMBER', 'FEEDER_TYPE', 'TROLLEY_TYPE']]
    log.debug(f"\n{df_feeder_action.head(5).to_string(index=False)}")

    log.debug('Dropping duplicates...')
    df_program = df_feeder_action.drop_duplicates(subset=['PROGRAM_NAME', 'MACHINE', 'COMPONENT', 'DESIGNATOR','BOARD_NUMBER'], keep='last')
    log.debug(f"\n{df_program.head(5).to_string(index=False)}")

    log.debug('Sorting df_program...')
    df_program = df_program.sort_values(by=['PROGRAM_NAME', 'MACHINE', 'COMPONENT', 'DESIGNATOR', 'BOARD_NUMBER'])
    log.debug(f"\n{df_program.head(5).to_string(index=False)}")

    log.info('Writing df_program detail into SCRIPT_OUTPUT_PROGRAM.xlsx ...')
    df_program.to_excel(f"{path_main}\\SCRIPT_OUTPUT_PROGRAM.xlsx", sheet_name="DETAIL", index=False) 
    
    log.debug('Adding LOCATION column...')
    df_program['LOCATION'] = 'Board: ' + df_program['BOARD_NUMBER']  + ', Machine: ' + df_program['MACHINE'] + ', Section: ' + df_program['SECTION_NUMBER'] + ', Feeder: ' + df_program['FEEDER_NUMBER'] + ', Lane: ' + df_program['LANE_NUMBER'] + ', Robot: ' + df_program['ROBOT_NUMBER'] + ', Head: ' + df_program['HEAD_NUMBER'] + ' (' + df_program['FEEDER_TYPE'] + ', ' + df_program['TROLLEY_TYPE'] + ')'
    df_program = df_program[['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR', 'SHAPE', 'BOARD_NUMBER', 'LOCATION']]
    log.debug(f"\n{df_program.head(5).to_string(index=False)}")

    log.debug('Grouping location...')
    df_program = df_program.groupby(['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR', 'SHAPE']).aggregate({'BOARD_NUMBER': lambda x: ','.join(sorted(x)) , 'LOCATION': lambda x: '\n'.join(sorted(x))}).reset_index()
    log.debug(f"\n{df_program.head(5).to_string(index=False)}")

    log.info('Writing df_program summary into SCRIPT_OUTPUT_PROGRAM.xlsx ...')
    with pd.ExcelWriter(f"{path_main}\\SCRIPT_OUTPUT_PROGRAM.xlsx", mode='a', if_sheet_exists='replace') as f:
        df_program.to_excel(f, sheet_name="SUMMARY", index=False) 

    log.info(f"Total of {str(len(df_program))} rows detected in program files.")
    log.info('All selected input files are successfully loaded, proceeding with the checking algorithm...')

    log.info('Algorithm 1: Starting to calculate count of part number in PNP_PROGRAM...')

    log.debug('Dropping duplicates from df_program...')
    df_program_qty = df_program.drop_duplicates(subset=['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR'], keep='last')
    log.debug(f"\n{df_program_qty.head(5).to_string(index=False)}")

    log.debug('Grouping by designator...')
    df_program_qty = df_program_qty.groupby(['PROGRAM_NAME', 'COMPONENT'])['DESIGNATOR'].count().reset_index()
    log.debug(f"\n{df_program_qty.head(5).to_string(index=False)}")

    log.debug('Renaming designator count to program qty and convert into int...')
    df_program_qty = df_program_qty.rename(columns={'DESIGNATOR':'PROGRAM_QTY'})
    df_program_qty['PROGRAM_QTY'] = df_program_qty['PROGRAM_QTY'].astype(int)
    log.debug(f"\n{df_program_qty.head(5).to_string(index=False)}")

    log.info('Algorithm 1 completed: Count of part number is calculated in PNP_PROGRAM.')

    log.info('Algorihm 2: Starting to merge 590 and MCTO into master source table...')

    log.debug('Keeping only BOM and MCTO from df_590_MCTO and drop duplicates...')
    df_590_MCTO_PV = df_input[['BOM', 'MCTO', 'PV']]
    df_590_MCTO_PV.drop_duplicates(subset=['BOM', 'MCTO', 'PV'], keep='last', inplace=True)
    log.debug(f"\n{df_590_MCTO_PV.head(5).to_string(index=False)}")

    log.debug('Keeping only BOM, MCTO, PNP_PROGRAM_SIDE1 and PNP_PROGRAM_SIDE2 from df_input and drop duplicates...')    
    df_590_MCTO_PV_program = df_input[['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2']]
    df_590_MCTO_PV_program.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2'], keep='last', inplace=True)
    log.debug(f"\n{df_590_MCTO_PV_program.head(5).to_string(index=False)}")

    log.debug('Inner joining df_590_MCTO on BOM and drop duplicates...')
    df_590_all = df_590.merge(df_590_MCTO_PV, how='inner', left_on='BOM', right_on='BOM')
    df_590_all['GROUP'] = 'BOM_590' 
    df_590_all = df_590_all[['BOM', 'MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'GROUP']]
    df_590_all.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'GROUP'], keep='last', inplace=True)
    log.debug(f"\n{df_590_all.head(5).to_string(index=False)}")

    log.debug('Inner joining df_590_MCTO on MCTO and drop duplicates...')
    df_MCTO_all = df_MCTO.merge(df_590_MCTO_PV, how='inner', left_on=['MCTO', 'PV'], right_on=['MCTO', 'PV'])
    df_MCTO_all['GROUP'] = 'MCTO'
    df_MCTO_all = df_MCTO_all[['BOM', 'MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'GROUP']]
    df_MCTO_all.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'GROUP'], keep='last', inplace=True)
    log.debug(f"\n{df_MCTO_all.head(5).to_string(index=False)}")

    # Check if 590/MCTO is having any data
    if len(df_590_all) < 1 or len(df_MCTO_all) < 1:
        raise ConnectionAbortedError (f"df_590_all or df_MCTO_all is empty, force exiting application...")

    log.debug('Concating df_590_all into df_MCTO_all...')
    df_590_MCTO_all = pd.concat([df_590_all, df_MCTO_all], ignore_index=True)
    log.debug(f"\n{df_590_MCTO_all.head(5).to_string(index=False)}")

    log.info('Algorithm 2 completed: Merged 590 and MCTO into master source table.')

    log.info('Algorithm 3: Starting to convert Memory Parts into 520-XXX...')

    log.debug('Running logic to generate component with 520-XXX memory...')
    df_Material = df_590_MCTO_all
    df_Material['MemoryDesc'] = np.where(~df_Material['COMPONENT'].str.contains('-'), df_Material['COMPDESC'], np.NaN)
    df_Material['MemoryDesc_dash'] = df_Material['MemoryDesc'].str.split('-').str[0].str.strip()
    df_Material['MemoryDesc_last'] = df_Material['MemoryDesc_dash'].str[-3:]
    df_Material['Last1'] = df_Material['MemoryDesc_last'].apply(digit_to_nondigit, keep='First').fillna('').replace('nan', '', regex=True).astype(str)
    df_Material['Last2'] = df_Material['MemoryDesc_last'].apply(digit_to_nondigit, keep='Last').fillna('').replace('nan', '', regex=True).astype(str)
    df_Material['MemoryDesc_:'] = df_Material['MemoryDesc'].str.split(':').str[-2]
    df_Material['COMPONENT2'] = np.where(df_Material['COMPONENT'].str.contains('-'), df_Material['COMPONENT'], np.where(df_Material['COMPDESC'].str.startswith('MTC'), df_Material['MemoryDesc_:'], np.where(df_Material['Last2'] != '', '520-' + df_Material['Last2'], np.where(df_Material['COMPDESC'].str.startswith('MT2'), '520-' + df_Material['Last1'].str[-2:], '520-' + df_Material['Last1']))))
    df_Material = df_Material.rename(columns={'COMPONENT':'COMPONENT3', 'COMPONENT2':'COMPONENT'})
    log.debug(f"\n{df_Material.head(5).to_string(index=False)}")

    log.debug('Dropping duplicates and inner join df_590_MCTO_program...')
    df_Material = df_Material[['BOM', 'MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR']]
    df_Material.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR'], keep='last', inplace=True)
    df_Material = df_Material.merge(df_590_MCTO_PV_program, how='inner', left_on=['BOM', 'MCTO', 'PV'], right_on=['BOM', 'MCTO', 'PV'])
    df_Material = df_Material[['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR']]
    df_Material.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR'], keep='last', inplace=True)
    df_Material = df_Material.sort_values(by=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT'])
    log.debug(f"\n{df_Material.head(5).to_string(index=False)}")

    log.info('Algorithm 3 completed: Memory parts are converted into 520-XXX')

    log.info('Algorithm 4: Starting to split designator into rows...')
    df_Material_expanded = split_into_rows(df_Material, column="DESIGNATOR")
    log.debug(f"\n{df_Material_expanded.head(5).to_string(index=False)}")
    log.info('Algoritm 4 completed: Designators are splited into rows.')

    log.info('Algorithm 5: Starting to check quantity...')
    df_checker = df_Material

    log.debug('Converting quantity to int...')
    df_checker['QUANTITY'] = df_checker['QUANTITY'].astype(int)
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Counting comma in designator + 1 as REFDES_QTY...')
    df_checker['REFDES_QTY'] = df_checker['DESIGNATOR'].str.count(',').astype(int) + 1
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Left joining df_program for side 1...')
    df_checker = df_checker.merge(df_program_qty, how='left', left_on=['PNP_PROGRAM_SIDE1', 'COMPONENT'], right_on=['PROGRAM_NAME', 'COMPONENT']).drop('PROGRAM_NAME', axis=1)
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Converting PROGRAM_QTY into int with null as zero, rename to PQ1...')
    df_checker['PROGRAM_QTY'] = df_checker['PROGRAM_QTY'].fillna(0).astype(int)
    df_checker = df_checker.rename(columns={'PROGRAM_QTY':'PQ1'})
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Left joining df_program_qty for side 2')
    df_checker = df_checker.merge(df_program_qty, how='left', left_on=['PNP_PROGRAM_SIDE2', 'COMPONENT'], right_on=['PROGRAM_NAME', 'COMPONENT']).drop('PROGRAM_NAME', axis=1)
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")
    
    log.debug('Converting PROGRAM_QTY into int with null as zero, rename to PQ2...')
    df_checker['PROGRAM_QTY'] = df_checker['PROGRAM_QTY'].fillna(0).astype(int)
    df_checker = df_checker.rename(columns={'PROGRAM_QTY':'PQ2'})
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Summing PQ1 and PQ2 into PROGRAM_QTY as int...')
    df_checker['PROGRAM_QTY'] = (df_checker['PQ1'] + df_checker['PQ2']).astype(int)
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Generating SAP_QTY_TALLY and PROGRAM_QTY_TALLY...')
    df_checker['SAP_QTY_TALLY?'] = np.where((df_checker['QUANTITY'] == df_checker['REFDES_QTY']), 'Yes', 'No')
    df_checker['PROGRAM_QTY_TALLY?'] = np.where((df_checker['QUANTITY'] == df_checker['PROGRAM_QTY']), 'Yes', 'No')
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.info('Algorithm 5 completed: Quantity checked.')

    log.info('Algorithm 6: Starting to check part number and designator...')

    log.debug('Splitting designator into rows...')
    df_checker = split_into_rows(df_checker, column='DESIGNATOR')
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Left joining df_program with LOCATION_SIDE1...')
    df_checker = df_checker.merge(df_program[['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR', 'BOARD_NUMBER', 'LOCATION']], how='left', left_on=['PNP_PROGRAM_SIDE1', 'COMPONENT', 'DESIGNATOR'], right_on=['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR']).drop(['PROGRAM_NAME'], axis=1)
    df_checker = df_checker.rename(columns={'BOARD_NUMBER':'BOARD_SIDE1', 'LOCATION':'LOCATION_SIDE1'})
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Left joining df_program with LOCATION_SIDE2...')
    df_checker = df_checker.merge(df_program[['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR', 'BOARD_NUMBER', 'LOCATION']], how='left', left_on=['PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR'], right_on=['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR']).drop(['PROGRAM_NAME'], axis=1)
    df_checker = df_checker.rename(columns={'BOARD_NUMBER':'BOARD_SIDE2', 'LOCATION':'LOCATION_SIDE2'})
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Dropping duplicates...')
    df_checker.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?', 'BOARD_SIDE1', 'BOARD_SIDE2', 'LOCATION_SIDE1', 'LOCATION_SIDE2'], keep='last', inplace=True)
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.debug('Generating Checker Result...')
    df_checker['CHECKER'] = np.where((~df_checker['BOARD_SIDE1'].isnull()) & (~df_checker['BOARD_SIDE2'].isnull()), 'Something wrong, both side mounting the same designator', np.where((~df_checker['BOARD_SIDE1'].isnull()), ('Mount at Side 1 on Board ' + df_checker['BOARD_SIDE1']), np.where((~df_checker['BOARD_SIDE2'].isnull()), ('Mount at Side 2 on Board ' + df_checker['BOARD_SIDE2']), 'Not found'))) 
    df_checker['LOCATION'] = np.where((~df_checker['LOCATION_SIDE1'].isnull()) & (~df_checker['LOCATION_SIDE2'].isnull()), 'Something wrong, both side mounting the same designator', np.where((~df_checker['LOCATION_SIDE1'].isnull()), ('Mount at Side 1 on ' + df_checker['LOCATION_SIDE1'].str.replace('\n', '\nMount at Side 1 on ')), np.where((~df_checker['LOCATION_SIDE2'].isnull()), ('Mount at Side 2 on ' + df_checker['LOCATION_SIDE2'].str.replace('\n', '\nMount at Side 2 on ')), 'Not found'))) 
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")
    
    log.debug('Extrating designator num from end and split into letter/number...')
    df_checker['DESIGNATOR_letter'] = df_checker['DESIGNATOR'].apply(extract_num_from_end, keep='letter').replace('', np.NaN, regex=True).astype(str)
    df_checker['DESIGNATOR_number'] = df_checker['DESIGNATOR'].apply(extract_num_from_end, keep='number').replace('', 0, regex=True).astype(int)
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")
    
    log.debug('Sorting designator...')
    df_checker = df_checker.sort_values(by=['BOM', 'MCTO', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR_letter', 'DESIGNATOR_number'])
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")
    
    log.debug('Grouping by designator and checker...')
    df_checker = df_checker.groupby(['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?']).aggregate({'DESIGNATOR': lambda x: ','.join(x), 'CHECKER': lambda x: '\n'.join(sorted(x)), 'LOCATION': lambda x: '\n'.join(sorted(x))}).reset_index()
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")
    
    log.debug('Removing string duplicates on designator and checker...')
    df_checker = df_checker[['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?', 'CHECKER', 'LOCATION']]
    df_checker['DESIGNATOR'] = df_checker['DESIGNATOR'].apply(string_remove_duplicate, delimiter=',')
    df_checker['CHECKER'] = df_checker['CHECKER'].apply(string_remove_duplicate, delimiter='\n')
    df_checker['LOCATION'] = df_checker['LOCATION'].apply(string_remove_duplicate, delimiter='\n')
    df_checker = split_into_rows(df_checker, column='LOCATION', sep='\n')
    df_checker = df_checker.groupby(['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?', 'CHECKER']).aggregate({'LOCATION': lambda x: '\n'.join(sorted(x))}).reset_index()
    log.debug(f"\n{df_checker.head(5).to_string(index=False)}")

    log.info('Algorithm 6 completed: Part number and designator checked.')

    log.info('Algorithm 7: Starting to check extra programmed parts...')
    df_checker_extra = df_program[['PROGRAM_NAME', 'COMPONENT', 'DESIGNATOR', 'BOARD_NUMBER', 'LOCATION']]

    log.debug('Left joining df_590_MCTO_program for side 1...')
    df_checker_extra = df_checker_extra.merge(df_590_MCTO_PV_program, how='left', left_on=['PROGRAM_NAME'], right_on=['PNP_PROGRAM_SIDE1']).drop(['PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2'], axis=1)
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Left joining df_590_MCTO_program for side 2...')
    df_checker_extra = df_checker_extra.merge(df_590_MCTO_PV_program, how='left', left_on=['PROGRAM_NAME'], right_on=['PNP_PROGRAM_SIDE2']).drop(['PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2'], axis=1)
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Combining BOM_x and BOM_y ...')
    df_checker_extra['BOM'] = df_checker_extra['BOM_x'].fillna(df_checker_extra['BOM_y'])
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Combining MCTO_x and MCTO_y ...')
    df_checker_extra['MCTO'] = df_checker_extra['MCTO_x'].fillna(df_checker_extra['MCTO_y'])
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Combining PV_x and PV_y ...')
    df_checker_extra['PV'] = df_checker_extra['PV_x'].fillna(df_checker_extra['PV_y'])
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Left joining df_590_MCTO_program on BOM, MCTO and PV...')
    df_checker_extra = df_checker_extra.merge(df_590_MCTO_PV_program, how='left', left_on=['BOM', 'MCTO', 'PV'], right_on=['BOM', 'MCTO', 'PV'])
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Creating NOARD_SIDE1 and LOCATION_SIDE1...')
    df_checker_extra['BOARD_SIDE1'] = np.where((df_checker_extra['PROGRAM_NAME'] == df_checker_extra['PNP_PROGRAM_SIDE1']), df_checker_extra['BOARD_NUMBER'], np.NaN)
    df_checker_extra['LOCATION_SIDE1'] = np.where((df_checker_extra['PROGRAM_NAME'] == df_checker_extra['PNP_PROGRAM_SIDE1']), df_checker_extra['LOCATION'], np.NaN)
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Creating BOARD_SIDE1 and LOCATION_SIDE2...')
    df_checker_extra['BOARD_SIDE2'] = np.where((df_checker_extra['PROGRAM_NAME'] == df_checker_extra['PNP_PROGRAM_SIDE1']), df_checker_extra['BOARD_NUMBER'], np.NaN)
    df_checker_extra['LOCATION_SIDE2'] = np.where((df_checker_extra['PROGRAM_NAME'] == df_checker_extra['PNP_PROGRAM_SIDE2']), df_checker_extra['LOCATION'], np.NaN)
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    log.debug('Removing rows with null BOM...')
    df_checker_extra = df_checker_extra[['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR', 'BOARD_SIDE1', 'BOARD_SIDE2', 'LOCATION_SIDE1', 'LOCATION_SIDE2']]
    df_checker_extra = df_checker_extra[(df_checker_extra['BOM'].notnull())]
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
    
    log.debug('Left joining df_Material_expanded with indicator...')
    df_checker_extra = df_checker_extra.merge(df_Material_expanded[['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR']], how='left', left_on=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR'], right_on=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR'], indicator=True)
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
    
    log.debug('Keeping left_only merge...')
    df_checker_extra = df_checker_extra[(df_checker_extra._merge=='left_only')].drop('_merge', axis=1)
    log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

    if df_checker_extra.empty:
        log.debug('df_checker_extra is empty, assigning df_checker to df_checker_all...')
        df_checker_all = df_checker
        log.info('No extra parts programmed.')
    else:
        log.debug('There is extra parts programmed.')

        log.debug('Dropping duplicates...')
        df_checker_extra.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR', 'BOARD_SIDE1', 'BOARD_SIDE2', 'LOCATION_SIDE1', 'LOCATION_SIDE2'], keep='last', inplace=True)
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
        
        log.debug('Setting QUANTITY and REFDES_QTY as 0...')
        df_checker_extra['QUANTITY'] = 0
        df_checker_extra['REFDES_QTY'] = 0
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

        log.debug('Left joining df_program_qty for side 1 as PQ1...')
        df_checker_extra = df_checker_extra.merge(df_program_qty, how='left', left_on=['PNP_PROGRAM_SIDE1', 'COMPONENT'], right_on=['PROGRAM_NAME', 'COMPONENT']).drop('PROGRAM_NAME', axis=1)
        df_checker_extra['PROGRAM_QTY'] = df_checker_extra['PROGRAM_QTY'].fillna(0).astype(int)
        df_checker_extra = df_checker_extra.rename(columns={'PROGRAM_QTY':'PQ1'})
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

        log.debug('Left joining df_program_qty for side 2 as PQ2...')
        df_checker_extra = df_checker_extra.merge(df_program_qty, how='left', left_on=['PNP_PROGRAM_SIDE2', 'COMPONENT'], right_on=['PROGRAM_NAME', 'COMPONENT']).drop('PROGRAM_NAME', axis=1)
        df_checker_extra['PROGRAM_QTY'] = df_checker_extra['PROGRAM_QTY'].fillna(0).astype(int)
        df_checker_extra = df_checker_extra.rename(columns={'PROGRAM_QTY':'PQ2'})
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

        log.debug('Summing PQ1 and PQ2 as PROGRAM_QTY...')
        df_checker_extra['PROGRAM_QTY'] = (df_checker_extra['PQ1'] + df_checker_extra['PQ2']).astype(int)
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

        log.debug('Generating SAP_QTY_TALLY and PROGRAM_QTY_TALLY...')
        df_checker_extra['SAP_QTY_TALLY?'] = np.where((df_checker_extra['QUANTITY'] == df_checker_extra['REFDES_QTY']), 'Yes', 'No')
        df_checker_extra['PROGRAM_QTY_TALLY?'] = np.where((df_checker_extra['QUANTITY'] == df_checker_extra['PROGRAM_QTY']), 'Yes', 'No')
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

        log.debug('Dropping duplicates...')
        df_checker_extra.drop_duplicates(subset=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'QUANTITY', 'DESIGNATOR', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?', 'BOARD_SIDE1', 'BOARD_SIDE2', 'LOCATION_SIDE1', 'LOCATION_SIDE2'], keep='last', inplace=True)
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
        
        log.debug('Generating Checker Result...')
        df_checker_extra['CHECKER'] = np.where((~df_checker_extra['BOARD_SIDE1'].isnull()) & (~df_checker_extra['BOARD_SIDE2'].isnull()), 'Something wrong, both side extra mounting the same designator', np.where((~df_checker_extra['BOARD_SIDE1'].isnull()), ('Extra Mount at Side 1 on Board ' + df_checker_extra['BOARD_SIDE1']), np.where((~df_checker_extra['BOARD_SIDE2'].isnull()), ('Extra Mount at Side 2 on Board ' + df_checker_extra['BOARD_SIDE2']), 'Not found'))) 
        df_checker_extra['LOCATION'] = np.where((~df_checker_extra['LOCATION_SIDE1'].isnull()) & (~df_checker_extra['LOCATION_SIDE2'].isnull()), 'Something wrong, both side mounting the same designator', np.where((~df_checker_extra['LOCATION_SIDE1'].isnull()), ('Extra Mount at Side 1 on ' + df_checker_extra['LOCATION_SIDE1'].str.replace('\n', '\nExtra Mount at Side 1 on ')), np.where((~df_checker_extra['LOCATION_SIDE2'].isnull()), ('Extra Mount at Side 2 on ' + df_checker_extra['LOCATION_SIDE2'].str.replace('\n', '\nExtra Mount at Side 2 on ')), 'Not found'))) 
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
        
        log.debug('Extrating designator num from end and split into letter/number...')
        df_checker_extra['DESIGNATOR_letter'] = df_checker_extra['DESIGNATOR'].apply(extract_num_from_end, keep='letter').replace('', np.NaN, regex=True).astype(str)
        df_checker_extra['DESIGNATOR_number'] = df_checker_extra['DESIGNATOR'].apply(extract_num_from_end, keep='number').replace('', 0, regex=True).astype(int)
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
        
        log.debug('Sorting designator...')
        df_checker_extra = df_checker_extra.sort_values(by=['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'DESIGNATOR_letter', 'DESIGNATOR_number'])
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
        
        log.debug('Grouping by designator and checker...')
        df_checker_extra = df_checker_extra.groupby(['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'QUANTITY', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?']).aggregate({'DESIGNATOR': lambda x: ','.join(x), 'CHECKER': lambda x: '\n'.join(sorted(x)), 'LOCATION': lambda x: '\n'.join(sorted(x))}).reset_index()
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")
        
        log.debug('Adding COMPDESC as null...')
        df_checker_extra['COMPDESC'] = np.NaN
        df_checker_extra = df_checker_extra[['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?', 'CHECKER', 'LOCATION']]
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

        log.debug('Removing string duplicates on designator and checker...')
        df_checker_extra['DESIGNATOR'] = df_checker_extra['DESIGNATOR'].apply(string_remove_duplicate, delimiter=',')
        df_checker_extra['CHECKER'] = df_checker_extra['CHECKER'].apply(string_remove_duplicate, delimiter='\n')
        df_checker_extra['LOCATION'] = df_checker_extra['LOCATION'].apply(string_remove_duplicate, delimiter='\n')
        df_checker_extra = split_into_rows(df_checker_extra, column='LOCATION', sep='\n')
        df_checker_extra = df_checker_extra.groupby(['BOM', 'MCTO', 'PV', 'PNP_PROGRAM_SIDE1', 'PNP_PROGRAM_SIDE2', 'COMPONENT', 'COMPDESC', 'QUANTITY', 'DESIGNATOR', 'REFDES_QTY', 'PROGRAM_QTY', 'SAP_QTY_TALLY?', 'PROGRAM_QTY_TALLY?', 'CHECKER']).aggregate({'LOCATION': lambda x: '\n'.join(sorted(x))}).reset_index()
        log.debug(f"\n{df_checker_extra.head(5).to_string(index=False)}")

        log.info('Algorithm 7 completed: Extra programmed part checked.')

        log.debug('Concating df_checker and df_checker_extra...')
        df_checker_all = pd.concat([df_checker, df_checker_extra], ignore_index=True)
        log.debug(f"\n{df_checker_all.head(5).to_string(index=False)}")

    log.info('All checking algorithm has been completed.')
    log.info('Writing Checker output table into SCRIPT_OUTPUT.xlsx...')
    df_checker_all.to_excel(path_main + '\\SCRIPT_OUTPUT.xlsx', sheet_name='OUTPUT', index=False)

    log.info('Successfully completed without any errors!!!')
    log.info('Closing application...' + '\n')
    time.sleep(5)

    return


if __name__ == '__main__':
    log, path_main, path_590, path_MCTO, path_program, path_checker, input_columns, output_columns = init()
    try:
        main(log, path_main, path_590, path_MCTO, path_program, path_checker, input_columns, output_columns)

    except ConnectionAbortedError as e:
        log.error(f"{str(e)}")
        df_checker_all = pd.DataFrame(columns=[output_columns])
        df_checker_all.to_excel(f"{path_main}\\SCRIPT_OUTPUT.xlsx", sheet_name='OUTPUT', index=False)
        time.sleep(5)
        sys.exit(0)

    except Exception as e:
        log.critical('Force exiting application...')
        log.exception(f"Unexpected Error: {str(e)}")
        df_checker_all = pd.DataFrame(columns=[output_columns])
        df_checker_all.to_excel(f"{path_main}\\SCRIPT_OUTPUT.xlsx", sheet_name='OUTPUT', index=False)
        time.sleep(5)