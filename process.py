import csv
from pdfminer.high_level import extract_text
import numpy as np
import os


def is_name(text):
    """
    Accepts a string and returns True if it maches name conventions
    :param text: The text string to check
    :returns: bool
    """

    # Make sure the text is all uppercase and has a space
    res = text == text.upper() and ' ' in text

    # Make sure the text has no disallowed characters    
    disallowed = ['WMO#', 'WBAN:', '5% DB', '5% WB']
    res = res and not any([x in text for x in disallowed])
    
    return res


def parse_text(in_text):
    """
    Extracts valuable info from a string of text from a PDF file in ASHRAE format
    :param in_text: the string of text from pdfminer.high_level.extract_text
    :returns: tuple of (station name, dew point 1%, moist count dry bulb 1%, dew point 2%, moist count dry bulb 2%)
    """
    try:
        dp1_idx = in_text[in_text.index('DP\n( d )')+1:].index('DP\n( d )')
        dp1 = in_text[in_text.index('DP\n( d )')+1:][dp1_idx+9:dp1_idx+13]
    except:
        dp1 = 'N/A'
    
    try:
        dp2_idx = in_text[in_text.index('DP\n( g )')+1:].index('DP\n( g )')
        dp2 = in_text[in_text.index('DP\n( g )')+1:][dp2_idx+9:dp2_idx+13]
    except:
        dp2 = 'N/A'
    
    try:
        mcdb1_idx = in_text[in_text.index('MCDB\n( f )')+1:].index('MCDB\n( f )')
        mcdb1 = in_text[in_text.index('MCDB\n( f )')+1:][mcdb1_idx+11:mcdb1_idx+15]
    except:
        mcdb1 = 'N/A'
    
    try:
        mcdb2_idx = in_text[in_text.index('MCDB\n( i )')+1:].index('MCDB\n( i )')
        mcdb2 = in_text[in_text.index('MCDB\n( i )')+1:][mcdb2_idx+11:mcdb2_idx+15]
    except:
        mcdb2 = 'N/A'
    
    try:
        name = [x for x in in_text.split('\n') if is_name(x)]

        name = name[0] if len(name) > 0 else 'NAME NOT AVAILABLE'
    except:
        name = 'NAME NOT AVAILABLE'
    
    return name, dp1, mcdb1, dp2, mcdb2


def process_numbers(worker_id, numbers_in):
    """
    Multiprocessing worker to extract ASHRAE data
    :param worker_id: The worker id to save files to
    :param numbers_in: The USAF station numbers to extract info for
    :returns: None
    """
    result_arr = []
    base_dir = r'C:\Users\njsto\Desktop\ts\Zipped Climate Data and Chapter 14\I-P climate data_2017'
    
    # Iterate over given numbers
    for number in numbers_in:
        try:
            file_text = extract_text(os.path.join(base_dir, f'{number}_p.pdf'))
            parsed_info = parse_text(file_text)
            result_arr.append([number, *parsed_info])
        except Exception as e:
            print(e)
            print(number)
            print('\n')
    
    # Write all results to file
    with open(f'save{worker_id}.csv', 'w') as csv_file:
        writer = csv.writer(csv_file)
        for row in result_arr:
            writer.writerow(row)