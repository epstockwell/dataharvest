import numpy as np
import os
import multiprocessing
import glob

from pdfminer.high_level import extract_text
from process import process_numbers, parse_text


if __name__ == "__main__":
    # Open the Station List
    base_dir = r'I-P climate data_2017'
    out_txt = extract_text(os.path.join(base_dir, 'StnList_p.pdf'))

    # Extract all nonempty cells from the PDF
    text_array = out_txt.split('\n')
    text_array = [element for element in text_array if element != '']

    # Extract all 6-digit numbers
    num_array = []
    for element in text_array:
        if len(element) == 6:
            try:
                num = int(element)
                num_array.append(element)
            except ValueError:
                pass
    num_array = np.array([int(x) for x in num_array])

    # Reduce array to fit between first and last stations
    first_station = 712850
    last_station = 719640
    num_array = num_array[np.where(num_array == first_station)[0].max():np.where(num_array == last_station)[0].max() + 1]


    # Split array into separate workers and process
    processes = []
    for i in range(8):
        p = multiprocessing.Process(target=process_numbers, args=(i, num_array[i::8], ))
        p.start()
        processes.append(p)

    for process in processes:
        process.join()


    # Write master CSV
    data = []

    # Obtain existing CSV Data
    for file in glob.glob('*.csv'):
        if 'result' not in file:
            reader = csv.reader(open(file, 'r'))
            for row in reader:
                if len(row) > 0:
                    data.append(row)
    #     os.remove(file)

    # Split the location string
    new_data = []
    for row in data:
        try:
            new_row = [row[0], row[1].split(', ')[1], row[1].split(', ')[0], row[2], row[3], row[4], row[5]]
        except:
            new_row = row
        new_data.append(new_row)
        
    # Save to master CSV
    with open('result.csv', 'w', newline='') as out_file:
        writer = csv.writer(out_file)
        writer.writerow(['USAF Code', 'State', 'Location', 'Dew Point 1%', 'Moist Count Dry Bulb 1%', 'Dew Point 2%', 'Moist Count Dry Bulb 2%'])
        for row in new_data:
            writer.writerow(row)
