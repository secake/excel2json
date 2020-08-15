from openpyxl import load_workbook
import sys, getopt
import json


def main(argv):
    input_file_path = ''
    output_file_path = ''
    sheet=''
    header_line = 1
    try:
        
        json_fd = open('./config.json', mode='r')
        config = json_fd.read()
        config = json.loads(config)
       
        input_file_path = config['input_file_path']
        output_file_path = config['output_file_path']
        sheet = config['sheet']
        header_line = config['header_line']
        wb = load_workbook(input_file_path)
        if sheet not in wb:
            print("error:","sheet:",sheet," is not found in file:",input_file_path)
        ws = wb[sheet]
        json_str = excel_2_json(ws, header_line)
  
        fd = open(output_file_path, mode='w')
        fd.write(json_str)
        fd.close()
    except Exception as e:
        print("error:",repr(e))
        sys.exit(2)


def excel_2_json(ws, header_line):
    json_list = list()
    

    header = list()
    row_num = -1
    for row in ws.values:
        row_num += 1
        col_num = -1
        json_item = dict()

        if row_num < int(header_line)-1:
            continue        
        for value in row:
            col_num += 1
            if row_num == (header_line)-1:
                header.append(value)
                continue
            
            json_item[header[col_num]] = value
        
        is_ok = False
        for k in json_item.keys():
            if json_item[k]:
                is_ok = True
        if is_ok:
            json_list.append(json_item)

    json_str = json.dumps(json_list, indent=2)
    return json_str





























if __name__ == "__main__":
    main(sys.argv[1:])
