import os
import json
import logging
from logging.handlers import TimedRotatingFileHandler
import re
from datetime import datetime, date, timedelta
import itertools
from difflib import SequenceMatcher
import zipfile

import requests
from bs4 import BeautifulSoup
import dpath.util
from PIL import Image, ImageFont, ImageDraw

from dotenv import load_dotenv
load_dotenv()

mntu_password = os.getenv('mntu_password')
mntu_base_url = os.getenv('mntu_base_url') 

logger = logging.getLogger('log')
clr_dark_green = (47, 96, 85)
clr_light_green = (92,141,135)
clr_white = (255,255,255)
clr_black = (0, 0, 0)
clr_gray = (191, 191, 191)
clr_red = (255, 0, 0)
clr_blue = (0, 255, 0)
font_18 = ImageFont.truetype("./res/arial.ttf", 18)
font_21 = ImageFont.truetype("./res/arial.ttf", 21)
padding_px = 20 #sum of left \ right content padding
row_height=25 # px


def init_project_structure():
    """
    Init project directory structure
    """
    if not os.path.exists("./logs"):        os.makedirs("./logs")
    if not os.path.exists("./res"):         os.makedirs("./res")
    if not os.path.exists("./res/docx"):    os.makedirs("./res/docx")
    if not os.path.exists("./res/pics"):    os.makedirs("./res/pics")
    if not os.path.exists('./res/json'):    os.makedirs("./res/json")
    if not os.path.exists("./res/users_db"):   
        with open('./res/users_db', 'w', encoding="utf-8") as f: 
            json.dump({}, f, indent=1, ensure_ascii=False)
        f.close()


def log_options_init():    
    """
    Initializes the logging options
    """
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter(fmt='%(asctime)s %(levelname)s %(message)s',datefmt='%m-%d-%y %H:%M:%S')
    fh = TimedRotatingFileHandler('./logs/log', when='midnight')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    logger.info('> log options initialized')


def grab_docx_files():
    """
    Downloads docx files going through site via DFS and store files to `./res/docx` 
    Returns:
        dict: dictionary containing the paths of `.docx` files found on the website.
    """
    json_data={}
    logger.info('call grab_docx_files()')
    mntu_start_url = mntu_base_url+'index.php?p=30&id_f=114'
    s = requests.Session() 
    s.get(mntu_start_url) # GET request to site
    s.post(mntu_start_url, data={'password_lib': mntu_password}) # login to site.

    def dfs(node, path_on_site="", path_in_file=""):    #favtutor.com/blogs/depth-first-search-python
            #print('DFS: '+path_on_site+'\n'+path_in_file) #debug
            trs = BeautifulSoup(s.get(node).content , "html.parser").find('table', {"class": "forumline"}).find_all('tr')[2:]
            
            for tr in trs:
                file_name = tr.select_one('a').contents[0]
                file_href = tr.select_one('a')['href']
                temp_path_in_file=path_in_file+'/'+file_name  # header dir name
                
                print('['+file_name+']')
                if ".docx" in file_name:
                    dpath.util.new(json_data, temp_path_in_file, 'null' ) 
                    '''
                    temp = s.get(app.mntu_base_url + file_href)
                    with open('./res/docx/'+file_name+'.docx', 'wb') as output:
                        output.write(temp.content)
                    output.close()
                    '''
                elif ".doc" in file_name:
                    pass
                else:
                    if "викладачі" not in file_name.lower():
                        temp_path_on_site=path_on_site+'/'+file_name
                        dpath.util.new(json_data, temp_path_in_file, {})
                        dfs( mntu_base_url+file_href, temp_path_on_site, temp_path_in_file )        
    dfs(mntu_start_url)
    return json_data


def docxs_handler():
    """
    Docx driver handler, which extracts data / parse it on XML level and stack in json.
    """
    logger.info('call docxs_handler')
    dir_content = os.listdir('./res/docx')
    docx_files = [x for x in dir_content if not x.startswith('~')] #exclude temp / open files
    for file_name in docx_files:
        json_data={}
        print("=> "+file_name) #keypoint debug
        zip = zipfile.ZipFile('./res/docx/'+file_name, 'r')
        file = zip.open('word/document.xml', "r").read() #bytes like obj (without utf)
        
        file_name = file_name.split('.')[0]
        path_to_json="./res/json/"+file_name+".json"
        
        table_for_this_week = get_tabel(file)
        if(table_for_this_week != 0):
            if not os.path.exists(path_to_json):       
                open(path_to_json, 'w').close()

            json_data = fetch_schedule_data(table_for_this_week, file_name)
        
            with open(path_to_json, 'w', encoding="utf-8") as f:
                json.dump(json_data, f, indent=1, ensure_ascii=False)
            f.close()    
            
            #input("Press Enter to continue...") #dubug
            json_to_pic(json_data, file_name)


def build_canvas(json_data):
    """
    Creates blank canvas with dimensions calculated based `json_data` structure.
    
    It computes:
        - number of rows (total_time_rows)
        - maximum 1st column width (max_time_width),
        - maximum width for other groups/columns (max_column_width)
    Args:
        json_data (dict): data representing timetable to visualize
    
    Returns:
        list: list canvas obj and dimensions [img, total_time_rows, max_time_width, *max_column_width].
    """
    img = Image.new("RGB", (0,0), clr_dark_green)
    draw = ImageDraw.Draw(img)

    total_time_rows = 0     #rows num
    max_time_width = 0      #max width column with time (1st column)
    max_column_width = []   #max width of other columns
    for column in json_data:
        total_time_rows = 0 
        max_subject_width = 0  #max len
        for day in list(json_data[column])[1:]:
            total_time_rows = total_time_rows + len(json_data[column][day])
            for time in json_data[column][day]:
                text_box = draw.textbbox((0, 0), time, font=font_18) # calc time size (x1,y1,x2,y2)
                if(text_box[2]>max_time_width):
                    max_time_width = text_box[2]
                for subject in json_data[column][day][time]:
                    text_box = draw.textbbox((0, 0), subject, font=font_18) # calc subject size (x1,y1,x2,y2)
                    if(text_box[2]>max_subject_width):
                        max_subject_width = text_box[2]
        max_column_width.append(max_subject_width+padding_px)
    max_time_width = max_time_width + padding_px

    img = Image.new(img.mode, (img.size[0]+ row_height*(total_time_rows+2), img.size[1]), clr_dark_green)
    img = Image.new(img.mode, (img.size[0], img.size[1]+row_height+max_time_width+sum(max_column_width)), clr_dark_green) # Day_name_labels + time_w + columns
    return [img, total_time_rows, max_time_width, *max_column_width]


def json_to_pic(json_data, file_name):
    """
    Converts timetable data (in `json_data`) into image (PNG format).

    Args:
        json_data (dict): Data structure containing the timetable to be rendered.
        file_name (str): Name of the output image file.
    
    Returns:
        None
    """
    cancas_and_dimensions = build_canvas(json_data)
    img = cancas_and_dimensions[0]
    max_time_width = cancas_and_dimensions[2]
    column_width_px =  cancas_and_dimensions[len(cancas_and_dimensions)-len(json_data):]
    #print(column_width_px) #keypoint debug
    
    #in block below draw the days of the week, rotate image.
    draw = ImageDraw.Draw(img)
    for column in json_data:
        axis_x_pos=0
        for day in reversed(list(json_data[column])[1:]):
            total_rows_size_in_day = len(json_data[column][day])*row_height
            text_box = draw.textbbox((0, 0), day, font=font_18) # get size (x1,y1,x2,y2)
            draw.text(( axis_x_pos + (total_rows_size_in_day-text_box[2])/2, (row_height-text_box[3])/2), day, font=font_18,  fill=clr_white) # print days of week

            axis_x_pos=axis_x_pos+total_rows_size_in_day
            #draw.line((axis_x_pos, 0, axis_x_pos, img.size[1]), fill=clr_red) #keypoint debug. Line between day labels
        break
    img = img.rotate(90, Image.NEAREST, expand = 1) #rotate img

    dates = get_dates()
    start_week = dates[1]
    end_week = dates[2]
    header_rozklad_date_range_text="Розклад ("+str(start_week).split(' ')[0].replace('-','.')+" - "+ str(end_week).split(' ')[0].replace('-','.')+")"   #text like - "Розклад (2022.04.11 - 2022.04.17)"
    #print(header_rozklad_date_range_text) #keypoint debug
    text_box = draw.textbbox((0, 0), header_rozklad_date_range_text, font=font_21) 
    draw = ImageDraw.Draw(img)
    draw.text(((img.size[0]-text_box[2])/2, (row_height-text_box[3])/2), header_rozklad_date_range_text, font=font_21,  fill=clr_white) 
    draw.line((row_height, row_height, img.size[0], row_height), fill=clr_black) 
    draw.line((row_height, 0, row_height, img.size[1]), fill=clr_black) #vertical line before time

    text_box = draw.textbbox((0, 0), "ЧАС", font=font_21) 
    draw.text((row_height+(max_time_width-text_box[2])/2, row_height-2+(row_height-text_box[3])/2), "ЧАС", font=font_21,  fill=clr_white)

    #block which prints time & horizontal lines
    for column in json_data:
        draw.rectangle(((row_height+1, row_height*2), (row_height+ max_time_width, img.size[1])), fill=clr_light_green) #time column background
        draw.rectangle(((row_height+max_time_width+1, row_height*2+1), (img.size[0], img.size[1])), fill=clr_white) # main area white background
        axis_y_pos=row_height*2
        for day in list(json_data[column])[1:]:
            for time_id, time in enumerate(json_data[column][day]):
                text_box = draw.textbbox((0, 0), time, font=font_18)  
                draw.text((row_height + (max_time_width-text_box[2])/2, axis_y_pos-1 + (row_height-text_box[3])/2 ), time, font=font_18,  fill=clr_black) 
                
                if(time_id==0):
                    draw.line((0, axis_y_pos, img.size[0], axis_y_pos), fill=clr_black) # black line dividing the days of the week
                else:
                    draw.line((row_height+1, axis_y_pos, img.size[0], axis_y_pos), fill=clr_gray) # silver lines between each line.
                axis_y_pos = axis_y_pos + row_height
        break

    axis_x_pos=row_height+max_time_width
    draw.rectangle(((axis_x_pos, row_height+1), (img.size[0], row_height*2-1)), fill=clr_light_green)
    for column_id, column in enumerate(json_data):
        axis_y_pos = row_height*2
        text_box = draw.textbbox((0, 0), column, font=font_18)  

        draw.text((axis_x_pos + (column_width_px[column_id]-text_box[2])/2, row_height+(row_height-text_box[3])/2), column, font=font_18,  fill=clr_black) 
        draw.line((axis_x_pos, row_height, axis_x_pos, img.size[1]), fill=clr_black) #vertical line before each column of the group
        
        for day in list(json_data[column])[1:]:
            for time_id, time in enumerate(json_data[column][day]):
                priority_subject = ""
                priority_color = ""
                for subject_id, subject in enumerate(json_data[column][day][time]): #priotities BLUE>BLACK>RED
                    #subject_color = ImageColor.getcolor(json_data[column][day][time][subject]["state"], "RGB") #hex->rgb
                    if("#0070C0" == json_data[column][day][time][subject]["state"]):
                        priority_subject=subject
                        priority_color = "#0070C0"
                    elif("#000000" == json_data[column][day][time][subject]["state"] and priority_color != "#0070C0"):
                        priority_subject=subject
                        priority_color = "#000000"
                    elif("#FF0000" == json_data[column][day][time][subject]["state"]) and priority_color != "#0070C0" and priority_color != "#000000":
                        priority_subject=subject
                        priority_color = "#FF0000"
                priority_subject
                if(len(priority_subject)):
                    text_box = draw.textbbox((0, 0), priority_subject, font=font_18)
                    draw.text((axis_x_pos + (column_width_px[column_id]-text_box[2])/2, axis_y_pos+(row_height-text_box[3])/2), priority_subject, font=font_18,  fill=priority_color) 
                    if(priority_color=="#FF0000"):
                        draw.line((axis_x_pos + (column_width_px[column_id]-text_box[2])/2,  axis_y_pos + row_height/2, axis_x_pos + (column_width_px[column_id]-text_box[2])/2 + text_box[2], axis_y_pos + row_height/2), width=2, fill=clr_red) #vertical line before each column of "group"
                axis_y_pos = axis_y_pos + row_height
        axis_x_pos = axis_x_pos + column_width_px[column_id]
    img.save("./res/pics/"+file_name+".png")


def fetch_schedule_data(table, file_name):
    """
    Converts schedule table (partial XML of docx content) into structured JSON format.
    
    Arguments:
    - table (str): partial XML representing timetable.
    - file_name (str): The name of file where schedule data is being saved

    Returns:
    - json_data (dict): A nested dictionary representing the schedule.

    """
    column_headers = get_column_headers(table, file_name)
    #print(column_headers)
    json_data = init_json(file_name, column_headers)

    day=""
    time=""
    rows = re.findall(r'<w:tr .*?>(.+?)</w:tr>', table)
    for row in rows[1:]: #row id from 0... \ [1:] - coz 1st row - headers
        cells = re.findall(r'<w:tc>(.+?)</w:tc>', row)
        for cell_id, cell in enumerate(cells):
            if(cell_id > 2+len(column_headers)): #day+time+N columns
                break                            # break if out of expercted len
            else:
                cell_text=""
                cell_paragraphs = re.findall(r'<w:p .*?>(.+?)</w:p>', cell)
                for cell_p in cell_paragraphs:
                    p_text = re.findall(r'(?:<w:t>(.+?)</w:t>)|(?:<w:t .*?>(.+?)</w:t>)', cell_p)
                    p_text = ''.join(list(itertools.chain(*p_text)))
                    cell_text=cell_text+p_text+" "

                    if(cell_id==0):
                        if(p_text.isupper()):
                            day=p_text
                            for column in column_headers:
                                json_data[column][day]={}
                    
                    elif(cell_id==1):
                        time=p_text
                        for column in column_headers:
                            json_data[column][day][time]={}
                    
                    elif(cell_id < 2+len(column_headers) and cell_id>1):
                        p_colors = re.findall(r'(?:<w:color w:val="(.+?)"/>)', cell_p)
                        subject =str_cleaner([p_text])[0]
                        #print( time +"|"+ subject) # keypoint debug
                        if(len(subject) and subject[0].isalpha()):
                            subject = shorten_text(subject)
                            
                            #print(subject)
                            json_data[column_headers[cell_id-2]][day][time][subject]={}
                            if("л." in p_text):
                                json_data[column_headers[cell_id-2]][day][time][subject]["type"]="lecture"
                            elif("пр." in p_text):
                                json_data[column_headers[cell_id-2]][day][time][subject]["type"]="practice"
                            else:
                                json_data[column_headers[cell_id-2]][day][time][subject]["type"]="?"
                            if("FF0000" in p_colors):
                                json_data[column_headers[cell_id-2]][day][time][subject]["state"]="#FF0000"
                            elif("0070C0" in p_colors):
                                json_data[column_headers[cell_id-2]][day][time][subject]["state"]="#0070C0"
                            else:
                                json_data[column_headers[cell_id-2]][day][time][subject]["state"]="#000000"       
                    #print("["+p_text+"]")   
    return json_data


def init_json(file_name, column_headers):
    """
    Initializes JSON structure for storing schedule data, based on column headers and the given file name.
    """
    json_data={}
    known_columns=[]
    for column_name in column_headers:
        complete_group_names=[]
        known_columns.append(column_name)
        json_data[column_name]={}
        for temp in column_name.split(','):
            #print("["+temp+"]") #keypoint debug
            if(re.match("^[А-ЩЬЮЯҐЄІЇа-щьюяґєії]{1,}-\d{1,}$", temp)):
                complete_group_names.append(temp)
            else:
                if(re.match("^[А-ЩЬЮЯҐЄІЇа-щьюяґєії]{1,}", temp)):
                    matches = re.findall("\d{1,}", file_name )
                    complete_group_names.append(temp+"-"+matches[0])
                elif(re.match("\d{1,}", temp)):
                    matches = re.findall("[А-ЩЬЮЯҐЄІЇа-щьюяґєії]{1,}", file_name )
                    complete_group_names.append(matches[0]+"-"+temp)
        json_data[column_name]["groups"]=complete_group_names
    return json_data


def str_cleaner(list):
    """
    Helper func, which cleans list of strings by removing unwanted words, characters, and formatting issues. 
    This includes:
    - Removing specific phrases such as "Google meet", "парний тиждень", etc.
    - Removing prefixes like "ст.", "викл.", and names of professors (FIO).
    - Removing numeric values, commas, and extra spaces.
    """

    for e_id, element in enumerate(list):
        list[e_id] = list[e_id].replace("Google meet", "")
        list[e_id] = list[e_id].replace("Google Meet", "")
        list[e_id] = list[e_id].replace("Google class", "")
        list[e_id] = list[e_id].replace("Google сlass", "")

        even_day = re.findall(r'парний тиждень|парний тижд.|непарний тиждень|непарний тижд.', list[e_id])
        for e in even_day:
            list[e_id] = list[e_id].replace(e, '')
            list[e_id] = list[e_id].replace("-", ' ')
            list[e_id] = list[e_id].replace("–", ' ')

        surnames = re.findall(r'([А-ЩЬЮЯҐЄІЇ][а-щьюяґєії\'’]+\.*?,?\s?[А-ЩЬЮЯҐЄІЇа]+\.[А-ЩЬЮЯҐЄІЇ]+\.?)', list[e_id]) #throw out full names of the teachers
        for surname in surnames:
            list[e_id] = list[e_id].replace(surname, '')

        prefixs = re.findall(r'ст\.+ ?викл\.+|ст\.+вик\.+|пр\.+,?|доц\.+,?|викл\.+|проф\.+|доц |ст\.+| л\.+| л,|викл |лр.', list[e_id]) 
        for p in prefixs:
            list[e_id] = list[e_id].replace(p, '')

        digits = re.findall(r'\d+\/\d+|\d+\.\d+', list[e_id]) 
        for d in digits:
            list[e_id] = list[e_id].replace(d, '')
        list[e_id] = list[e_id].replace(",", '')

        space_s = re.findall(r'\s\s+', list[e_id]) 
        for s in space_s:
            list[e_id] = list[e_id].replace(s, ' ')
        list[e_id] = list[e_id].strip()
    return list 


def shorten_text(subject):
    dict =	{
    "Веб-орієнтована розробка програмного забезпечення": "Веб-орієнтована розробка",
    "Іноземна мова (за спрямуванням)": "Іноземна мова",
    "Іноземна мова (за фаховим спрямуванням)": "Іноземна мова",
    "Іноземна мова (за професійним спрямуванням)": "Іноземна мова",
    "Конструювання програмного забезпечення": "Конструювання ПЗ",
    "Основи програмування та алгоритмічні мови": "Основи програмування та алг.",
    "Інноваційне підприємництво та управління стартап проєктами (з гр.ФНМПІ-91)": "Підприємництво та стартапи",
    "Архітектура та проектування програмного забезпечення": "Архітектура та проектування ПЗ",
    "Бухгалтерський облік і звітність у ком.банках" : "Бух. облік і звітність",
    "Математика (Алгебра і початки аналізу та геометрія)" : "Математика (алгебра та геометрія)",
    "Математика (алгебра та геометрія)" : "Математика (алгебра та геометрія)",
    "Загальна теорія здоров'я діагностика і моніторинг стану здоров'я" : "Теорія здоров'я",
    "Фізична терапія при захворюваннях та порушеннях опорно-рухового апарату" : "Фіз. терапія опорно-рухового апарату",
    "Моделювання та аналіз програмного забезпечення":"Моделювання та аналіз ПЗ",
    "Якість програмного забезпечення та тестування":"Якість ПЗ та тестування",
    "Інформаційно-комунікаційні технології в менеджменті":"Комунікаційні технології",
    "Долікарська медична допомога у невідкладних станах":"Долікарська медична допомога",
    "Безпека інформаційних систем/Безпека програм та даних":"Безпека програм та БД",
    "Рекреаційна рухова активність та оздоровчий фітнес":"Оздоровчий фітнес",
    "Теорія оздоровчого харчування дієтотерапія":"Оздорове харчування",
    "Історія економіки та економічної думки (з гр.МТ-11)":"Історія економіки"
    }
    matches = sorted(dict, key=lambda x: SequenceMatcher(None, x, subject).ratio(), reverse=True)
    if(round (SequenceMatcher(None, matches[0], subject).ratio(),2)>0.75):
        return dict[matches[0]]
    else:
        return subject


def get_column_headers(table, file_name):
    """
    Extracts column headers from table by parsing through.
    Skips the first two columns (Day and Time), processes remaining columns.
    
    Args:
    - table (str): XML structure 
    - file_name (str)

    Returns:
    - list: A list of column headers extracted from the table.
    """
    column_headers=[]
    rows = re.findall(r'<w:tr .*?>(.+?)</w:tr>', table)
    for  row in rows:
        cells = re.findall(r'<w:tc>(.+?)</w:tc>', row)
        for cell in cells[2:]: # [2:] skip День,Час
            cell_paragraphs = re.findall(r'<w:p .*?>(.+?)</w:p>', cell)
            for cell_p in cell_paragraphs:
                p_text = re.findall(r'(?:<w:t>(.+?)</w:t>)|(?:<w:t .*?>(.+?)</w:t>)', cell_p)
                p_text = ''.join(list(itertools.chain(*p_text)))
                #print(p_text) #keypoint debug
                if(SequenceMatcher(None, "Дисципліна", p_text).ratio()>0.75): #if Дисципліна" = group only one
                    column_headers.append(file_name)
                    return column_headers
                else:
                    column_headers.append(re.sub(r'\([^)]*\)', '', p_text).replace(" ", "")) #removes "(123 students)"
        return column_headers


def get_tabel(file):
    """
    Fetches XML table with schedule from provided file, which coresponds THIS date.
    """
    dates = get_dates() 
    tables = re.findall(r'<w:tbl>(.+?)</w:tbl>', file.decode('utf-8'))
    for table in reversed(tables):                # start from last table
        rows = re.findall(r'<w:tr .*?>(.+?)</w:tr>', table)
        for row in rows[1:]:
            cells = re.findall(r'<w:tc>(.+?)</w:tc>', row)
            for cell in cells:
                cell_paragraphs = re.findall(r'<w:p .*?>(.+?)</w:p>', cell)
                for cell_p in cell_paragraphs:
                    p_text = re.findall(r'(?:<w:t>(.+?)</w:t>)|(?:<w:t .*?>(.+?)</w:t>)', cell_p)
                    p_text = ''.join(list(itertools.chain(*p_text)))
                    #print("["+p_text+"]")
                    
                    date_of_this_row=datetime(1, 1, 1) 
                    try: 
                        if(re.match("\d{2}.\d{2}.\d{2}", p_text)):
                            date_of_this_row = datetime.strptime(p_text, '%d.%m.%yр.')
                        elif(re.match(".\d{2}.\d{2}.\d{2}", p_text)):
                            date_of_this_row = datetime.strptime(p_text, '.%d.%m.%yр.')
                        elif(re.match("\d{2}.\d{2}.", p_text)):
                            date_of_this_row = datetime.strptime(p_text, '%d.%m.')
                            date_of_this_row = date_of_this_row.replace(year=dates[0].year)
                        elif(re.match("\d{2}.\d{2}", p_text)):
                            date_of_this_row = datetime.strptime(p_text, '%d.%m')
                            date_of_this_row = date_of_this_row.replace(year=dates[0].year)
                        if (date_of_this_row >=dates[1] and date_of_this_row <= dates[2]):    
                            #print(str(date_of_this_row).split(' ')[0])
                            #print("ok")
                            return table
                    except:
                        pass
    return 0   


def get_dates(): 
    """
    Returns the current date and the start and end dates of the current or next week.
    """
    #today=datetime(2022, 4, 12)      
    today = datetime(date.today().year, date.today().month, date.today().day)
    if(today.weekday()<=4): # if Mon-Fri print this week
        start_week=today - timedelta(days=today.weekday())
        end_week=start_week + timedelta(days=6, hours=23, minutes=59)
    else: #else print next week
        start_week=today + timedelta(days=(7-today.weekday()))
        end_week=start_week + timedelta(days=6, hours=23, minutes=59)
    return [today, start_week, end_week]