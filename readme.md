# Uni timetable parser & formatter

This kinda pet project done due to issue where university constantly made timetable updates without notifying students. Its designed to fetch, parse, process, and format university schedule data from `.docx` files on its site on level of XML markup of docx file,  into own structured JSON format and further turn into easy to use, up to date png`s.

## Table of Contents

- [Features](#features)
- [Technologies](#technologies)
- [Installation](#installation)
- [Usage](#usage)
- [Example_Output](#example_output)
- [University_Information](#university_information)
- [Acknowledgements](#acknowledgements)
---

## Features

This project allows you to:

- Fetch `.docx` files containing university schedules from it`s site.
- Parse received data on XML level of docx structure.
- Clean up the data by removing unnecessary information, student names, extra info, etc.
- Store schedule data as JSON object for easy use in other future system
- Process these timetable data to get easy to use png

The script processes weekly schedule data, extracts the relevant details (such as subject names, groups, times, and days), and cleans up any extraneous or unwanted text.


## Technologies

This project uses the following Python libraries and technologies:

- **Python 3.8**
- **requests**: To make HTTP GET/POST requests.
- **BeautifulSoup**: For parsing from XML tables.
- **Regular expressions**: For regex-based pattern matching and text cleaning.
- **dpath**: External lib to navigate nested dictionaries.
- **PIL (Pillow)**: For image processing, generating visual output.
- **datetime**: For handling time-related data.
- **logging**: To log minor debug info.

## Installation

Follow these steps to set up the project locally:

1. Clone the repository:
   ```bash
   git clone ...

2. Make project enviroment:
    ```bash
    virtualenv venv --python=python3.8
3. activate enviroment:
    ```bash
    .\venv\Scripts\activate     # Note, for Windows
4. set-up dependencies and libs needed:
    ```bash
    pip install -r requirements.txt
5. Create and configure the .env file:
    ```bash
    # .env example
    mntu_password=
    mntu_base_url=http://lib.istu.edu.ua/

6. You are ready to go locally !

## Usage
    # Note: On my production/use case scenario it worked as 4h Cron task.
    
Example output provided further, which intended to use as simple image, fast share in internal use, between groups/students versus long chain of actions like: login to uni site -> search own course / file / groups -> download file -> search needed week of study -> get need info...

## Example_Output

Below is an example of the generated timetable image which :

![Uni Timetable Example](/res/pics/І-01І-02І-03К-01.png)

#### File Path:
`Uni_timetable/res/pics/І-01І-02І-03К-01.png`
#### Some other output examples are at:
`Uni_timetable/res/pics/...`
#### Also, input examples are at:
`Uni_timetable/res/docx/...`
#### As well as partially processed data examples at:
`Uni_timetable/res/json/...`

## University_Information

This project is developed to use with **MNTU university, by Academician Yuriy Buhai**.

Original named as: **МНТУ ім. академіка Юрія Бугая (@mntu.kyiv)**.

http://istu.edu.ua/

http://lib.istu.edu.ua/

## Acknowledgements

Special thanks to [davecra/OpenXmlFileViewer](https://github.com/davecra/OpenXmlFileViewer) for providing usefull open source tool to work with XML / .docx files for this project.
