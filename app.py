import bot_functions    

if __name__ == "__main__" :
    bot_functions.init_project_structure()
    bot_functions.log_options_init()
    #a = bot_functions.grab_docx_files()
    #print(a)
    bot_functions.docxs_handler()