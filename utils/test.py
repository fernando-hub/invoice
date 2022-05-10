import configparser

OperatingSystem="windows"
ClearScreen="cls"
def setup():
        """
        global OperativeSystem
        OperativeSystem=input("Windows=1, Linux/Mac=2 ")
        print (OperativeSystem)
        if  OperativeSystem !="1" and OperativeSystem !="2":
            OperativeSystem="2"

        if  OperativeSystem=="2":
            ClearScreen="clear"
        """

       
        config = configparser.ConfigParser(allow_no_value=True)
        config['PATHD'] = {'locationexport': 'C:/user///',
                             'OperatingSystem': 'windows',
                             'Email':'fernandoforce@gmail.com,fermay123@hotmail.com'}
                             
   
        

        with open('app.config', 'w') as configfile:
          config.write(configfile)
setup()