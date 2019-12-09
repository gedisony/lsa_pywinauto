from pywinauto import application
#from pywinauto.win32functions import SetForegroundWindow
import time
import os,shutil
from datetime import datetime

class user_maintenance_screen:

    def __init__(self,app):
        self.app=app

    def connect(self ):
        try:
            #app.window(best_match='User Maintenance - '+server)  does not work. picks the prevoious window.
            print( '    connecting to an existing instance of the User Manintnance Screen' )
            self.screen=self.app.window(title_re='User Maintenance*')     
            print(  '      connected '  + str(self.screen.is_visible()) )
            #self.screen.minimize()
            #self.screen.restore()
            self.screen.set_focus()
            #SetForegroundWindow(self.app)

        except application.findwindows.ElementNotFoundError:
            print( '      cannot connect' )
            main_screen=self.app.window(title_re='Lawson Security*', found_index=0)  #there will be 2 found

            main_screen.wait('visible')

            #main_screen.print_control_identifiers()

            print( '    opening user maintenance screen' )
            main_screen.set_focus()
            #main_screen.type_keys('%{w}') #alt-w
            main_screen.type_keys('^U') #ctrl-U

            self.screen=self.app.window(title_re='User Maintenance*')  

        #print( '    wait - visible' )
        self.screen.wait('visible')
        #self.screen.print_control_identifiers()
        #print( '    done' )

    def add_users_adv(self, user_ids):
        
        al=user_ids.split(',')


        #clear the contents of the query results pane
        self.screen[' &Clear All'].click()

        #click the Advanced tab where we can enter ID's
        #self.screen.ThunderRT6Frame8.click_input( coords=(45, -15), absolute=False) 
        #self.screen['F&ind Now'].set_focus()
        self.screen.type_keys('{UP}{UP}{RIGHT}')
    
        for user_id in al:
            #print(user_id)
            #self.screen.ComboBox3.draw_outline()
            self.screen.ComboBox3.type_keys('Addins{UP}') #pick activity list (cannot directly pick Activity List)
            
            #self.screen.Edit2.draw_outline()  #strange that this is Edit2 not Edit3
            self.screen.Edit2.type_keys(user_id)
            
            self.screen['&AddButton'].click()  #self.screen.Edit2.type_keys('{ENTER}') will also work

        self.screen.VSFlexGrid8L2.set_focus()
        for user_id in al[1:]:
            self.screen.VSFlexGrid8L2.type_keys('{DOWN}')
            self.screen.VSFlexGrid8L2.type_keys('x')
            self.screen.VSFlexGrid8L2.type_keys('O') #need to send keys twice

    def find_now(self):
        self.screen['F&ind Now'].click()

class report_maintenance_screen:

    def __init__(self,app):
        self.app=app

    def connect(self ):
        try:
            #app.window(best_match='User Maintenance - '+server)  does not work. picks the prevoious window.
            print( '    connecting to an existing instance of the Report Screen' )
            self.screen=self.app.window(title_re='Report Maintenance*')     
            print(  '      connected '  + str(self.screen.is_visible()) )
            self.screen.set_focus()

        except application.findwindows.ElementNotFoundError:
            print( '      cannot connect' )
            main_screen=self.app.window(title_re='Lawson Security*', found_index=0)  #there will be 2 found

            main_screen.wait('visible')

            #main_screen.print_control_identifiers()

            print( '    opening report maintenance screen' )
            main_screen.set_focus()
            time.sleep(1)
            #main_screen.set_focus()
            #time.sleep(1)
            print( '    Ctrl+F11' )
            main_screen.type_keys('^{F11}') #ctrl-F4
            time.sleep(1)
            main_screen.type_keys('^{F11}') #ctrl-F4
            self.screen=self.app.window(title_re='Report Maintenance*')  
            
        print( '    wait - visible' )
        self.screen.wait('visible')
        
        #print( '    done' )

        #having a hard time getting the controls...
            #self.screen['StatusBar'].print_control_identifiers()
            #child_window=self.screen.child_window[0].print_control_identifiers()
            #self.screen.child_window(class_name="StatusBar20WndClass").print_control_identifiers()

            #a=self.screen['VSFlexGrid8L']
            #a.draw_outline()
            #self.screen['StatusBar'].draw_outline()

        #refresh the list so that the {DOWN} key will be item 1 on the list
        self.screen.type_keys('%s')

        #the report we want is the first item on the list
        self.screen.type_keys('{DOWN}')

        #shortcut for Run Report
        self.screen.type_keys('%r')

        rr=self.app.window(title_re='Run Report*', found_index=0) 
        #rr.print_control_identifiers()
        rr['Run ReportVSFlexGrid8L'].type_keys('{SPACE}') #check the csv box on the list(first item)
        rr['&OkButton'].click()
        

        #refresh the list (wait until report is ready) 
        self.screen.type_keys('%s')

        print('    waiting for report to finish.')
        self.screen['VSFlexGrid8L'].draw_outline( )
        time.sleep(7)
        
        print('    waiting for report to finish..')
        time.sleep(7)
        self.screen.type_keys('%s')
        self.screen.type_keys('{DOWN}')

        print('    waiting for report to finish...')
        time.sleep(5)
        self.screen.type_keys('%s')
        self.screen.type_keys('{DOWN}')

        #view report
        self.screen.type_keys('%v')
        vr=self.app.window(title_re='View Report*', found_index=0) 
        #vr.print_control_identifiers()
        vr['View ReportVSFlexGrid8L'].type_keys('{SPACE}')  #check the csv box on the list
        vr['&View File'].click()

        print('    waiting to finish file download..')
        
        '''
            self.app.wait_cpu_usage_lower(threshold=5) did not work with python 3.7
            download and install python 3.6
            in vs Code Ctrl-Shift+P. Python:select Interpreter (pick appropriate version)
            install pywinauto in 3.6 by referring to the 3.6 python (not the 3.7 default)
            C:\\Users\\......\\Python\\Python36\\python -m pip install pywinauto
        '''
        self.app.wait_cpu_usage_lower(threshold=5,timeout=55) #a timeout=None will still time out. specify a large enough value

        #temp_admin-yau_ed_all_users_2019.Aug.05
        xls_filename='temp_admin-yau_ed_all_users_' + datetime.today().strftime('%Y.%b.%d') + '.csv'
        xls_basedir=r'C:\Users\1071102\AppData\Local\Temp'
        xls_full_path=os.path.join(xls_basedir,  xls_filename )
        print('    excel will open ' + xls_filename)


        excel_app = application.Application(backend='win32')
        time.sleep(5)
        excel_app.connect(title_re='temp_admin-yau*')
        #xlmain = excel_app["temp.csv [Read Only] - Excel"]
        xlmain = excel_app.Window_(title_re="temp_admin-yau*") 


        #xlmain.type_keys('%F')
        xlmain.close()

        print('    excel closed')

        time.sleep(2)

        shutil.copy( xls_full_path ,r'\\sftp\windp\FTP\Lawsontest' )

        print('    '+xls_full_path+' copied to windp - lawsontest')
