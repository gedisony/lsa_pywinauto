#from pywinauto.application import Application
from multiprocessing import Process
from pywinauto import application
from pywinauto import mouse
from classes import user_maintenance_screen, report_maintenance_screen
import base64
import time
import shutil

#config file has server and credentials information so that code can be published on github without exposing 
#add config.py to gitignore
import config

def main():

    #print(config.lsa_admin_cred.password() )
    #return 

    mode = input("1  Run Lawson Security Admin only\n" \
        + "2  LSA + LI2D + sec audit report\n" \
        + "3  LID export sso config\n" \
        + "4  LSA + (csv of users)\n" )  

    options[mode](mode)

    #start_lawson_lsa()

    #run_lid_sso_config_export()
    return    
    

def start_lawson_lsa(action=None):
    #app = Application().start("C:\Program Files (x86)\Lawson Software\Security Administrator\LawsonSecAdmin.exe")
    app=application.Application(backend='win32')

    # https://pywinauto.readthedocs.io/en/latest/getting_started.html

    #connect to the running app, if can't, start it up.
    try:
        app.connect(path = r"C:\Program Files (x86)\Lawson Software\Security Administrator\LawsonSecAdmin.exe")
        time.sleep(1)

        #while debugging all of a sudden, app.connect() will nto throw an exception. manually do it here. 
        raise  application.ProcessNotFoundError

    except application.ProcessNotFoundError:
        app.start(r"C:\Program Files (x86)\Lawson Software\Security Administrator\LawsonSecAdmin.exe")

        #login_dialog = app.window(title='Login')   --old version of lsa have this 
        #login_dialog = app.window(title_re="*")  #hard time trying to guess the corrent title
        #server_select_dialog = app.top_window()  #this works
        #after top_window() works, was able to figure out the title from print_control_identifiers()
        server_select_dialog = app.window(title="Server Information")  


        #server_select_dialog.print_control_identifiers()
        #server_select_dialog.type_keys('{VK_DOWN}')
        server_select_dialog['Server URLComboBox'].draw_outline()
        server_select_dialog['Server URLComboBox'].type_keys(config.lsa_admin_cred.hostname)  #server url here
        #server_select_dialog.LoginComboBox.draw_outline()
        #server_select_dialog.Edit4.draw_outline()
        #server_select_dialog.LoginEdit1.type_keys(server) #don't specify a server name so that it will pick hrprap-2.staff.fusd.local instead of hrprap-2.staff.fusd.local:443
        #server_select_dialog.LoginEdit2.type_keys(user)
        #server_select_dialog.LoginEdit3.type_keys(passw)
        server_select_dialog['Connect'].click()
        #connecting...
        server_select_dialog.wait_not('visible')

        time.sleep(1)
        #there is an wuthenticating dialog but it looks like it is the same dialog as the login dialog(next screen) 
        #  only certain elements are visible/hidden
        #  TODO: figure out which ewlements are hidden/visible and use that instead of time.sleep() 
        #   to determine login screen ready.
        #autenticating_message_dlg=app.top_window()
        #autenticating_message_dlg.print_control_identifiers();quit()
        #autenticating_message_dlg.wait_not('visible')
        time.sleep(3)

        #login_dialog = app.top_window()  #fot the title adter using the top_window() + print_control_identifiers()
        login_dialog=app.window(title_re='Authenticating to*', found_index=0)
        #print_control_identifiers does not give a useful result lookls like an internet explorer kind of control and the inner controld are not available.  
        #login_dialog.print_control_identifiers(); quit()

        #login_dialog['WindowsForms10.Window.8.app.0.297b065_r62_ad14'].draw_outline()   #ok
        #login_dialog['Shell Embedding'].draw_outline()
        #time.sleep(1)
        login_dialog.type_keys('{TAB}')
        #password box
        login_dialog.type_keys('{TAB}')
        #sign in button
        login_dialog.type_keys('{TAB}')
        #invisible
        login_dialog.type_keys('{TAB}')
        #username
        login_dialog.type_keys(config.lsa_admin_cred.username) 
        login_dialog.type_keys('{TAB}')
        #passw
        login_dialog.type_keys(config.lsa_admin_cred.password()) 
        login_dialog.type_keys('{TAB}')
        time.sleep(1)

        login_dialog.type_keys('{ENTER}')

        #wait for finish authentication
        time.sleep(3)

        #this does not work; tabbing works but typing keys doesn't
        #login_dialog['Shell Embedding'].type_keys('{TAB}')
        #time.sleep(1)

    if action=='0' or action=='1':
        return
    elif action=='2':
        run_sec_audit(app)
    else :   
        list_users (app)         

    #list_users (app)         
    #run_sec_audit(app)

    """
    #main_screen.close()
    #mouse.click(button='left', coords=(150, 150)) #works but is absolute

    """

def list_users( app ):
    um=user_maintenance_screen( app )
    um.connect()
    #activity list not RMID
    um.add_users_adv('1077439,1077421,1077457,1048965') 
    um.find_now()



def run_sec_audit( app ):
    
    p = Process(target=run_lid_sso_config_export, args=())
    # you have to set daemon true to not have to wait for the process to join
    p.daemon = True
    p.start()

    r_scr=report_maintenance_screen( app )
    r_scr.connect()

def run_lid_sso_config_export(action=None):
    app=application.Application(backend='win32')

    app.start(r"C:\Program Files\Lawson Software\Lawson Interface Desktop\univwin64.exe")

    main_dialog = app.top_window()  #this works
    #main_dialog = app.window(title="Terminal")  #this also works
    #main_dialog.print_control_identifiers()
    #after top_window() works, was able to figure out the title from print_control_identifiers()

    #the connect button
    #main_dialog['BUTTON_TOOL5'].draw_outline()
    main_dialog['BUTTON_TOOL5'].click()
    #time.sleep(1)
    #print("aaa")

    #comm_type_dlg=app.top_window()
    comm_type_dlg=app.window(title="Communications Type")

    #comm_type_dlg.draw_outline()
    #comm_type_dlg.print_control_identifiers()

    #comm_type_dlg['RadioButton0'].draw_outline()  
    comm_type_dlg['OKButton'].draw_outline()   
    comm_type_dlg['OKButton'].click()


    login_dlg=app.window(title="Connect To Server")
    #login_dlg.draw_outline(); login_dlg.print_control_identifiers()
    #time.sleep(1)
    #login_dlg['ServerComboBox'].draw_outline()   
    login_dlg['ServerComboBox'].type_keys(config.sso_config_cred.hostname)
    
    #login_dlg['PortEdit'].SetText('1026')
    login_dlg['Login IDEdit'].set_edit_text(config.sso_config_cred.username)                     #.SetText('xxx')
    login_dlg['DomainEdit'].set_edit_text(config.sso_config_cred.domain)                         #.SetText('xxx')
    login_dlg['PasswordEdit'].type_keys(config.sso_config_cred.password() )   #SetText / type_keys minor difference between the two 
    
    login_dlg['OKButton'].click()
    time.sleep(1)

    #main_dialog.close(); return

    main_dialog.type_keys('ssoconfig{SPACE}-c')
    main_dialog.type_keys('{ENTER}')
    time.sleep(10)

    #Please enter the password used for Lawson security utilities:
    main_dialog.type_keys(config.sso_config_cred.password())  #just the password again

    main_dialog.type_keys('{ENTER}')
    time.sleep(1)

    main_dialog.type_keys('5')   #(5) Manage Lawson Services
    main_dialog.type_keys('{ENTER}')
    time.sleep(1)

    #-------Manage Lawson Services-------
    main_dialog.type_keys('6')   #(6) Export service and identity info
    main_dialog.type_keys('{ENTER}')
    time.sleep(1)

    ##Do you want to export all the services?
    main_dialog.type_keys('1')  #  (1) YES
    main_dialog.type_keys('{ENTER}')
    time.sleep(1)

    #Do you want to export the identities ("ALL" or "NONE") (<empty>)
    main_dialog.type_keys('ALL')
    main_dialog.type_keys('{ENTER}')
    time.sleep(1)

    #Enter file name to save export as (<empty>):ssoconfig_dump
    main_dialog.type_keys('ssoconfig_dump')
    main_dialog.type_keys('{ENTER}')
    time.sleep(1)

    #Choose format that Lawson Software should export credential information as.
    print('    make sure the ssocinfig_dump files does not exist or there will be an extra prompt and automation will fail ')
    main_dialog.type_keys('2')   #(2) Opaque
    main_dialog.type_keys('{ENTER}')
    time.sleep(20)

    # Export finished.
    main_dialog.type_keys('13')   #(13) Exit
    main_dialog.type_keys('{ENTER}')
    time.sleep(1)

    #the created file will be on:
    #\\hrprap-2\c$\LSFPROD\launt\STAFF_lawson
    print('    copying ssoconfig_dump to sid-dev (youll get permission error if you have just restarted pc and have not accessed \\hrprap-2) ')
    try:
        shutil.move( r'\\hrprap-2\c$\LSFPROD\launt\STAFF_lawson\ssoconfig_dump',r'\\sid-dev\temp\ssoconfig_dump' )
    except:
        main_dialog.type_keys('NOTE:Manually{SPACE}copy{SPACE}ssoconfig_dump{SPACE}to{SPACE}sid-dev')
        time.sleep(10)
        
    else:
        print('    ssoconfig_dump copied to sid-dev')

    main_dialog.close()


options = {'0' : start_lawson_lsa,
           '1' : start_lawson_lsa,
           '2' : start_lawson_lsa,
           '3' : run_lid_sso_config_export,
           '4' : start_lawson_lsa,
}



if __name__== "__main__":
    main()

