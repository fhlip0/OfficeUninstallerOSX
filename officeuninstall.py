#!/usr/bin/env python
# coding: utf-8
import sys, getopt
import os
import subprocess
import urllib
import time

def _restart_dock(user):
    command = "sudo -u "+user+" osascript -e 'delay 3' -e 'tell Application \"Dock\"' -e 'quit' -e 'end tell' -e 'delay 3'"
    p = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    p.communicate()

def _remove_icon(user, positions):
    for position in positions:
        command = 'sudo -u '+user+' /usr/libexec/PlistBuddy -c "Delete persistent-apps:'+str(position)+'" /Users/'+user+'/Library/Preferences/com.apple.dock.plist'
        p = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        p.communicate()


def _get_users():
    command = "dscacheutil -q user | grep -A 3 -B 2 -e uid:\\ 5'[0-9][0-9]' | grep name"
    p = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    out, err = p.communicate()
    out = out.strip()
    lines = out.split("\n")
    users = []
    for line in lines:
        users.append(line.replace("name: ", "").strip())
    print "Users:", users
    return users

def _get_doc_icon_positions(app_name, user):
    app_name = urllib.quote("/Applications/Microsoft")
    app_name = app_name.replace("/", "\\/")
    print "Encoded App name:", app_name
    command = "sudo -u "+user+" defaults read com.apple.dock persistent-apps | grep _CFURLString\\\" | awk '/"+app_name+"/ {print NR}'"
    p = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    out, err = p.communicate()
    out = out.strip()
    lines = out.split("\n")
    positions = []
    try:
        for line in lines:
            pos = int(line) - 1
            positions.append(pos)
    except Exception as e:
        pass
    return positions

def dockicons(argv):
    app_name = '/Applications/Microsoft'
    app_ignore = []
    try:
        opts, args = getopt.getopt(argv,"hr:i:",["remove=","ignore="])
    except getopt.GetoptError:
        print 'dock-icon-remove.py -r <partial_name_to_remove> -i <partial_name_to_ignore>'
        sys.exit(2)
    
    for opt, arg in opts:
        if opt == '-h':
            print 'dock-icon-remove.py -r <partial_name_to_remove> -i <partial_name_to_ignore>'
            sys.exit()
        elif opt in ("-i", "--ignore"):
            app_ignore_raw = app_ignore.append(arg)
        elif opt in ("-r", "--remove"):
            app_name = arg

#   if app_name is None:
#      print "Missing remove flag"
#     sys.exit(1)

    print "Removing Dock icons for all users - This will take a moment and screen will flicker"
    time.sleep(3)

    print "Removing", app_name, "icon from dock and ignoring:", app_ignore
    
    users = _get_users()
    for user in users:
        ignored_positions = set()
        doc_positions = _get_doc_icon_positions(app_name, user)
        for app in app_ignore:
            ignore_positions = _get_doc_icon_positions(app, user)
            for p in ignore_positions:
                ignored_positions.add(p)
    
        clean_positions = []
        for pos in doc_positions:
            if pos not in ignored_positions:
                clean_positions.append(pos - len(clean_positions))

        print "Ignored Positions:", ignored_positions
        if len(clean_positions) == 0:
            print "User", user, "does not have it in dock"
            continue
        
        print "User", user, "has it in dock"+str(clean_positions)
        _remove_icon(user, clean_positions)
        _restart_dock(user)
    print "."
    print "Dock Icon Removal Complete! - EXITING"


#if __name__ == "__main__":
#   main(sys.argv[1:])



    
def office2011():
    print "Uninstalling Office from Applications folder"
    time.sleep(3)
    os.system('rm -r /Applications/Microsoft\ Office\ 2011')
    print "Uninstalling Communicator"
    os.system('rm -r /Applications/Microsoft\ Communicator.app')
    print "Uninstalling Messenger"
    os.system('rm -r /Applications/Microsoft\ Messenger.app')
    print "Uninstalling Remote Desktop Connection"
    os.system('rm -r /Applications/Remote\ Desktop\ Connection.app')
    print "."
    print "Removing Preference Files (takes a minute)"
    print "Also can produce a lot of -No Directory Found- totally normal!!!"
    print "."
    time.sleep(5)
    os.system('rm -r /Library/Preferences/com.microsoft.Word.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.Excel.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.Powerpoint.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.Outlook.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.outlook.databasedaemon.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.DocumentConnection.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.office.setupassistant.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.Word.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.Excel.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.Powerpoint.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.Outlook.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.outlook.databasedaemon.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.outlook.office_reminders.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.DocumentConnection.plist')
    os.system('rm -r /ibrary/Preferences/com.microsoft.office.setupassistant.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.office.plist')
    os.system('rm -r /Library/Preferences/com.microsoft.error_reporting.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.Word.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.Excel.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.Powerpoint.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.Outlook.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.outlook.databasedaemon.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.DocumentConnection.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.office.setupassistant.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.registrationDB.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.e0Q*.*.plist')
    os.system('rm -r /Library/Preferences/ByHost/com.microsoft.Office365.*.plist')
    print "."
    print "Preferences Done - uninstalling backend components"
    print "."
    time.sleep(5)
    os.system('rm -r /Library/LaunchDaemons/com.microsoft.office.licensing.helper.plist')
    os.system('/Library/Preferences/com.microsoft.office.licensing.plist')
    os.system('/Library/Application Support/Microsoft/MERP2.0')
    os.system('rm -r ~/Library/Caches/com.microsoft.browserfont.cache')
    os.system('rm -r ~/Library/Caches/com.microsoft.office.setupassistant')
    os.system('rm -r ~/Library/Caches/Microsoft/Office')
    os.system('rm -r ~/Library/Caches/Outlook')
    os.system('rm -r ~/Library/Caches/com.microsoft.Outlook')
    os.system('rm -r /Library/Fonts/Microsoft')
    print "."
    print "Office 2011 Uninstall Complete"
    print "."

def office2016():
    print "Starting Uninstall of Office 2016"
#Office 2016 uninstall script v 0.01
#deletes application from primary Applications Folder


    os.system('rm -r /Applications/Microsoft\ Powerpoint.app')
    print "Uninstalling Powerpoint"
    os.system('rm -r /Applications/Microsoft\ Word.app')
    print "Uninstalling Word"
    os.system('rm -r /Applications/Microsoft\ Outlook.app')
    print "Uninstalling Outlook"
    os.system('rm -r /Applications/Microsoft\ OneNote.app')
    print "Uninstalling OneNote"
    os.system('rm -r /Applications/Microsoft\ Excel.app')
    print "Uninstalling Excel"

    #Now for Library / Containers items
    print "Clearing Library Container items"
    os.system('rm -r ~/Library/Containers/com.microsoft.errorreporting')
    os.system('rm -r ~/Library/Containers/com.microsoft.Excel')
    os.system('rm -r ~/Library/Containers/com.microsoft.netlib.shipassertprocess')
    os.system('rm -r ~/Library/Containers/com.microsoft.Office365ServiceV2')
    os.system('rm -r ~/Library/Containers/com.microsoft.Outlook')
    os.system('rm -r ~/Library/Containers/com.microsoft.Powerpoint')
    os.system('rm -r ~/Library/Containers/com.microsoft.RMS-XPCService')
    os.system('rm -r ~/Library/Containers/com.microsoft.Word')
    os.system('rm -r ~/Library/Containers/com.microsoft.onenote.mac')

    #Remove group containers
    print "Clearing Group Containter items"

    os.system('rm -r ~/Library/Group\ Containers/UBF8T346G9.MS')
    os.system('rm -r ~/Library/Group\ Containers/UBF8T346G9.Office')
    os.system('rm -r ~/Library/Group\ Containers/UBF8T346G9.OfficeOsfWebHost')
    print "."
    print "."
    print "."
    print "WARNING: Some items may say -No such file or directory found- this is normal"
    print "."
    print "."
    print "."
    print "Office 2016 Uninstall Complete\n"


    # map the inputs to the function blocks
print "Uninstalling Office 2011 or Office 2016?"
version = input('Enter either "2011" or "2016": ')


options = {2011 : office2011,
    2016 : office2016,
}
options[version]()


print "Clear Office dock icons? (yes or no) :"

cleardock = raw_input().lower()
if cleardock == "yes":
    dockicons(sys.argv[1:])
else:
    print "Uninstaller exiting"





