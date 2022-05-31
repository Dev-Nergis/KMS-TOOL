import os, time, progressbar, platform

os.system("echo @off")
os.chdir("C:\Program Files\Microsoft Office\Office16")
os.system("Cscript ospp.vbs /remhst")
os.system("Cscript ospp.vbs /ckms-domain")
os.system("Cscript ospp.vbs /actype:0")
os.system("Cscript ospp.vbs /skms-domain:kms.textcord.xyz")
os.system("Cscript ospp.vbs /cachst:TRUE")
os.system("Slmgr //b /ckms >nul")
os.system("Slmgr //b /ckms-domain >nul")
os.system("Slmgr //b /skms-domain kms.textcord.xyz >nul")
os.system("cls")

def lodbar():
    widgets = ['Loading: ', progressbar.AnimatedMarker()]
    bar = progressbar.ProgressBar(widgets=widgets).start()
      
    for i in range(50):
        time.sleep(0.1)
        bar.update(i)

System = platform.system()
Release = platform.release()
Edition = platform.win32_edition()

print(f"{System}\n{Release}\n{Edition}")

print("""                                                              
    █████ █████ █████ █████ █████ █████ █████ █████ █████ █████ █████

    Don't Use This Script Without Admin Permission.

    ██   ██ ███    ███ ███████     ████████  ██████   ██████  ██     
    ██  ██  ████  ████ ██             ██    ██    ██ ██    ██ ██     
    █████   ██ ████ ██ ███████        ██    ██    ██ ██    ██ ██     
    ██  ██  ██  ██  ██      ██        ██    ██    ██ ██    ██ ██     
    ██   ██ ██      ██ ███████        ██     ██████   ██████  ███████

    Made by: Dev_Nergis                                              

    █████ █████ █████ █████ █████ █████ █████ █████ █████ █████ █████
""")

print("Offical KMS Server: kms.textcord.xyz")
print("Key Type: GVLK")

print("\033[1;35m")
print("01. Windows 10/11 Pro")
print("\033[1;36m")
print("02. Windows 10/11 Pro N")
print("\033[1;35m")
print("03. Windows 10/11 Pro for Workstations")
print("\033[1;36m")
print("04. Windows 10/11 Pro for Workstations N")
print("\033[1;35m")
print("05. Windows 10/11 Pro Education")
print("\033[1;36m")
print("06. Windows 10/11 Pro Education N")
print("\033[1;35m")
print("07. Windows 10/11 Education")
print("\033[1;36m")
print("08. Windows 10/11 Education N")
print("\033[1;35m")
print("09. Windows 10/11 Education KN")
print("\033[1;36m")
print("10. Windows 10/11 Enterprise")
print("\033[1;35m")
print("11. Windows 10/11 Enterprise N")
print("\033[1;36m")
print("12. Windows 10/11 Enterprise KN")
print("\033[1;35m")
print("13. Windows 10/11 Enterprise G")
print("\033[1;36m")
print("14. Windows 10/11 Enterprise G N")

print("\033[1;37m")
print("Microsoft Offical")
print("\033[1;30m")
print("00. Office 2010/2013/2016/2019/2021 (Maybe all Office Versions)")

print("\033[1;34m")
Ed = str(input("Choice your Windows Edition or Office: "))
print("""
---------------  ---------------  ---------------  ---------------  ---------------  --------------- 
-:::::::::::::-  -:::::::::::::-  -:::::::::::::-  -:::::::::::::-  -:::::::::::::-  -:::::::::::::- 
---------------  ---------------  ---------------  ---------------  ---------------  ---------------                                                                                               
""")

print("\033[1;33m")
if Ed == "00":
    lodbar()
    os.system("Cscript ospp.vbs /act")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "01":
    lodbar()
    os.system("Slmgr //b /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "02":
    lodbar()
    os.system("Slmgr //b /ipk MH37W-N47XK-V7XM9-C7227-GCQG9 >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "03":
    lodbar()
    os.system("Slmgr //b /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "04":
    lodbar()
    os.system("Slmgr //b /ipk 9FNHH-K3HBT-3W4TD-6383H-6XYWF >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "05":
    lodbar()
    os.system("Slmgr //b /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "06":
    lodbar()
    os.system("Slmgr //b /ipk YVWGF-BXNMC-HTQYQ-CPQ99-66QFC >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "07":
    lodbar()
    os.system("Slmgr //b /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2 >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "08":
    lodbar()
    os.system("Slmgr //b /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "09":
    lodbar()
    os.system("Slmgr //b /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "10":
    lodbar()
    os.system("Slmgr //b /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43 >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "11":
    lodbar()
    os.system("Slmgr //b /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4 >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "12":
    lodbar()
    os.system("Slmgr //b /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4 >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "13":
    lodbar()
    os.system("Slmgr //b /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
elif Ed == "14":
    lodbar()
    os.system("Slmgr //b /ipk 44RPN-FTY23-9VTTB-MP9BX-T84FV >nul")
    os.system("Slmgr //b /ato >nul")
    print("\n")
    print("\033[1;32m Done!")
    input("\n Press Enter to exit... \n")
else:
    print("\033[1;31m")
    print("\n")
    print("Error! Please try again...")
    input("\n Press Enter to exit... \n")
#######################################################################################################################################################################################################