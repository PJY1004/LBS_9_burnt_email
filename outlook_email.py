import os
import time
import shutil
from datetime import datetime, timedelta
import win32com.client

folder_path_for_alarm = "./"

new_mail_body = "<b>Number,Site,Test Name,Channel,Low,Measured(mA),High,Force,Loc</b><br>"
mail_line = 1

while (1):
    if folder_path_for_alarm[len(folder_path_for_alarm) - 1] != "/":
        folder_path_for_alarm = folder_path_for_alarm + "/"

    mail_body = ""
    count = 0

    ############### 변수 folder_path_for_alarm에 저장된 (Burnt Warning) txt 파일의 마지막 수정 시간을 구하고, 5분 이내면 이메일로 보냄 ##############################

    each_file_gen_time = os.path.getmtime("LBUF_Burnt_Log.csv")
    five_minute_after_alarm = datetime.fromtimestamp(each_file_gen_time) + timedelta(minutes=5)

    if (five_minute_after_alarm > datetime.now()):  # 현재 시간과 5분 차이 이내면 이메일 보냄.
        count += 1
        shutil.copy("LBUF_Burnt_Log.csv", "(copy)LBUF_Burnt_Log.csv")
        f = open("(copy)LBUF_Burnt_Log.csv", 'r')
        lines = f.readlines()
        if len(lines) <= 1:
            mail_body += "Burnt Alarm System 시작됨."
        else:
            read_start_line = mail_line
            for line in lines[read_start_line:]:
                mail_body += (line + "<br>")
                mail_line += 1

            if "Number,Site,Test Name,Channel,Low,Measured(mA),High,Force,Loc" in mail_body:
                mail_body.replace("Number,Site,Test Name,Channel,Low,Measured(mA),High,Force,Loc","<b>기준 이상값 발생!</b><br>Number,Site,Test Name,Channel,Low,Measured(mA),High,Force,Loc<br>")
            else:
                mail_body = ("<b>기준 이상값 발생!</b><br>Number,Site,Test Name,Channel,Low,Measured(mA),High,Force,Loc<br>" + mail_body)

        f.close()
        os.remove("(copy)LBUF_Burnt_Log.csv")
        mail_body += "<br>"

    new_mail_body = mail_body

    #################################################################################################################################################################

    if count != 0:  # 폴더 내에 마지막으로 5분 이내에 수정된 파일이 없으면 이메일 보내지x
        outlook = win32com.client.Dispatch("Outlook.Application")
        send_mail = outlook.CreateItem(0)

        send_mail.To = "ghjeong@lbsemicon.com;smcho@lbsemicon.com;"  # 메일 수신인 ex) 정규희 수석 , 조성민 팀장에게 송신하려면 send_mail.To = "ghjeong@lbsemicon.com;smcho@lbsemicon.com;"
        # 정규희 수석 : ghjeong@lbsemicon.com;
        # 조성민 팀장 : smcho@lbsemicon.com;
        send_mail.CC = "ejlee@lbsemicon.com;jhlim@lbsemicon.com;LBS_SOC@lbsemicon.com;"  # 메일 참조.. ex) 이은진 사원 , 임진형 사원에게 참조하려면 send_mail.CC = "ejlee@lbsemicon.com;jhlim@lbsemicon.com;"
        # 변형찬 담당 : hcbyun@lbsemicon.com;
        # 이은진 사원 : ejlee@lbsemicon.com;
        # 임진형 사원 : jhlim@lbsemicon.com;
        # 제품기술2팀 : LBS_SOC@lbsemicon.com;
        send_mail.Subject = "Burnt Warning"  # 메일 제목
        send_mail.HTMLBody = mail_body  # 메일 내용
        send_mail.Send()
    time.sleep(300)
