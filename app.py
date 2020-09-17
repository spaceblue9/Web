import json
import os
from flask import Flask
from flask import request
from flask import make_response
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
from datetime import datetime
import pandas as pd
import gspread

#อย่าลืม pip install gunicorn
#Connect Google sheet--------------------------------------------------------
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("My Project 48150-8c433b5a72ae.json", scope)
client = gspread.authorize(creds)

sheet = client.open("SiData+Response")
worksheet = sheet.worksheet("sheet2")
#---------------------------------------------------------------------------


# Flask
app = Flask(__name__)
@app.route('/', methods=['POST']) 

def MainFunction():

    #รับ intent จาก Dailogflow
    question_from_dailogflow_raw = request.get_json(silent=True, force=True)

    #เรียกใช้ฟังก์ชัน generate_answer เพื่อแยกส่วนของคำถาม
    answer_from_bot = generating_answer(question_from_dailogflow_raw)
    
    #ตอบกลับไปที่ Dailogflow
    r = make_response(answer_from_bot)
    r.headers['Content-Type'] = 'application/json' #การตั้งค่าประเภทของข้อมูลที่จะตอบกลับไป

    return r

def generating_answer(question_from_dailogflow_dict):

    #Print intent ที่รับมาจาก Dailogflow
    #print(json.dumps(question_from_dailogflow_dict, indent=4 ,ensure_ascii=False))

    #เก็บต่า ชื่อของ intent ที่รับมาจาก Dailogflow
    intent_group_question_str = question_from_dailogflow_dict["queryResult"]["intent"]["displayName"] 
    #print('intent {}'.format(intent_group_question_str))

    #ลูปตัวเลือกของฟังก์ชั่นสำหรับตอบคำถามกลับ
    if intent_group_question_str == 'บุคลากร - custom - yes':
        #print('question {}'.format(question_from_dailogflow_dict))
        answer_str = personTeam(question_from_dailogflow_dict)
    elif intent_group_question_str == 'สรุปรายงาน SiData+ - custom - yes':
        #print('question {}'.format(question_from_dailogflow_dict))
        answer_str = googlesheet(question_from_dailogflow_dict)
    else: answer_str = "ผมไม่เข้าใจ คุณต้องการอะไร"

    #สร้างการแสดงของ dict 
    answer_from_bot = {"fulfillmentText": answer_str}
    
    #แปลงจาก dict ให้เป็น JSON
    answer_from_bot = json.dumps(answer_from_bot, indent=4) 
    
    return answer_from_bot

def personTeam(respond_dict):
    nam = str(respond_dict["queryResult"]["outputContexts"][1]["parameters"]["Name.original"])
    #print('debug {}'.format(nam))
    if nam == 'เอก':
        answer_function = 'https://si.mahidol.ac.th/siit/admin/personal_images/10011350.jpg'
    elif nam == 'เจมส์':
        answer_function = 'https://si.mahidol.ac.th/siit/admin/personal_images/10034416.jpg'
    elif nam == 'ต้น':
        answer_function = 'https://si.mahidol.ac.th/siit/admin/personal_images/10024137.jpg'
    elif nam == 'คิด' or nam == 'สมคิด':
        answer_function = 'https://si.mahidol.ac.th/siit/admin/personal_images/10011317.jpg'
    elif nam == 'พี่อร':
        answer_function = 'https://si.mahidol.ac.th/siit/admin/personal_images/10003731.jpg'
    elif nam == 'ปับ':
        answer_function = 'https://si.mahidol.ac.th/siit/admin/personal_images/10031115.jpg'
    elif nam == 'หนิง':
        answer_function = 'https://si.mahidol.ac.th/siit/admin/personal_images/10018037.jpg'
    else: answer_function = "ผมไม่สามารถหาข้อมูลให้ได้ครับ ขอโทษด้วยครับ"
    return answer_function

def googlesheet(respond_dict):
    #print('googlesheet {}'.format(respond_dict))
    name1 = str(respond_dict["queryResult"]["outputContexts"][1]["parameters"]["Parameter1.original"])
    #print('name1 = {}'.format(name1))
    dataframe = pd.DataFrame(worksheet.get_all_records())
    answer = ''
    #print('debug {}'.format(name1))
    if name1 == 'วันนี้':
        date_timestamp = (datetime.now()).strftime("%d-%m-%Y")
        df1 = dataframe
        df1 = (df1[df1.Timestamp.str.contains(date_timestamp,case=False)])
        for index,row in df1.iterrows():
            answer1 = "เรื่องที่: {} วันที่: {} \nชื่อผู้แจ้ง: {} \nเบอร์: {}\nรายละเอียด: {}\nเอกสารแนบ: {} \n".format(index+1,row['Timestamp'],row['กรุณากรอก ชื่อ/นามสกุล'],row['เบอร์โทรติดต่อ'],row['ข้อคำถาม'],row['แนบเอกสารเพิ่มเติม'])
            answer = answer + answer1 + '\n'
        answer_function = answer
    elif name1 == 'ทั้งหมด':
        df = dataframe
        for index,row in df.iterrows():
            answer1 = "เรื่องที่: {} วันที่: {} \nชื่อผู้แจ้ง: {} \nเบอร์: {}\nรายละเอียด: {}\nเอกสารแนบ: {} \n".format(index+1,row['Timestamp'],row['กรุณากรอก ชื่อ/นามสกุล'],row['เบอร์โทรติดต่อ'],row['ข้อคำถาม'],row['แนบเอกสารเพิ่มเติม'])
            answer = answer + answer1 + '\n'
        answer_function = answer
    elif name1 == 'สรุป':
        answer_function = 'สรุปรายงานครับ : https://datastudio.google.com/s/vYib4KKo3jM'
    else: answer_function = "ผมไม่สามารถหาข้อมูลให้ได้ครับ ขอโทษด้วยครับ"
    if answer_function == '':
        answer_function = 'ไม่มีข้อความที่ฝากไว้คะ'
    return answer_function

#Flask
if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    print("Starting app on port %d" % port)
    app.run(debug=False, port=port, host='0.0.0.0', threaded=True)
