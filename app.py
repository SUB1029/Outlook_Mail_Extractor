from flask import Flask, request, send_file, render_template
import win32com.client
import datetime
import re
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract():
    start_date = request.form['start_date']
    end_date = request.form['end_date']

    start = datetime.datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.datetime.strptime(end_date, "%Y-%m-%d") + datetime.timedelta(hours=23, minutes=59, seconds=59)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 받은 편지함
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    results = []

    for message in messages:
        try:
            if message.Class != 43:
                continue

            subject = message.Subject
            received_time = message.ReceivedTime
            body = message.Body

            if '[Submitted]' in subject and start <= received_time <= end:
                match = re.search(r"Created by\s*:\s*(.*)", body)
                if match:
                    created_line = match.group(1).strip()
                    if re.search(r"(Part|part|PART|팀|IT)", created_line):
                        results.append({
                            "No": len(results)+1,
                            "요청부서": created_line,
                            "요청월": received_time.strftime("%Y-%m"),
                            "요청일자": received_time.strftime("%Y-%m-%d"),
                            # 필요한 항목은 여기서 확장 가능
                        })
        except Exception as e:
            print("Error:", e)

    df = pd.DataFrame(results)
    output_path = "output.xlsx"
    df.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
