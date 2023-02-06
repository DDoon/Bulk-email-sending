# 이메일 대량발송

속성: 이메일 업무자동화
작성일시: 2023년 2월 6일 오후 5:36
최종 편집일시: 2023년 2월 6일 오후 5:58

# # 개요

---

- 이메일을 대량으로 보내는 경우 사용(이메일마다 보내는 파일이 다를 때 사용)

# # 파일 구조

---

1. **파일 구조**
    - **account.py**: 자신의 이메일 아이디와 비밀번호 입력하는 파일
    - **list32.xlsx**: 데이터가 담긴 엑셀파일(받는사람, 메일제목, 메일내용, 첨부파일)
    - **mail.py**: 데이터가 담긴 엑셀파일(받는사람, 메일제목, 메일내용, 첨부파일)

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled.png)

1. **account.py: 자신의 이메일 아이디와 비밀번호 입력**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%201.png)

1. **mail.py: 이메일 보내는 파일** 

```python
from account import MY_ID, MY_PW
from email.message import EmailMessage
from smtplib import SMTP_SSL
from pathlib import Path
from openpyxl import load_workbook

def send_mail(받는사람, 제목, 본문, 첨부파일=False):

    # 템플릿 생성
    msg = EmailMessage()

    # 보내는 사람 / 받는 사람 / 제목 입력
    msg["From"] = MY_ID
    msg["To"] = 받는사람.split()
    msg["Subject"] = 제목

    # 본문 구성
    msg.set_content(본문)
    
    # 파일 첨부
    if 첨부파일:
        파일명 = Path(첨부파일).name
        with open(첨부파일, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=파일명)
        
    with SMTP_SSL("smtp.naver.com", 465) as smtp:
        smtp.login(MY_ID, MY_PW)
        smtp.send_message(msg)
    
    # 완료 메시지
    print(받는사람, "성공", sep="\t")

        
# 엑셀파일에서 가져온 정보를 활용해 함수 반복 실행
wb = load_workbook('list32.xlsx', data_only=True)
ws = wb.active

for 행 in ws.iter_rows(min_row=2):
    받는사람 = 행[0].value
    제목 = 행[1].value
    본문 = 행[2].value
    첨부파일 = 행[3].value

    send_mail(받는사람, 제목, 본문, 첨부파일)
```

1. **엑셀파일.xlsx: 데이터가 담긴 엑셀파일(받는사람, 메일제목, 메일내용, 첨부파일)**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%202.png)

# # 네이버 환경설정

---

1. 네이버 메일에 들어가 **왼쪽 하단에 환경설정 클릭**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%203.png)

1. 환경설정 상단에 **POP3/IMAP 설정 클릭**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%204.png)

1. **POP3/IMAP 설정 클릭 → IMAP/SMTP 설정 클릭 → IMAP/SMTP 사용에서 사용함 클릭 → 확인**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%205.png)

# # 메일 보내기(코드실행 하기)

---

1. 엑셀 파일 설정

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%206.png)

1. VScode **오른쪽 상단에 실행버튼**이 있음 클릭

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%207.png)

1. 밑에 터미널 창에서 보낸거 확인 가능

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%208.png)

- **Done : 보내기까지 완료시간**
- **Running: 파이썬 경로**
- **~@naver.com: 보낸 메일주소**
- **clear: 보낸메일확인문구**

1. **메일 확인**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%209.png)

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2010.png)

- 위와 같이 엑셀에 있는 데이터로 메일제목, 첨부파일, 메일 내용 보내기 가능

# # 첨부파일 여러개 보내기

---

- 한 사람에게 첨부파일을 여러개 보내야하는 상황 발생
1. 엑셀파일에 첨부파일2 를 추가

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2011.png)

1. 코드 추가

```python
from account import MY_ID, MY_PW
from email.message import EmailMessage
from smtplib import SMTP_SSL
from pathlib import Path
from openpyxl import load_workbook

def send_mail(받는사람, 제목, 본문, 첨부파일, 첨부파일2=False):

    # 템플릿 생성
    msg = EmailMessage()

    # 보내는 사람 / 받는 사람 / 제목 입력
    msg["From"] = MY_ID
    msg["To"] = 받는사람.split()
    msg["Subject"] = 제목

    # 본문 구성
    msg.set_content(본문)
    
    # 파일 첨부
    if 첨부파일:
        파일명 = Path(첨부파일).name
        with open(첨부파일, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=파일명)
    if 첨부파일:
        파일명 = Path(첨부파일2).name
        with open(첨부파일2, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=파일명)

        
    with SMTP_SSL("smtp.naver.com", 465) as smtp:
        smtp.login(MY_ID, MY_PW)
        smtp.send_message(msg)
    
    # 완료 메시지
    print(받는사람, "성공", sep="\t")

        
# 엑셀파일에서 가져온 정보를 활용해 함수 반복 실행
wb = load_workbook('list32.xlsx', data_only=True)
ws = wb.active

for 행 in ws.iter_rows(min_row=2):
    받는사람 = 행[0].value
    제목 = 행[1].value
    본문 = 행[2].value
    첨부파일 = 행[3].value
    첨부파일2 = 행[4].value

    
    send_mail(받는사람, 제목, 본문, 첨부파일, 첨부파일2)
```

- def send_mail에 첨부파일2 작성

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2012.png)

- 파일 첨부부분에 if문(첨부파일2) 작성

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2013.png)

- for문에 첨부파일2 작성

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2014.png)

- **위와 같이 작성한 후에 저장하기!!!!**
- 그리고 실행해주면 한 사람에게 첨부파일 여러개가 들어가진다.

# # 이슈 및 해결방안(★★★★★)

---

1. **파일명(띄어쓰기 금지)**
    - 보내는 파일의 파일명에 띄어쓰기가 있으면 안보내짐. ex) 프로젝트 보고서 최종.pdf
    - **띄어쓰기 대신에 _을 추가해줌 ex) 프로젝트_보고서_최종.pdf**

1. **엑셀에 파일명을 쉽게 작성하는 방법**
    - 파일명이 다를 수가 있다.
    - ex) 수료증, 상장 등의 파일은 그 사람의 이름이 들어가 있기 때문
    - 파일명 하나하나 엑셀에 작성하면 불편함

- **검색창에 cmd 입력 → 명령 프롬포트 실행**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2015.png)

- 프롬포트 실행 → 파일이 담겨있는 디렉토리(폴더)경로 설정
- cd : 경로 변경
- C ~ : 변경 경로

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2016.png)

- **dir/b>test.txt 입력후 엔터**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2017.png)

- 엔터해주면 아까 입력한 경로에  파일명이 담겨있는 .txt파일이 생김
- **텍스트파일 실행후 파일명을 엑셀에 복붙해주면 끝!**

![Untitled](%E1%84%8B%E1%85%B5%E1%84%86%E1%85%A6%E1%84%8B%E1%85%B5%E1%86%AF%20%E1%84%83%E1%85%A2%E1%84%85%E1%85%A3%E1%86%BC%E1%84%87%E1%85%A1%E1%86%AF%E1%84%89%E1%85%A9%E1%86%BC%20993cc652434e40f4bd1cd4430234d6ab/Untitled%2018.png)

# # 참고사이트

---

[https://dataanalytics.tistory.com/176](https://dataanalytics.tistory.com/176)

[https://github.com/hleecaster/automation-in-python/blob/main/5. 메일 보내기.ipynb](https://github.com/hleecaster/automation-in-python/blob/main/5.%20%EB%A9%94%EC%9D%BC%20%EB%B3%B4%EB%82%B4%EA%B8%B0.ipynb)

[https://www.inflearn.com/course/나도코딩-업무자동화-파이썬](https://www.inflearn.com/course/%EB%82%98%EB%8F%84%EC%BD%94%EB%94%A9-%EC%97%85%EB%AC%B4%EC%9E%90%EB%8F%99%ED%99%94-%ED%8C%8C%EC%9D%B4%EC%8D%AC)