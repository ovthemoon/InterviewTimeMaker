import pandas as pd
import openpyxl, sys, os
import tkinter as tk
from tkinter import simpledialog, messagebox,filedialog,ttk,Text,Scrollbar
from collections import defaultdict

    
#데이터로 사용할 엑셀 파일 경로 가져오기
def get_file_path():
    root = tk.Tk()
    root.withdraw()  # GUI 창을 보이지 않게 합니다.
    file_path = filedialog.askopenfilename(title="파일을 선택하세요", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    if not file_path.endswith(".xlsx"):
        tk.messagebox.showerror("오류","잘못된 파일 형식입니다.엑셀 파일만 등록 가능합니다.")
        return None
    return file_path

#저장할 파일 경로 가져오기
def get_save_path():
    root = tk.Tk()
    root.withdraw()  # GUI 창을 숨김
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files","*.xlsx"), ("All files", "*.*")])
    if not file_path.endswith(".xlsx"):
        tk.messagebox.showerror("오류","잘못된 파일 형식입니다.엑셀 파일만 등록 가능합니다.")
        return None
    return file_path

# 면접관의 가능한 시간대를 입력받는 대화상자를 표시
def get_interviewer_slots():
    slots = simpledialog.askstring("입력", "면접관의 가능한 시간을 다음과 같은 형식으로 입력해주세요: '월 10시~12시 화 10시~22시 수 10시~17시'")
    return slots

#면접을 진행하기 위한 최소 인원
def get_interviewer_number():
    number=simpledialog.askinteger("입력","최소 면접관 인원을 입력해주세요")
    if not number:
        return 1
    return number

#interviewer 오름차순정렬 규칙
def time_sort_key_interviewer(item):
    days = ['월', '화', '수', '목', '금', '토', '일']
    
    parts = item.split()
    day = parts[0]
    hour = parts[1].replace("시", "")
    minute = parts[2].replace("분", "")
    
    return days.index(day), int(hour), int(minute)

#interviewee 오름차순정렬 규칙
def time_sort_key_interviewee(item):
    days = ['월', '화', '수', '목', '금', '토', '일']
    
    parts = item[1].split()
    day = parts[0]
    hour = parts[1].replace("시", "")
    minute = parts[2].replace("분", "")
    
    return days.index(day), int(hour), int(minute)

#시간 데이터 파싱 함수
def extract_time_slots(row, valid_slots=None):
    times = row['면접가능시간'].replace(',', ' ').split()
    slots = []
    
    for i in range(0, len(times), 2):
        day = times[i]
        start_time, end_time = map(int, times[i+1].replace("시", "").split("~"))
        
        for hour in range(start_time, end_time):
            for minute in [0, 30]:
                if hour == end_time and minute == 0:
                    break
                slot = f"{day} {hour}시 {minute}분"
                
                # valid_slots 인자가 제공되면, 해당 슬롯이 valid_slots에 있는지 확인
                if valid_slots:
                    if slot in valid_slots:
                        slots.append(slot)
                else:
                    slots.append(slot)
                    
    return slots

#작성된 데이터 미리보기 함수
def show_dataframe(df):
    root = tk.Tk()
    root.title("DataFrame Viewer")

    frame = ttk.Frame(root)
    frame.grid(row=0, column=0, sticky='nsew')

    text_widget = Text(frame, wrap=tk.NONE)
    text_widget.insert(tk.END, df.to_string())
    text_widget.config(state=tk.DISABLED)

    y_scrollbar = Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
    y_scrollbar.grid(row=0, column=1, sticky='ns')
    text_widget.config(yscrollcommand=y_scrollbar.set)

    x_scrollbar = Scrollbar(frame, orient=tk.HORIZONTAL, command=text_widget.xview)
    x_scrollbar.grid(row=1, column=0, sticky='ew')
    text_widget.config(xscrollcommand=x_scrollbar.set)

    text_widget.grid(row=0, column=0, sticky='nsew')
    
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)


def main():
    # tkinter의 기본 윈도우 생성
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기

    # 면접에 필요한 최소 면접관 수 
    interviewer_number= get_interviewer_number()

    dataFile_path = get_file_path()  # 사용자가 선택한 파일 경로를 얻습니다.
    if not dataFile_path:  # 파일을 선택하지 않고 취소를 누른 경우
        print("파일을 선택하지 않았습니다. 프로그램을 종료합니다.")
        return
    sheet_name1="면접관가능시간"
    sheet_name2="면접자가능시간"
    df1=pd.read_excel(dataFile_path,sheet_name=sheet_name1,engine='openpyxl')
    df2=pd.read_excel(dataFile_path,sheet_name=sheet_name2,engine='openpyxl')


    # 1. 데이터 파싱 및 면접관의 가능한 시간 슬롯 계산
    interviewer_slots={}
    for index, row in df1.iterrows():
        name = row['이름']
        interviewer_slots[name] = extract_time_slots(row)
        
    #시간대별 가능한 면접관들 리스트
    slot_names = defaultdict(list)

    for name, slots in interviewer_slots.items():
        for slot in slots:
            slot_names[slot].append(name)

    valid_slots = {k: v for k, v in slot_names.items() if len(v) >= interviewer_number}
    sorted_valid_slots = sorted(valid_slots.keys(), key=time_sort_key_interviewer)
    ##print(valid_slots.keys())
    ##print(valid_slots.values())



    # 2. 참가자의 가능한 시간 슬롯 계산
    interview_slots = {}
    unassigned_names = []
    phone_numbers = {}
    
    for index, row in df2.iterrows():
        name = row['이름']
        phone_number = row['전화번호']
        phone_numbers[name] = phone_number
        interview_slots[name] = extract_time_slots(row, valid_slots)

        if not interview_slots[name]:
            unassigned_names.append(name)
        
    
        

    # 3. 가능한 시간 슬롯의 개수를 기반으로 참가자 정렬
    sorted_names = sorted(interview_slots.keys(), key=lambda k: len(interview_slots[k]))

    # 4. 면접 스케줄 만들기
    interview_schedule = {}
    used_slots = set()

    for name in sorted_names:
        slots = interview_slots[name]
        for slot in slots:
            if slot not in used_slots:
                interview_schedule[name] = slot
                used_slots.add(slot)
                if name in interview_slots:
                    del interview_slots[name]
                break

    unassigned_names.extend(interview_slots.keys())
    
    if unassigned_names:
        unassigned_str = ', '.join(unassigned_names)
        tk.messagebox.showwarning("면접 시간 미배정",f"다음 참가자들이 면접 시간을 배정받지 못했습니다: {unassigned_str}")
    # 5. 결과 출력(프로그램내 확인)(오름차순)
    sorted_schedule = sorted(interview_schedule.items(), key=time_sort_key_interviewee)

    #그 시간대에 가능한 면접관 연결
    for name, slot in sorted_schedule:
        available_interviewers = slot_names.get(slot, [])
        interviewers_str = ', '.join(available_interviewers)
        phone_number = phone_numbers.get(name, '번호 없음')
        print(f"{name} ({phone_number}): {slot} (가능한 면접관: {interviewers_str})")

    #6. 결과를 엑셀 파일로 저장
    result_data = []

    for name, slot in sorted_schedule:
        available_interviewers = slot_names.get(slot, [])
        interviewers_str = ', '.join(available_interviewers)
        phone_number = phone_numbers.get(name, '번호 없음')
        result_data.append([name, phone_number, slot, interviewers_str])
        
    for name in unassigned_names:
        phone_number = phone_numbers.get(name, '번호 없음')
        result_data.append([name, phone_number, "배정 안됨", ""])

    df_result = pd.DataFrame(result_data, columns=['이름', '전화번호', '면접 시간', '가능한 면접관'])
    print(df_result)
    #show_dataframe(df_result)

    #저장 경로 받아오기
    file_path = get_save_path()

    if not file_path:  # 사용자가 '저장하기' 창에서 취소를 누를 경우
        print("저장을 취소했습니다.")
        tk.messagebox.showinfo("저장 취소", "저장을 취소했습니다.")
    else:
        df_result.to_excel(file_path, index=False, engine='openpyxl')
        print(f"'{file_path}' 파일이 저장되었습니다.")
        tk.messagebox.showinfo("파일 저장", "파일이 저장되었습니다.")


if __name__ == "__main__":
    main()


