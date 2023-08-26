import pandas as pd
import openpyxl
import os
import tkinter as tk
from tkinter import simpledialog, messagebox
from collections import defaultdict

# 면접관의 가능한 시간대를 입력받는 대화상자를 표시
def get_interviewer_slots():
    slots = simpledialog.askstring("입력", "면접관의 가능한 시간을 다음과 같은 형식으로 입력해주세요: '월 10시~12시 화 10시~22시 수 10시~17시'")
    return slots

#면접을 진행하기 위한 최소 인원
def get_interviewer_number():
    number=simpledialog.askinteger("입력","최소 면접관 인원을 입력해주세요")
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

def main():
    # tkinter의 기본 윈도우 생성
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기

    # 면접에 필요한 최소 면접관 수 
    interviewer_number= get_interviewer_number()

    #파일경로지정
    dataFile_path="두레박면접시간.xlsx"
    sheet_name1="면접관가능시간"
    sheet_name2="면접자가능시간"
    df1=pd.read_excel(dataFile_path,sheet_name=sheet_name1,engine='openpyxl')
    df2=pd.read_excel(dataFile_path,sheet_name=sheet_name2,engine='openpyxl')


    # 1. 데이터 파싱 및 면접관의 가능한 시간 슬롯 계산
    interviewer_slots={}
    for index, row in df1.iterrows():
        name = row['이름']
        times = row['면접가능시간'].replace(',',' ').split()

        slots = []
        for i in range(0, len(times), 2):
            day = times[i]
            start_time, end_time = map(int, times[i+1].replace("시", "").split("~"))
            
            for hour in range(start_time, end_time):
                for minute in [0, 30]:
                    if hour == end_time and minute == 0:
                        break
                    slot = f"{day} {hour}시 {minute}분"
                    slots.append(slot)

        interviewer_slots[name] = slots

    #시간대별 가능한 면접관들 리스트
    slot_names = defaultdict(list)

    for name, slots in interviewer_slots.items():
        for slot in slots:
            slot_names[slot].append(name)

    valid_slots = {k: v for k, v in slot_names.items() if len(v) >= interviewer_number}
    sorted_valid_slots = sorted(valid_slots.keys(), key=time_sort_key_interviewer)
    print(valid_slots.keys())
    print(valid_slots.values())



    # 2. 참가자의 가능한 시간 슬롯 계산
    interview_slots = {}
    unassigned_names = []
    phone_numbers = {}
    
    for index, row in df2.iterrows():
        name = row['이름']
        print(row['이름'])
        times = row['면접가능시간'].split()
        phone_number=row['전화번호']
        phone_numbers[name] = phone_number

        slots = []
        for i in range(0, len(times), 2):
            day = times[i]
            start_time, end_time = map(int, times[i+1].replace("시", "").split("~"))
            
            for hour in range(start_time, end_time):
                for minute in [0, 30]:
                    if hour == end_time and minute == 0:
                        break
                    slot = f"{day} {hour}시 {minute}분"
                    if slot in valid_slots.keys():  # 면접관의 가능한 시간대와 겹치는지 확인
                        slots.append(slot)
        if not slots:  # slots가 비어있다면
            unassigned_names.append(name)
        
        interview_slots[name] = slots
        print(interview_slots[name])
    if unassigned_names:
        unassigned_str = ', '.join(unassigned_names)
        print(f"다음 참가자들이 면접 시간을 배정받지 못했습니다: {unassigned_str}")
        print("\n면접관이 가능한 시간대를 늘리거나 면접자가 빈 시간대에 가능한지 여부를 체크해 보세요")

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
                break
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

    file_path = '면접시간.xlsx'

    if os.path.exists(file_path):
        answer = input(f"'{file_path}' 파일이 이미 존재합니다. 수정하시겠습니까? (yes or no) ").lower()
        if answer == 'yes' or answer == 'y':
            df_result.to_excel(file_path, index=False, engine='openpyxl')
            print(f"'{file_path}' 파일이 수정되었습니다.")
        else:
            print("저장을 취소했습니다.")
    else:
        df_result.to_excel(file_path, index=False, engine='openpyxl')
        print(f"'{file_path}' 파일이 생성되었습니다.")

        # 작업 완료 메시지 표시
        messagebox.showinfo("알림", "작업이 완료되었습니다!")

if __name__ == "__main__":
    main()


