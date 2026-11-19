import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
def get_file_path():
    root = tk.Tk()
    # Function to get the input Excel file path
    root.withdraw()  # Hide the main GUI window
    file_path = filedialog.askopenfilename(
        title="파일을 선택하세요",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if file_path and not file_path.endswith(".xlsx"):
        messagebox.showerror("오류", "잘못된 파일 형식입니다. 엑셀 파일만 등록 가능합니다.")
        return None
    return file_path

def save_file_path():
    root = tk.Tk()
    # Function to get the output Excel file path
    root.withdraw()  # Hide the main GUI window
    file_path = filedialog.asksaveasfilename(
        title="저장할 파일 경로를 선택하세요",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    return file_path
def input_interviewer_mincount():
    root = tk.Tk()
    # 정수 입력받기
    num = simpledialog.askinteger("입력", "최소 면접관 수를 입력하세요:",
                                  parent=root,
                                  minvalue=1)  # 최소값 1로 설정

    # 입력된 정수를 반환
    return num
    
def parse_time_slots(df):
    # Function to parse time slots from the dataframe
    parsed_slots = []  # List to store the parsed time slots
    availability_counts = {}  # Dictionary to store the count of available slots per candidate

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        candidate_name = row['1. 성함']
        availability_counts[candidate_name] = 0
        # Iterate through each column that represents a time slot
        for col in df.columns[7:]:  # Time slot columns start from the 8th column in the dataframe
            # Check if the cell is not NaN (meaning the person is available at this time)
            if pd.notna(row[col]):
                # Extract the time range from the column header and split it into start and end times
                time_range = col[col.index('[')+1: col.index(']')].split('~')
                start_time = [int(x) for x in time_range[0].split(':')]
                end_time = [int(x) for x in time_range[1].split(':')]

                # Calculate the total minutes for the start and end times
                start_minutes = start_time[0] * 60 + start_time[1]
                end_minutes = end_time[0] * 60 + end_time[1]

                # Generate time slots in 15-minute increments
                time_slots = [f"{(minutes // 60):02d}:{(minutes % 60):02d}"
                              for minutes in range(start_minutes, end_minutes, 15)]

                # For each available day, create an entry in the parsed_slots list
                for day in row[col].split(', '):
                    for slot in time_slots:
                        parsed_slots.append({'Name': candidate_name, 'Day': day.strip(), 'Time': slot})
                        availability_counts[candidate_name] += 1

    # Convert the list of dictionaries to a DataFrame for easier manipulation
    parsed_slots_df = pd.DataFrame(parsed_slots)

    # Add a new column to the dataframe for the available slot counts
    parsed_slots_df['AvailableSlotsCount'] = parsed_slots_df['Name'].map(availability_counts)

    # Sort the dataframe based on the available slot counts, then Day, and then Name
    days = ['월', '화', '수', '목', '금', '토', '일']
    parsed_slots_df['Day'] = pd.Categorical(parsed_slots_df['Day'], categories=days, ordered=True)
    parsed_slots_df.sort_values(by=['AvailableSlotsCount', 'Day', 'Name'], ascending=[False, True, True], inplace=True)

    return parsed_slots_df


def parse_interviewer_slots(df):
    parsed_slots = []
    
    # Iterate over the DataFrame rows
    for index, row in df.iterrows():
        interviewer = row['면접관']
        
        # Iterate over each time slot column, skipping the interviewer's name
        for time_slot in df.columns[1:]:
            if pd.notna(row[time_slot]):
                # Split the cell by ', ' and take only the day part (ignoring the date part)
                days = [day.split(' ')[0] for day in row[time_slot].split(', ')]
                
                # Extract the time range from the column name
                time_range = time_slot.split(' ')[-1].strip('[]').split('~')
                start_time_str, end_time_str = time_range
                start_hour, start_minute = map(int, start_time_str.split(':'))
                end_hour, end_minute = map(int, end_time_str.split(':'))

                # Calculate the total minutes for the start and end times
                start_minutes = start_hour * 60 + start_minute
                end_minutes = end_hour * 60 + end_minute

                # Generate time slots in 15-minute increments
                time_slots = [f"{(minutes // 60):02d}:{(minutes % 60):02d}"
                              for minutes in range(start_minutes, end_minutes, 15)]

                # For each day and each time slot, create an entry
                for day in days:
                    # Strip the day to remove any whitespace and ignore dates
                    day = day.strip()
                    for slot in time_slots:
                        parsed_slots.append({
                            'Interviewer': interviewer,
                            'Day': day,
                            'Time': slot
                        })
    
    # Convert the list of dictionaries to a DataFrame
    return pd.DataFrame(parsed_slots)

def filter_available_times(parsed_slots_df, min_interviewers):
    day_order = ['월', '화', '수', '목', '금']
    # 요일별, 시간대별로 가능한 면접관 수를 계산합니다.
    counts = parsed_slots_df.groupby(['Day', 'Time'])['Interviewer'].size().reset_index(name='Count')
    
    # 최소 면접관 수를 충족하는 시간대만 필터링합니다.
    available_times = counts[counts['Count'] >= min_interviewers]
    
    # 해당 시간대에 가능한 면접관의 이름을 리스트로 만듭니다.
    parsed_slots_df['InterviewerList'] = parsed_slots_df.groupby(['Day', 'Time'])['Interviewer'].transform(lambda x: ', '.join(x))
    
    # 가능한 면접관의 이름을 추가합니다.
    available_times = available_times.merge(parsed_slots_df[['Day', 'Time', 'InterviewerList']].drop_duplicates(), on=['Day', 'Time'], how='left')

    # 반환된 available_times에 요일을 Categorical 데이터로 변환합니다.
    available_times['Day'] = pd.Categorical(available_times['Day'], categories=day_order, ordered=True)

    # 요일과 시간대별로 정렬합니다.
    available_times.sort_values(by=['Day', 'Time'], inplace=True)
    
    return available_times
def assign_interview_slots_and_list_failures(candidate_slots_df, available_times_df):
    # 배정 시도를 위한 임시 리스트
    assignments_list = []
    failed_assignments_list = []

    candidate_names = candidate_slots_df['Name'].unique()
    print('배정시도')
    
    for candidate_name in candidate_names:
        print(candidate_name)
        candidate_slots = candidate_slots_df[candidate_slots_df['Name'] == candidate_name]
        assigned = False

        for _, slot in candidate_slots.iterrows():
            available_time = available_times_df[
                (available_times_df['Day'] == slot['Day']) & 
                (available_times_df['Time'] == slot['Time'])
            ]
            
            # 해당 시간대에 면접관이 충분한 경우 배정
            if not available_time.empty:
                assignments_list.append({
                    'Name': slot['Name'], 
                    'Day': slot['Day'], 
                    'Time': slot['Time'], 
                    'Interviewer': available_time.iloc[0]['InterviewerList']
                })
                
                assigned = True
                
                available_times_df = available_times_df.drop(available_time.index)
                break  # 참가자당 한 시간만 배정

        if not assigned:
            failed_assignments_list.append({'Name': candidate_name})

    # 배정된 시간과 실패한 참가자를 DataFrame으로 변환
    assigned_slots = pd.DataFrame(assignments_list)
    failed_assignments = pd.DataFrame(failed_assignments_list)

     # 요일을 순서대로 정렬하기 위해 Categorical 타입을 사용
    days_order = ['월', '화', '수', '목', '금', '토', '일']
    assigned_slots['Day'] = pd.Categorical(assigned_slots['Day'], categories=days_order, ordered=True)
    
    # 요일과 시간대별로 정렬
    assigned_slots.sort_values(by=['Day', 'Time'], inplace=True)

    return assigned_slots, failed_assignments





def save_to_excel(parsed_candidates_df,parsed_interviewers_df,interview_time_df,assigned_slots_df, failed_assignments_df,output_file_path):
    with pd.ExcelWriter(output_file_path) as writer:
        parsed_candidates_df.to_excel(writer,sheet_name='1.면접자가능시간')
        parsed_interviewers_df.to_excel(writer,sheet_name='2.면접관가능시간')
        interview_time_df.to_excel(writer,sheet_name='3.면접가능시간')
        assigned_slots_df.to_excel(writer,sheet_name='4.면접시간배정')
        failed_assignments_df.to_excel(writer, sheet_name='5.배정실패')
    print(f"The file was saved as {output_file_path}")

def main():
    root = tk.Tk()
    root.withdraw()
    # Main function to run the script
    messagebox.showinfo("파일 선택","면접 시간 데이터베이스를 선택해 주세요.")
    input_file_path = get_file_path()  # Get the input file path
    print(input_file_path)
    if not input_file_path:
        messagebox.showerror("파일 선택 오류","파일 경로가 잘못되었습니다. 프로그램을 종료합니다.")
        return
    messagebox.showinfo("경로 선택","파일을 저장할 경로를 선택해 주세요")
    output_file_path = save_file_path()  # Get the output file path
    if not output_file_path:
        messagebox.showerror("경로 선택 오류","경로가 잘못되었습니다. 프로그램을 종료합니다.")
        return
    try:

        # Load the data from the Excel file
        df_candidates = pd.read_excel(input_file_path, sheet_name='면접자가능시간')
        parsed_candidates_df = parse_time_slots(df_candidates)

        df_interviewers = pd.read_excel(input_file_path, sheet_name='면접관가능시간')
        parsed_interviewers_df = parse_interviewer_slots(df_interviewers)
        
        min_interviewers = input_interviewer_mincount()
        
        filtered_available_times_df = filter_available_times(parsed_interviewers_df, min_interviewers)

        assigned_slots, failed_assignments = assign_interview_slots_and_list_failures(parsed_candidates_df, filtered_available_times_df)

        # Save to Excel file
        save_to_excel(parsed_candidates_df, parsed_interviewers_df, filtered_available_times_df, assigned_slots,failed_assignments, output_file_path)
        messagebox.showinfo("완료","면접 시간 배정이 완료되었습니다. 결과를 확인해 주세요.")
        print("배정에 실패한 참가자:", failed_assignments)
    except Exception as e:
        mesagebox.showerror("오류","오류가 발생했습니다: "+ste(e))
    root.destroy()

if __name__ == "__main__":
    main()
