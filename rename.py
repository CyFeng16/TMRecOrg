import glob
import os
from datetime import timedelta
import pandas as pd

EXCEL_PATTERN = "*-[0-9]*-[0-9a-zA-Z]*.xlsx"


def extract_meeting_times(excel_file_path: str) -> tuple:
    df = pd.read_excel(excel_file_path, header=7)
    df.columns = df.iloc[0]
    df = df[1:]
    df["首次入会时间"] = pd.to_datetime(df["首次入会时间"])
    df["最后退会时间"] = pd.to_datetime(df["最后退会时间"])

    return df["首次入会时间"].min(), df["最后退会时间"].max()


def read_meeting_info_from_excel(excel_file_path: str) -> dict:
    if not os.path.exists(excel_file_path):
        raise FileNotFoundError(f"Excel file {excel_file_path} not found")

    df = pd.read_excel(excel_file_path, header=None)
    meeting_theme, meeting_number = df.iloc[0, 1], df.iloc[1, 1]
    earliest_join, latest_leave = extract_meeting_times(excel_file_path)

    return {
        "directory_path": os.path.dirname(excel_file_path),
        "meeting_theme": meeting_theme,
        "meeting_number": meeting_number,
        "earliest_join_time": earliest_join,
        "latest_leave_time": latest_leave,
        "base_name": f"【{earliest_join.strftime('%Y-%m-%d')}】{meeting_theme}",
        "original_excel_name": excel_file_path,
    }


def get_file_name_patterns(meeting_info: dict, delta_range=range(-5, 6)) -> dict:
    if not isinstance(meeting_info, dict):
        raise TypeError("meeting_info must be a dictionary")

    patterns = {}
    for key in ["video", "transcription", "summary"]:
        patterns[key] = [
            generate_pattern(meeting_info, delta, key) for delta in delta_range
        ]

    meeting_info["patterns"] = patterns
    return meeting_info


def generate_pattern(meeting_info: dict, delta: int, file_type: str) -> str:
    time_format = "%Y%m%d%H%M%S"
    if file_type == "video":
        time_stamp = (
            meeting_info["earliest_join_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        return f"TM-{time_stamp}-{meeting_info['meeting_number']}-*.mp4"
    elif file_type == "transcription":
        time_stamp = (
            meeting_info["latest_leave_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        return f"TencentMeeting_({time_stamp})_Transcription.txt"
    elif file_type == "summary":
        time_stamp = (
            meeting_info["latest_leave_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        return f"TencentMeeting_{time_stamp}_Summary.txt"


def find_matching_files(directory: str, patterns: list) -> list:
    if not os.path.isdir(directory):
        raise FileNotFoundError(f"Directory {directory} not found")

    return list(
        {
            file
            for pattern in patterns
            for file in glob.glob(os.path.join(directory, pattern))
        }
    )


def rename_files(meeting_info: dict) -> None:
    if not isinstance(meeting_info, dict):
        raise TypeError("meeting_info must be a dictionary")

    directory, base_name = meeting_info.get("directory_path"), meeting_info.get(
        "base_name"
    )
    if not directory or not base_name:
        raise ValueError("Invalid meeting information provided")

    file_types = [
        "original_excel_name",
        "original_video_name",
        "original_transcription_name",
        "original_summary_name",
    ]
    suffixes = [".xlsx", ".mp4", "_Transcription.txt", "_Summary.txt"]

    for file_type, suffix in zip(file_types, suffixes):
        original_file = meeting_info.get(file_type)
        if original_file and os.path.exists(original_file):
            new_name = os.path.join(directory, base_name + suffix)
            os.rename(original_file, new_name)


def process_meetings(directory_path: str):
    if not os.path.isdir(directory_path):
        raise FileNotFoundError(f"Directory {directory_path} not found")

    excel_files = find_matching_files(directory_path, [EXCEL_PATTERN])
    if not excel_files:
        print("No Excel files found.")
        return

    for excel_file in excel_files:
        try:
            meeting_info = read_meeting_info_from_excel(excel_file)
            meeting_info = get_file_name_patterns(meeting_info)

            for file_type in ["video", "transcription", "summary"]:
                files = find_matching_files(
                    directory_path, meeting_info["patterns"][file_type]
                )
                if len(files) == 1:
                    meeting_info[f"original_{file_type}_name"] = files[0]

            rename_files(meeting_info)
        except Exception as e:
            print(f"Error processing file {excel_file}: {e}")
