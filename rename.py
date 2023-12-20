import glob
import os
import warnings
from datetime import timedelta

import pandas as pd
from tqdm import tqdm

# 定义用于匹配Excel文件的模式
EXCEL_PATTERN = "*-[0-9]*-[0-9a-zA-Z]*.xlsx"


def extract_meeting_times(excel_file_path: str) -> tuple:
    """
    从Excel文件中提取会议的首次入会时间和最后退会时间。
    """
    # 读取Excel文件
    df = pd.read_excel(excel_file_path, header=7)
    # 设置列名
    df.columns = df.iloc[0]
    df = df[1:]
    # 转换时间列为datetime类型
    df["首次入会时间"] = pd.to_datetime(df["首次入会时间"])
    df["最后退会时间"] = pd.to_datetime(df["最后退会时间"])

    # 返回最早入会时间、最早和最晚离开时间
    return (
        df["首次入会时间"].min(),
        df["最后退会时间"].min(),
        df["最后退会时间"].max(),
    )


def read_meeting_info_from_excel(excel_file_path: str) -> dict:
    """
    从Excel文件中读取会议的基本信息。
    """
    # 检查文件是否存在
    if not os.path.exists(excel_file_path):
        raise FileNotFoundError(f"Excel file {excel_file_path} not found")

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        # 读取会议主题和会议编号
        df = pd.read_excel(excel_file_path, header=None, engine="openpyxl")
        meeting_theme, meeting_number = df.iloc[0, 1], df.iloc[1, 1]
        # 调用extract_meeting_times函数获取时间信息
        earliest_join, earliest_leave, latest_leave = extract_meeting_times(
            excel_file_path
        )

    # 返回包含会议信息的字典
    return {
        "directory_path": os.path.dirname(excel_file_path),
        "meeting_theme": meeting_theme,
        "meeting_number": meeting_number,
        "earliest_join_time": earliest_join,
        "earliest_leave_time": earliest_leave,
        "latest_leave_time": latest_leave,
        "base_name": f"【{earliest_join.strftime('%Y-%m-%d')}】{meeting_theme}",
        "original_excel_name": excel_file_path,
    }


def get_file_name_patterns(meeting_info: dict, delta_range=range(-90, 90)) -> dict:
    """
    根据会议信息和时间偏移，生成不同文件类型的文件名模式。
    """
    # 检查输入类型
    if not isinstance(meeting_info, dict):
        raise TypeError("meeting_info must be a dictionary")

    patterns = {}
    for key in ["video", "transcription", "summary"]:
        patterns[key] = []
        for delta in delta_range:
            patterns[key].extend(generate_pattern(meeting_info, delta, key))

    meeting_info["patterns"] = patterns
    return meeting_info


def generate_pattern(meeting_info: dict, delta: int, file_type: str) -> list:
    """
    生成特定文件类型的文件名模式。
    """
    time_format = "%Y%m%d%H%M%S"
    if file_type == "video":
        time_stamp = (
            meeting_info["earliest_join_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        return [f"TM-{time_stamp}-{meeting_info['meeting_number']}-*.mp4"]
    elif file_type == "transcription":
        time_stamp_1 = (
            meeting_info["earliest_leave_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        time_stamp_2 = (
            meeting_info["latest_leave_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        return [
            f"TencentMeeting_({time_stamp_1})_Transcription.txt",
            f"TencentMeeting_({time_stamp_2})_Transcription.txt",
        ]
    elif file_type == "summary":
        time_stamp_1 = (
            meeting_info["earliest_leave_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        time_stamp_2 = (
            meeting_info["latest_leave_time"] + timedelta(seconds=delta)
        ).strftime(time_format)
        return [
            f"TencentMeeting_{time_stamp_1}_Summary.txt",
            f"TencentMeeting_{time_stamp_2}_Summary.txt",
        ]


def find_matching_files(directory: str, patterns: list) -> list:
    """
    在指定目录下查找匹配给定模式的文件。
    """
    # 检查目录是否存在
    if not os.path.isdir(directory):
        raise FileNotFoundError(f"Directory {directory} not found")

    # 查找匹配的文件并返回
    return list(
        {
            file
            for pattern in patterns
            for file in glob.glob(os.path.join(directory, pattern))
        }
    )


def rename_files(meeting_info: dict) -> None:
    """
    根据会议信息重命名文件。
    """
    # 检查输入类型
    if not isinstance(meeting_info, dict):
        raise TypeError("meeting_info must be a dictionary")

    # 获取目录和基础文件名
    directory, base_name = meeting_info.get("directory_path"), meeting_info.get(
        "base_name"
    )
    if not directory or not base_name:
        raise ValueError("Invalid meeting information provided")

    # 文件类型和后缀名
    file_types = [
        "original_excel_name",
        "original_video_name",
        "original_transcription_name",
        "original_summary_name",
    ]
    suffixes = [".xlsx", ".mp4", "_Transcription.txt", "_Summary.txt"]

    # 重命名文件
    for file_type, suffix in zip(file_types, suffixes):
        original_file = meeting_info.get(file_type)
        if original_file and os.path.exists(original_file):
            new_name = os.path.join(directory, base_name + suffix)
            os.rename(original_file, new_name)


def process_meetings(directory_path: str):
    """
    处理指定目录下的会议文件。
    """
    # 检查目录是否存在
    if not os.path.isdir(directory_path):
        raise FileNotFoundError(f"Directory {directory_path} not found")

    # 查找匹配的Excel文件
    excel_files = find_matching_files(directory_path, [EXCEL_PATTERN])
    if not excel_files:
        print("No Excel files found.")
        return

    # 遍历并处理每个Excel文件
    for excel_file in tqdm(excel_files, desc="Processing files"):
        try:
            meeting_info = read_meeting_info_from_excel(excel_file)
            meeting_info = get_file_name_patterns(meeting_info)

            # 对于每种文件类型，查找并重命名文件
            for file_type in ["video", "transcription", "summary"]:
                files = find_matching_files(
                    directory_path, meeting_info["patterns"][file_type]
                )
                if len(files) == 1:
                    meeting_info[f"original_{file_type}_name"] = files[0]

            rename_files(meeting_info)
        except Exception as e:
            print(f"Error processing file {excel_file}: {e}")
