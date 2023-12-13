import os
import pandas as pd
import glob
from datetime import datetime, timedelta

EXCEL_PATTERN = "*-[0-9]*-[0-9a-zA-Z]*.xlsx"


def read_meeting_info_from_excel(excel_file_path):
    """
    Read the meeting information from the given Excel file.
    """
    df = pd.read_excel(excel_file_path, header=None)
    meeting_theme = df.iloc[0, 1]
    meeting_number = df.iloc[1, 1]
    scheduled_start_time = datetime.strptime(df.iloc[3, 1], "%Y-%m-%d %H:%M:%S")
    duration_str = df.iloc[5, 1]
    duration_parts = duration_str.split(":")
    duration = timedelta(
        hours=int(duration_parts[0]),
        minutes=int(duration_parts[1]),
        seconds=int(duration_parts[2]),
    )
    end_time = scheduled_start_time + duration

    return {
        "meeting_theme": meeting_theme,
        "meeting_number": meeting_number,
        "scheduled_start_time": scheduled_start_time,
        "duration": duration,
        "end_time": end_time,
    }


def rename_meeting_files_with_flexibility(directory, meeting_info, time_flexibility=5):
    """
    Rename files related to a specific meeting in the given directory with a flexibility in time matching.
    """
    renamed_files = []
    new_file_base = f"【{meeting_info['scheduled_start_time'].strftime('%Y-%m-%d')}】{meeting_info['meeting_theme']}"
    print("New file base:", new_file_base)

    for file in os.listdir(directory):

        if meeting_info["meeting_number"] in file:
            file_timestamp_str = "".join(filter(str.isdigit, file.split("-")[0]))
            try:
                file_timestamp = datetime.strptime(file_timestamp_str, "%Y%m%d%H%M%S")
            except ValueError:
                continue

            actual_end_time = (
                meeting_info["scheduled_start_time"] + meeting_info["duration"]
            )
            time_difference_start = abs(
                (file_timestamp - meeting_info["scheduled_start_time"]).total_seconds()
            )
            time_difference_end = abs(
                (file_timestamp - actual_end_time).total_seconds()
            )

            if (
                time_difference_start <= time_flexibility
                or time_difference_end <= time_flexibility
            ):
                old_file_path = os.path.join(directory, file)
                file_extension = os.path.splitext(file)[1]
                new_file_name = new_file_base + file_extension
                new_file_path = os.path.join(directory, new_file_name)

                print(f"Renaming {old_file_path} to {new_file_path}")
                # os.rename(old_file_path, new_file_path)
                # renamed_files.append(new_file_name)

    return renamed_files


def find_matching_video_file(directory, meeting_info):
    start_time_str = meeting_info["scheduled_start_time"].strftime("%Y%m%d%H%M%S")
    meeting_number = meeting_info["meeting_number"]
    video_pattern = f"TM-{start_time_str}-{meeting_number}-*.mp4"

    matched_files = glob.glob(os.path.join(directory, video_pattern))

    if len(matched_files) != 1:
        print(f"Found {len(matched_files)} video files")
        return None

    return matched_files[0]


def find_matching_transcription_files(directory, meeting_info):
    matched_files = []
    time_flexibility = 2  # 1 second flexibility

    for seconds_diff in range(-time_flexibility, time_flexibility + 1):
        adjusted_time = meeting_info["end_time"] + timedelta(seconds=seconds_diff)
        time_str = adjusted_time.strftime("%Y%m%d%H%M%S")

        summary_pattern = f"TencentMeeting_{time_str}_Summary*.txt"
        transcription_pattern = f"TencentMeeting_({time_str})_Transcription*.txt"

        matched_files += glob.glob(os.path.join(directory, summary_pattern))
        matched_files += glob.glob(os.path.join(directory, transcription_pattern))

    if len(matched_files) != 2:
        print(f"Found {len(matched_files)} transcription files")
        return None, None

    return matched_files[0], matched_files[1]


def process_all_meetings(directory_path, excel_pattern=EXCEL_PATTERN):
    """
    Process all meetings based on the Excel files found with the given pattern.
    """
    rename_peer = []

    for excel_file in glob.glob(os.path.join(directory_path, EXCEL_PATTERN)):
        meeting_info = read_meeting_info_from_excel(excel_file)
        new_base_name = f"【{meeting_info['scheduled_start_time'].strftime('%Y-%m-%d')}】{meeting_info['meeting_theme']}"

        video = find_matching_video_file(directory_path, meeting_info)
        summary, transcription = find_matching_transcription_files(
            directory_path, meeting_info
        )

        if None in [video, summary, transcription]:
            print(f"Skipping {excel_file}")
            continue

        new_excel_name = new_base_name + ".xlsx"
        new_video_name = new_base_name + ".mp4"
        new_summary_name = new_base_name + "_Summary.txt"
        new_transcription_name = new_base_name + "_Transcription.txt"

        rename_peer.append([excel_file, os.path.join(directory_path, new_excel_name)])
        rename_peer.append([video, os.path.join(directory_path, new_video_name)])
        rename_peer.append([summary, os.path.join(directory_path, new_summary_name)])
        rename_peer.append(
            [transcription, os.path.join(directory_path, new_transcription_name)]
        )

    for peer in rename_peer:
        os.rename(peer[0], peer[1])
        print(f"Renamed {peer[0]} ——> {peer[1]}")


if __name__ == "__main__":
    process_all_meetings("<path>")
