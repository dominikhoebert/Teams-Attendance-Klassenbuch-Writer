import argparse
import pandas as pd


def parse():
    parser = argparse.ArgumentParser(
        description="Converst a MS Teams Attendance List to a TGM Klassenbuch List"
    )
    parser.add_argument(
        "-in",
        help="CSV Inputfile donwloadable from a MS Teams Meeting\nStructure:\nFull Name (Lastname Firstname),\nUser Action (Joined/Left),\nTimestamp (dd/mm/yyyy, HH:MM:SS)\nseparated by tabs\nDefault: meetingAttendanceList.csv",
        dest="attendance_table",
        type=str,
        default="meetingAttendanceList.csv",
    )
    parser.add_argument(
        "-l",
        help="2 or 4 Number of Lessons, 4 for Lernbüro\nDefault: 2",
        dest="lessons",
        type=int,
        default=2,
    )
    parser.add_argument(
        "-s",
        help="Only for 4 lessons, Time when 4 lessons will be split into 2 and 2\nFormat: HH:MM\nDefault: 14:00",
        type=str,
        dest="separator",
        default="14:00",
    )
    parser.add_argument(
        "-st",
        help="CSV Table of Students\nStructure:\nnr (Katalagnummer)\nname (Nachname Vorname)\nklasse\nseparated by ';'\nDefault: students.csv",
        dest="students_table",
        type=str,
        default="students.csv",
    )
    parser.add_argument(
        "-z",
        help="Calculates how many Minutes somebody is late, or leaves early, with an offset. Values: True when -z set\nDefault: False",
        dest="calc_z",
        default=False,
        action="store_true",
    )
    parser.add_argument(
        "-o",
        dest="offset",
        type=int,
        default=10,
        help="Only with calc_z, time someone can be late or leave early, without getting Z.\nDefault: 10",
    )
    args = parser.parse_args()
    return args


# Funktion Berechnet die anwesende Zeit jeder Person
def calc_attendance_time(attendance, start_time, end_time):
    attended_students = pd.unique(attendance.name)
    attended_times = {}
    ts_string = {}
    late_time = {}
    left_early = {}
    joined = None
    for student in attended_students:
        attended_times[student] = 0
        ts_string[student] = ""
        for index, row in attendance[attendance.name == student].iterrows():
            if row["User Action"] == "Joined":
                joined = row.Timestamp
                if student not in late_time:
                    late_time[student] = (row.Timestamp - start_time).seconds // 60
            elif row["User Action"] == "Left":
                if joined:
                    attended_times[student] += (row.Timestamp - joined).seconds // 60
                    ts_string[
                        student
                    ] += f'{joined.strftime("%H:%M")}-{row.Timestamp.strftime("%H:%M")}; '
                    joined = None
                else:
                    attended_times[student] += (
                        row.Timestamp - start_time
                    ).seconds // 60
                    ts_string[
                        student
                    ] += f'S:{start_time.strftime("%H:%M")}-{row.Timestamp.strftime("%H:%M")}; '
                early_time = (end_time - row.Timestamp).seconds // 60
                if student not in left_early or early_time < left_early[student]:
                    left_early[student] = early_time
                # print(student, end_time, row.Timestamp, early_time)
        if joined:
            attended_times[student] += (end_time - joined).seconds // 60
            ts_string[
                student
            ] += f'{joined.strftime("%H:%M")}-E:{end_time.strftime("%H:%M")}; '
    l = pd.DataFrame.from_dict(late_time, orient="index", columns=["late"])
    le = pd.DataFrame.from_dict(
        left_early, orient="index", columns=["left_early"]
    ).merge(l, left_index=True, right_index=True, how="outer")
    ts = pd.DataFrame.from_dict(ts_string, orient="index", columns=["ts"]).merge(
        le, left_index=True, right_index=True
    )
    at = pd.DataFrame.from_dict(attended_times, orient="index", columns=["time"])
    return at.merge(ts, left_index=True, right_index=True)


if __name__ == "__main__":
    args = parse()
    # print(args)
    attendance_table = args.attendance_table
    students_table = args.students_table
    lessons = args.lessons
    separator = args.separator
    calc_z = args.calc_z
    offset = args.offset

    if lessons != 2 and lessons != 4:
        print("Lessons has to be 2 or 4! Exiting...")
        exit()

    if offset < 0:
        print("Offset must be greater than 0! Exiting...")
        exit()

    try:
        students = pd.read_csv(students_table, delimiter=";")
    except FileNotFoundError as fe:
        print(fe)
        print(f"Students Table {students_table} not found! Exiting...")
        exit()

    if not set(["nr", "name", "klasse"]).issubset(students.columns):
        print(
            f"Students Table {students_table} Structure not correct! Hint: nr;name;klasse \n Exiting..."
        )
        exit()

    students.name = students.name.str.replace("  ", " ")
    # "Replace double Space from Klassenbuch Liste"

    try:
        attendance = pd.read_csv(attendance_table, encoding="utf-16", delimiter="\t")
    except FileNotFoundError as fe:
        print(fe)
        print(f"Attendance Table {attendance_table} not found! Exiting...")
        exit()

    if not set(["Full Name", "User Action", "Timestamp"]).issubset(attendance.columns):
        print(
            f"Attendance Table {attendance_table} Structure not correct! Hint: Full Name	User Action	Timestamp \nSeparated by tabs! \n Exiting..."
        )
        exit()

    attendance.Timestamp = pd.to_datetime(attendance.Timestamp, dayfirst=True)

    date = attendance.Timestamp.dt.date[0]  # Datum zwischenspeichern
    separator = pd.Timestamp(str(date) + " " + separator)
    # 4 Stundenblock seperator Zeitpunkt zu Timestamp konvertieren

    print(
        f"Chosen Settings:\n\tAttendance List: {attendance_table}\n\tStudents Table: {students_table}\n\t# Lessons: {lessons}\n\tSepareting Time: {separator}\n\tCalculating Z#: {calc_z}\n\tLateness Offset: {offset}"
    )

    attendance = attendance.merge(
        right=students, left_on="Full Name", right_on="name", how="left"
    )
    # Anwesenheit mit Klassenbuchliste verbinden

    print(
        ", ".join(pd.unique(attendance[attendance.name.isna()]["Full Name"])),
        "konnten nicht in der Klassenliste gefunden werden! Eventuell ist die Schreibweise in der Klassenliste anders.",
    )
    # Ausgabe nicht gefundener Personen die anwesend waren ausgeben

    attendance = attendance[attendance.nr.notnull()]  # nicht gefundene Personen löschen

    start_time = attendance.Timestamp.min()
    end_time = attendance.Timestamp.max()
    if lessons == 2:
        attended_times = calc_attendance_time(attendance, start_time, end_time)
    elif lessons == 4:
        attended_times1 = calc_attendance_time(
            attendance[attendance.Timestamp < separator], start_time, separator
        )
        attended_times2 = calc_attendance_time(
            attendance[attendance.Timestamp > separator], separator, end_time
        )
        attended_times = attended_times1.merge(
            attended_times2,
            left_index=True,
            right_index=True,
            how="outer",
            suffixes=("1", "2"),
        )
        attended_times.time1 = attended_times.time1.fillna(0)
        attended_times.time2 = attended_times.time2.fillna(0)
        attended_times["time"] = attended_times.time1 + attended_times.time2

    new_students = students.merge(
        attended_times, left_on="name", right_index=True, how="left"
    )  # Verbinde list der Anwesenheitszeiten mit der Klassenliste

    # Erstellt eine Liste an Klassen von denen SchülerInnen anwesend waren
    classes = pd.unique(new_students.klasse)
    empty_classes = []
    for clas in classes:
        sum_time = new_students[new_students.klasse == clas].time.sum()
        if sum_time == 0:
            empty_classes.append(clas)
    new_students = new_students[~new_students.klasse.isin(empty_classes)]

    hours = [i for i in range(1, lessons + 1)]
    col = [i for i in reversed(hours)]
    for i in hours:
        new_students[i] = ""

    # Berechnet die Zuspät gekommene Zeit
    if calc_z:
        if lessons == 2:
            new_students.loc[new_students.time.notnull(), hours] = "A"
            new_students.late = new_students.late.fillna(0)
            new_students.loc[
                new_students.late > offset, 1
            ] = "Z" + new_students.late.astype(int).astype(str)
            new_students.left_early = new_students.left_early.fillna(0)
            new_students.loc[
                new_students.left_early > offset, 2
            ] = "Z" + new_students.left_early.astype(int).astype(str)
            col += ["nr", "name", "klasse", "time", "ts", "left_early", "late"]
        elif lessons == 4:
            col += [
                "nr",
                "name",
                "klasse",
                "time",
                "time1",
                "ts1",
                "left_early1",
                "late1",
                "time2",
                "ts2",
                "left_early2",
                "late2",
            ]
            new_students.loc[new_students.ts1.notnull(), [1, 2]] = "A"
            new_students.loc[new_students.ts2.notnull(), [3, 4]] = "A"
            new_students.late1 = new_students.late1.fillna(0)
            new_students.late2 = new_students.late2.fillna(0)
            new_students.left_early1 = new_students.left_early1.fillna(0)
            new_students.left_early2 = new_students.left_early2.fillna(0)
            new_students.loc[
                new_students.late1 > offset, 1
            ] = "Z" + new_students.late1.astype(int).astype(str)
            new_students.loc[
                new_students.late2 > offset, 3
            ] = "Z" + new_students.late2.astype(int).astype(str)
            new_students.loc[
                new_students.left_early1 > offset, 2
            ] = "Z" + new_students.left_early1.astype(int).astype(str)
            new_students.loc[
                new_students.left_early2 > offset, 4
            ] = "Z" + new_students.left_early2.astype(int).astype(str)
    # Erzeugt A für Anwesend Tabelle
    else:
        if lessons == 2:
            new_students.loc[new_students.time.notnull(), hours] = "A"
            col += ["nr", "name", "klasse", "time", "ts", "left_early", "late"]
        elif lessons == 4:
            col += [
                "nr",
                "name",
                "klasse",
                "time",
                "time1",
                "ts1",
                "left_early1",
                "late1",
                "time2",
                "ts2",
                "left_early2",
                "late2",
            ]
            new_students.loc[new_students.ts1.notnull(), [1, 2]] = "A"
            new_students.loc[new_students.ts2.notnull(), [3, 4]] = "A"

    # Name des Ausgabefiles erstellen
    classes = [c for c in classes if c not in empty_classes]
    classes_string = " ".join(classes)
    output_excel_name = f'{date.strftime("%Y%m%d")} {classes_string}.xlsx'

    print(f"\n{output_excel_name} for the classes {', '.join(classes)} created!\n")

    # Excel File erstellen
    new_students[new_students.klasse == classes[0]][col].to_excel(
        output_excel_name, sheet_name=classes[0]
    )
    with pd.ExcelWriter(output_excel_name, mode="a") as writer:
        for c in classes[1:]:
            new_students[new_students.klasse == c][col].to_excel(writer, sheet_name=c)
