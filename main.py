import openpyxl as xl
from tkinter import filedialog
import re


class TimeTable():
    def __init__(self, sheet:xl.worksheet.worksheet.Worksheet, tt_range: str, fix_rows=True):
        self.rows = sheet[tt_range]
        self.fixrows(fix_rows)

    def fixrows(self, fix_rows):
        """Structure rows in the proper format for further processing)"""
        if fix_rows:
            rr = []
            for row in self.rows[1:]:
                rr.append(row[1:])

            self.rows = rr
        
        # Little more structuring.
        # Multiple possible classes in the same slot.
        self.rows = [[cell.value for cell in row] for row in self.rows]
        self.rows = [[cell.rstrip('\n').split('\n') if cell else [] for cell in row] for row in self.rows]

    
    def removeOtherCourses(self):
        """Removes other courses you are not registerd to
        Ensure you have correctly added your courses to your_courses.txt
        """

        input("Ensure you have added your course codes to the file 'your_courses.txt' before continuing. Press Enter to continue.")
        
        with open('your_courses.txt', 'r') as file:
            tlist = file.readlines()
        
        clist = []
        for c in tlist:
            if c[0] != '#' and c != '\n':
                clist.append(c.rstrip('\n'))

        self.clist = clist
        # print(self.clist)


        def isValidCourse(ele):
            """Inner Function verifies if course valid"""
            return any(map(lambda c: c in ele, self.clist))


        self.rows_updated = []
        for row in self.rows:
            nrow = []
            for cell in row:
                ncell = list(filter(isValidCourse, cell))
                nrow.append(ncell)
            self.rows_updated.append(nrow)

    def convertToCalendarEvents(self):
        timings = [
            '8:30-9:25',
            '9:30-10:25',
            '10:40-11:35',
            '11:40-12:35',
            '12:40-1:30',
            '13:30-14:25',
            '14:30-15:25',
            '15:40-16:35',
            '16:40-17:35',
            '17:40-18:35'
        ]
        timings = [timing.split('-') for timing in timings]


        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']     # Not used.
        dates = ['15/01/2024', '16/01/2024', '17/01/2024', '18/01/2024', '19/01/2024', '20/01/2024']    # Calendar starts from Monday 15th Jan 2024

        def getCourse(row, col):
            class_ele = "\n".join(self.rows_updated[i][j])      # Incase there are multiple courses (cuz Ben10 👍)
            if not class_ele:
                return
            
            room = re.search('{(.*?)}', class_ele)[0]
            
            class_ele = class_ele.replace(room, "")
            class_ele = class_ele.strip(': ')

            startt, endt = timings[row]
            date = dates[col]

            return date, startt, endt, class_ele, room


        self.calendar_events = []
        si, sj = len(self.rows_updated), len(self.rows_updated[0])
        for i in range(si):
            for j in range(sj):
                course = getCourse(i, j)
                if course:
                    self.calendar_events.append(course)

    def compile(self):
        self.removeOtherCourses()
        self.convertToCalendarEvents()

                 



if __name__ == '__main__':
    xl_file = filedialog.askopenfilename()
    print("Selected File:", xl_file)

    xl_file = xl.load_workbook(xl_file)

    # Time Table Range and verification
    tt_range = 'A4:G14'
    if input(f"TimeTable Range: {tt_range} Correct?"):
        tt_range = input("Input new Range: ")
        print("Range Update to:", tt_range)


    TT = TimeTable(xl_file.active, tt_range, fix_rows=True)
    TT.compile()
