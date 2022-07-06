import aiohttp
import asyncio
import pandas as pd
import xlsxwriter
from bs4 import BeautifulSoup


class Course:
    def __init__(self, code):
        self.code = code
        self.name = None
        self.credit = None
        course_code_split = code.split()
        self.url = f'https://ce.ieu.edu.tr/en/syllabus/type/read/id/{course_code_split[0]}+{course_code_split[1]}'
        self.department = course_code_split[0]
        self.evaluations = pd.DataFrame(columns=['metric', 'number', 'weight', 'grade'])


class CourseTable:
    def __init__(self, course_codes: list):
        self.courses = [Course(course_code) for course_code in course_codes]
        asyncio.run(self.scrape_courses())

    async def scrape_courses(self):
        sem = asyncio.Semaphore(50)
        async with aiohttp.ClientSession() as session:
            async def get_course_info(url):
                async with sem, session.get(url) as resp:
                    data = await resp.text()
                    soup = BeautifulSoup(data, "html.parser")
                    return [soup.find(id='course_name').string.strip(),
                            soup.find(id='evaluation_table1').findAll('tr')[1:],
                            soup.find(id='ects_credit').string.strip()]

            tasks = [asyncio.create_task(get_course_info(course.url)) for course in self.courses]
            results = await asyncio.gather(*tasks)

        for course, result in zip(self.courses, results):
            course.name = result[0]
            course.credit = int(result[2])
            for evaluation in result[1]:
                tds = evaluation.findAll('td')
                eval_weight = tds[2].find('div').string
                if eval_weight is not None and eval_weight != '-':
                    eval_name = tds[0].string
                    eval_number = tds[1].find('div').string
                    new_row = pd.DataFrame([[eval_name, int(eval_number), int(eval_weight) / 100, None]],
                                           columns=['metric', 'number', 'weight', 'grade'])
                    course.evaluations = pd.concat([course.evaluations, new_row], ignore_index=True)

    def to_excel(self, filename):
        excel_writer = pd.ExcelWriter(filename, engine='xlsxwriter')

        df_out = pd.DataFrame(columns=['metric', 'number', 'weight'])
        df_calc = pd.DataFrame([['AA', 4, 90, 100],
                                ['BA', 3.5, 85, 89.5],
                                ['BB', 3, 80, 84.5],
                                ['CB', 2.5, 75, 79.5],
                                ['CC', 2, 70, 74.5],
                                ['DC', 1.5, 65, 69.5],
                                ['DD', 1, 60, 64.5],
                                ['FD', 0.5, 50, 59.5],
                                ['FF', 0, 0, 49.5]])
        for course in self.courses:
            df_out = pd.concat([df_out, pd.DataFrame([[course.code, 'Number', 'Weight', 'Grade']],
                                                     columns=['metric', 'number', 'weight', 'grade'])],
                               ignore_index=True)
            df_out = pd.concat([df_out, course.evaluations],
                               ignore_index=True)

            # noinspection PyTypeChecker
            last_index = df_out.last_valid_index() + 1
            metric_num = len(course.evaluations.index) - 1
            formula = f'=SUMPRODUCT(Courses!C{last_index - metric_num}:C{last_index},' \
                      f'Courses!D{last_index - metric_num}:D{last_index})'
            df_out = pd.concat([df_out, pd.DataFrame([[None, None, None, formula]],
                                                     columns=['metric', 'number', 'weight', 'grade'])],
                               ignore_index=True)
        df_out.to_excel(excel_writer, sheet_name='Courses', header=False, index=False)
        df_calc.to_excel(excel_writer, sheet_name='letter_grades', header=False, index=False)

        workbook = excel_writer.book
        courses_sheet = excel_writer.sheets['Courses']
        grade_sheet = excel_writer.sheets['letter_grades']

        unlocked = workbook.add_format({'locked': False})
        unlocked.set_align('center')
        bold = workbook.add_format({'bold': True})
        bold.set_align('center')
        decimal_places = workbook.add_format({'num_format': '#,##0.00'})
        decimal_places.set_align('center')
        center_text = workbook.add_format()
        center_text.set_align('center')

        for col_i in range(len(df_out.columns) - 1):  # loop through all columns
            series = df_out.iloc[:, col_i]
            max_len = max((series.astype(str).map(len).max(),
                           len(str(series.name)))) + 1
            courses_sheet.set_column(col_i, col_i, max_len)  # set column width

        courses_sheet.protect()
        grade_sheet.protect()
        grade_sheet.hide()

        course_i = 0
        for row_i in range(len(df_out.index)):
            grade_val = df_out.iloc[row_i, 3]
            if grade_val == 'Grade':
                courses_sheet.set_row(row_i, None, bold)
            elif grade_val is None:
                courses_sheet.write(row_i, 3, None, unlocked)
            else:  # grade_val == '=SUMP...'
                grade_sheet.write(course_i, 4, f'=Courses!D{row_i + 1}')
                grade_sheet.write(course_i, 5,
                                  f'=INDEX(letter_grades!A1:D9,MATCH(E{course_i + 1},letter_grades!D1:D9,-1),2)')
                grade_sheet.write(course_i, 6, self.courses[course_i].credit)
                grade_sheet.write(course_i, 7, f'=F{course_i+1}*G{course_i+1}')
                course_i += 1

        grade_sheet.write(course_i, 6, f'=SUM(G1:G{course_i})')
        grade_sheet.write(course_i, 7, f'=SUM(H1:H{course_i})/G{course_i+1}')
        courses_sheet.write(1, 5, 'GPA', bold)
        courses_sheet.write(2, 5, f'=letter_grades!H{course_i+1}', decimal_places)

        courses_sheet.set_column('B:F', None, center_text)

        excel_writer.save()


if __name__ == '__main__':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    course_names = ['CE 302', 'MATH 250', 'CE 475', 'CE 306', 'ENG 310']  # Course codes that will be parsed

    ct = CourseTable(course_names)
    ct.to_excel('grades.xlsx')  # Output path of the Excel file
