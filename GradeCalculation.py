import aiohttp
import asyncio
import pandas as pd
import xlsxwriter
from bs4 import BeautifulSoup


class Course:
    def __init__(self, code):
        self.code = code
        self.name = None
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
                            soup.find(id='evaluation_table1').findAll('tr')[1:]]

            tasks = [asyncio.create_task(get_course_info(course.url)) for course in self.courses]
            results = await asyncio.gather(*tasks)

        for course, result in zip(self.courses, results):
            course.name = result[0]
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

        for course in self.courses:
            df_out = pd.concat([df_out, pd.DataFrame([[course.code, 'Number', 'Weight', 'Grade']],
                                                     columns=['metric', 'number', 'weight', 'grade'])],
                               ignore_index=True)
            df_out = pd.concat([df_out, course.evaluations],
                               ignore_index=True)

            # noinspection PyTypeChecker
            last_index = df_out.last_valid_index() + 1
            metric_num = len(course.evaluations.index) - 1
            formula = f'=SUMPRODUCT(C{last_index - metric_num}:C{last_index},D{last_index - metric_num}:D{last_index})'
            df_out = pd.concat([df_out, pd.DataFrame([[None, None, None, formula]],
                                                     columns=['metric', 'number', 'weight', 'grade'])],
                               ignore_index=True)
        df_out.to_excel(excel_writer, sheet_name='Courses', header=False, index=False)
        workbook = excel_writer.book
        worksheet = excel_writer.sheets['Courses']

        unlocked = workbook.add_format({'locked': False})
        bold = workbook.add_format({'bold': True})

        for col_i in range(len(df_out.columns) - 1):  # loop through all columns
            series = df_out.iloc[:, col_i]
            max_len = max((series.astype(str).map(len).max(),
                           len(str(series.name)))) + 1
            worksheet.set_column(col_i, col_i, max_len)  # set column width

        worksheet.protect()

        for row_i in range(len(df_out.index)):
            grade_val = df_out.iloc[row_i, 3]
            if grade_val == 'Grade':
                # row = f'A{row_i}:D{row_i}'
                # worksheet.write(row, bold)
                worksheet.set_row(row_i, None, bold)
            if grade_val is None:
                worksheet.write(row_i, 3, None, unlocked)

        excel_writer.save()

    def to_xlsxwriter(self, filename):

        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet('Courses')

        locked = workbook.add_format({'locked': True})
        unlocked = workbook.add_format({'locked': False})

        excel_writer = pd.ExcelWriter(filename, engine='xlsxwriter')

        df_out = pd.DataFrame(columns=['metric', 'number', 'weight'])

        for course in self.courses:
            df_out = pd.concat([df_out, pd.DataFrame([[course.code, 'Number', 'Weight', 'Grade']],
                                                     columns=['metric', 'number', 'weight', 'grade'])],
                               ignore_index=True)
            df_out = pd.concat([df_out, course.evaluations],
                               ignore_index=True)

            # noinspection PyTypeChecker
            last_index = df_out.last_valid_index() + 1
            metric_num = len(course.evaluations.index) - 1
            formula = f'=SUMPRODUCT(C{last_index - metric_num}:C{last_index},D{last_index - metric_num}:D{last_index})'
            df_out = pd.concat([df_out, pd.DataFrame([[None, None, None, formula]],
                                                     columns=['metric', 'number', 'weight', 'grade'])],
                               ignore_index=True)
        df_out.to_excel(excel_writer, sheet_name='Courses', header=False, index=False)

        for col_i in range(len(df_out.columns) - 1):  # loop through all columns
            series = df_out.iloc[:, col_i]
            max_len = max((series.astype(str).map(len).max(),
                           len(str(series.name)))) + 1
            excel_writer.sheets['Courses'].set_column(col_i, col_i, max_len)  # set column width
        excel_writer.save()


if __name__ == '__main__':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    course_names = ['CE 302', 'MATH 250', 'CE 475', 'CE 306', 'ENG 310']  # Course codes that will be parsed

    ct = CourseTable(course_names)
    ct.to_excel('grades.xlsx')  # Output path of the Excel file
