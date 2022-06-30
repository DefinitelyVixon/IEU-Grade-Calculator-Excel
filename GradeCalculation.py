import aiohttp
import asyncio
import pandas as pd
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
        df_out.to_excel(filename, sheet_name='Courses', engine='xlsxwriter', header=False, index=False)


if __name__ == '__main__':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    course_names = ['CE 302', 'MATH 250', 'CE 475', 'CE 306', 'ENG 310']  # Course codes that will be parsed

    ct = CourseTable(course_names)
    ct.to_excel('grades.xlsx')  # Output path of the Excel file
