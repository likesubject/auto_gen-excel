# -- coding: utf-8 --

from click.testing import CliRunner

from main import (ExcelAdapter, ColumnRawData, WorkTable, RedmineAdapter, gen_excel)


TEST_REDMINE_URL = 'http://192.168.67.133:7777/redmine'


class TestPowerpoint(object):
    def test_insert_table_and_text(self):
        excel_proxy = ExcelAdapter("template.xlsx", "release1.xlsx")
        error_flag = False
        with excel_proxy.context() as excel:
            try:
                for row_index in range(2, 10):
                    for column_index in range(1, 10):
                        excel.set_text('row:{0},column:{1}'.format(row_index, column_index),
                                       row_index=row_index, column_index=column_index)
                src_cell = excel.get_cell(1, 1)
                dst_cell = excel.get_cell(2, 1)
                excel.merge(src_cell, dst_cell)
            except Exception as e:
                error_flag = True
                print(e)

        assert not error_flag

    def test_get_template_content(self):
        excel_proxy = ExcelAdapter("template.xlsx", "release2.xlsx")
        columns = []
        with excel_proxy.context() as excel:
            for column_index, first_row_cell in enumerate(excel.get_cells(row_index=1)):
                second_row_cell = excel.get_cell(column_index=column_index + 1, row_index=2)
                columns.append(ColumnRawData(column_index + 1,
                                             first_row_cell.get_text(),
                                             second_row_cell.get_text(),
                                             excel))
        print(columns)
        assert len(columns) == 7


class TestRedmineAdapter(object):
    def test_generate_projects_with_month(self):
        redmine = RedmineAdapter(TEST_REDMINE_URL,
                                 key='5f4821802e9cd29fb2ac54a13fc98d15e760b865',
                                 month=6)
        projects = redmine.get_projects()
        _projects = []
        _projects.extend(projects.projects)
        assert len(_projects) >= 1

    def test_generate_projects_with_month_and_year(self):
        redmine = RedmineAdapter(TEST_REDMINE_URL,
                                 key='5f4821802e9cd29fb2ac54a13fc98d15e760b865',
                                 month=6,
                                 year=2019)
        projects = redmine.get_projects()
        _projects = []
        _projects.extend(projects.projects)
        assert len(_projects) == 0

    @staticmethod
    def generate_test_projects():
        redmine = RedmineAdapter(TEST_REDMINE_URL,
                                 key='5f4821802e9cd29fb2ac54a13fc98d15e760b865')
        projects = redmine.get_projects()
        return projects

    def test_get_projects(self):
        projects = self.generate_test_projects().projects
        _projects = []
        _projects.extend(projects)
        assert len(_projects) >= 1

    def test_get_spent_times(self):
        projects = self.generate_test_projects()
        for project in projects.projects:
            for user in project.users:
                spent_time = user.spent_time
                print(project, user, user.fullname, spent_time)

    def test_process(self):
        power_point = ExcelAdapter("template.xlsx", "release3.xlsx")
        work_table = WorkTable(power_point, self.generate_test_projects())
        work_table.process()


class TestCmd(object):
    def test_gen_ppt(self):
        runner = CliRunner()
        result = runner.invoke(gen_excel,
                               ['--key',
                                '5f4821802e9cd29fb2ac54a13fc98d15e760b865',
                                '--url',
                                TEST_REDMINE_URL,
                                '--month',
                                '6'])
        assert result.exit_code == 0

    def test_gen_ppt_with_year(self):
        runner = CliRunner()
        result = runner.invoke(gen_excel,
                               ['--key',
                                '5f4821802e9cd29fb2ac54a13fc98d15e760b865',
                                '--url',
                                TEST_REDMINE_URL,
                                '--month',
                                '6',
                                '--year',
                                '2016'])
        assert result.exit_code == 0

    def test_gen_ppt_with_project1(self):
        runner = CliRunner()
        result = runner.invoke(gen_excel,
                               ['--username',
                                'like',
                                '--password',
                                '6976630670',
                                '--url',
                                TEST_REDMINE_URL,
                                '--month',
                                '6',
                                '--project',
                                'spd'])
        assert result.exit_code == 0

    def test_gen_ppt_with_project2(self):
        runner = CliRunner()
        result = runner.invoke(gen_excel,
                               ['--username',
                                'like',
                                '--password',
                                '6976630670',
                                '--url',
                                TEST_REDMINE_URL,
                                '--month',
                                '6',
                                '--project',
                                'test_project1'])
        assert result.exit_code == 0
