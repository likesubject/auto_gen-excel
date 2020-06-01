# -- coding: utf-8 --

from click.testing import CliRunner

from main import (PowerPoint, WorkTable, Task, Project, RedmineInterface, gen_ppt)


class TestPowerpoint(object):
    def test_insert_table_and_text(self):
        power_point = PowerPoint("test.pptx", "new1.pptx")
        error_flag = False
        with power_point.context(7) as ppt:
            try:
                column_num = len(ppt.get_cells())
                row_num = ppt.get_row_num()
                print((column_num, row_num))
                for row_index in range(1, row_num ):
                    for column_index in range(1, column_num):
                        ppt.set_text('row:{0},column:{1}'.format(row_index, column_index),
                                     row_index=row_index, column_index=column_index)
                src_cell = ppt.get_cell(1, 0)
                dst_cell = ppt.get_cell(2, 0)
                ppt.merge(src_cell, dst_cell)
            except Exception as e:
                error_flag = True
                print(e)

        assert not error_flag


class TestWorkTable(object):
    @staticmethod
    def generate_test_projects():
        tasks = [Task('test_task' + str(i), i, i * 10, 'prepare', '2020/05/08', '2020/05/10') for i in range(0, 3)]
        projects = [Project('test_project' + str(i), i, tasks) for i in range(0, 2)]
        tasks = [Task('test_task' + str(i), i, i * 10, 'prepare', '2020/05/08', '2020/05/10') for i in range(0, 2)]
        projects1 = [Project('test_project' + str(i + 3), i, tasks) for i in range(0, 2)]
        projects.extend(projects1)
        tasks = [Task('test_task' + str(i), i, i * 10, 'prepare', '2020/05/08', '2020/05/10') for i in range(0, 1)]
        projects2 = [Project('test_project' + str(i + 5), i, tasks) for i in range(0, 2)]
        projects.extend(projects2)
        tasks = [Task('test_task' + str(i), i, i * 10, 'prepare', '2020/05/08', '2020/05/10') for i in range(0, 3)]
        projects3 = [Project('test_project' + str(i + 8), i, tasks) for i in range(0, 2)]
        projects.extend(projects3)
        return projects

    def test_process(self):
        power_point = PowerPoint("test.pptx", "new2.pptx")
        work_table = WorkTable(power_point, self.generate_test_projects())
        work_table.process()


class TestRedmine(object):
    @staticmethod
    def generate_test_projects():
        redmine = RedmineInterface('http://192.168.67.129:7777/redmine',
                                   key='5f4821802e9cd29fb2ac54a13fc98d15e760b865')
        projects = redmine.get_projects()
        return projects

    def test_process(self):
        power_point = PowerPoint("test.pptx", "new3.pptx")
        work_table = WorkTable(power_point, self.generate_test_projects())
        work_table.process()


class TestCmd(object):
    def test_gen_ppt(self):
        runner = CliRunner()
        result = runner.invoke(gen_ppt,
                               ['--key',
                                '5f4821802e9cd29fb2ac54a13fc98d15e760b865',
                                '--url',
                                'http://192.168.67.129:7777/redmine'])
        assert result.exit_code == 0
