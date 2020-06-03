# -- coding: utf-8 --

import click
import os
from win32com.client import Dispatch
import contextlib
from collections import OrderedDict
from jinja2 import Template
from redminelib import Redmine
from datetime import date


class Task(object):
    def __init__(self, name: str, uid: int, progress: int, status: str, start_time: str, end_time: str, **kwargs):
        self.name = name
        self.uid = uid
        self.child_tasks = []
        self.progress = progress  # 0-100
        self.status = status
        date_template = '{year}年{month}月{day}日'
        if isinstance(start_time, date):
            self.start_time = date_template.format(year=start_time.year, month=start_time.month, day=start_time.day)
        else:
            self.start_time = start_time
        if isinstance(end_time, date):
            self.end_time = date_template.format(year=end_time.year, month=end_time.month, day=end_time.day)
        else:
            self.end_time = end_time

    def have_child_tasks(self):
        return bool(len(self.child_tasks))

    def append_task(self, task):
        self.child_tasks.append(task)
        return self


class Project(object):
    def __init__(self, name: str, uid: int, tasks, **kwargs):
        self.name = name
        self.uid = uid
        self.tasks = tasks

    def append_task(self, task):
        self.tasks.append(task)
        return self

    def extend_task(self, task):
        self.tasks.extend(task)
        return self


class PowerPoint(object):
    def __init__(self, source_name: str, target_name: str):
        self.ppt_app = Dispatch('PowerPoint.Application')
        self.ppt_app.Visible = 1
        source_file_path = os.path.join(*(os.getcwd(), source_name))
        target_file_path = os.path.join(*(os.getcwd(), target_name))
        self.ppt = self.ppt_app.Presentations.Open(source_file_path)
        self.target_file_path = target_file_path
        self.target_name = target_name
        self.row_num = 2
        self._cached_cells = {}

    def get_table(self, slide_index, shape_index):
        return self.ppt.slides[slide_index].shapes[shape_index].Table

    def get_rows(self, slide_index, shape_index):
        return self.get_table(slide_index, shape_index).Rows

    def _calculate_hash_for_cells(self, slide_index: int, shape_index: int, row_index: int):
        assert row_index <= self.row_num
        hash_str = str(hash(slide_index)) + str(hash(shape_index)) + str(hash(row_index))
        return hash_str

    def _get_cached_cells(self, slide_index, shape_index, row_index):
        hash_str = self._calculate_hash_for_cells(slide_index, shape_index, row_index)
        return self._cached_cells.get(hash_str)

    def _set_cached_cells(self, slide_index, shape_index, row_index, cells):
        hash_str = self._calculate_hash_for_cells(slide_index, shape_index, row_index)
        return self._cached_cells.update({hash_str: cells})

    def get_cells(self, slide_index=2, shape_index=1, row_index=1):
        cells = self._get_cached_cells(slide_index, shape_index, row_index)
        if cells is None:
            cells = self.get_rows(slide_index, shape_index)[row_index].Cells
            self._set_cached_cells(slide_index, shape_index, row_index, cells)
        return cells

    def set_text(self, text, slide_index=2, shape_index=1, row_index=1, column_index=0):
        assert row_index <= self.row_num + 1
        self.get_cells(slide_index, shape_index, row_index)[column_index].Shape.TextFrame.TextRange.Text = text

    def get_text(self, slide_index=2, shape_index=1, row_index=1, column_index=0):
        assert row_index <= self.row_num + 1
        return self.get_cells(slide_index, shape_index, row_index)[column_index].Shape.TextFrame.TextRange.Text

    def get_cell(self, row_index=1, column_index=0, **kwargs):
        return self.get_cells(row_index=row_index, **kwargs)[column_index]

    @staticmethod
    def merge(src_cell, dst_cell):
        src_cell.Merge(dst_cell)

    @staticmethod
    def get_attr(obj, obj_type, attr):
        if obj_type == 'cell':
            if attr == 'text':
                return obj.Shape.TextFrame.TextRange.Text

    def _set_table_row_num(self, row_num):
        assert row_num >= 2
        self.ppt.slides[2].shapes[1].Table.Rows.Add(2)
        self.ppt.slides[2].shapes[1].Table.Rows[2].Delete()
        self.ppt.slides[2].shapes[1].Table.Rows.Add(2)
        for num in range(0, row_num - 2):
            self.ppt.slides[2].shapes[1].Table.Rows.Add(2)

    @contextlib.contextmanager
    def context(self, row_num: int = 2):
        if row_num < 2 and row_num != 0:
            row_num = 2
        self.row_num = row_num
        if row_num != 0:
            self._set_table_row_num(row_num)
        yield self
        self.ppt.SaveAs(self.target_file_path)
        self.ppt_app.Quit()

    def get_row_num(self):
        return self.row_num

    def get_text_frame(self, slide_index, shape_index):
        shape = self.ppt.slides[slide_index].shapes[shape_index]
        if shape.HasTextFrame:
            return shape.TextFrame
        return None

    def set_user_fullname(self, text, slide_index=2, shape_index=0):
        text_frame = self.get_text_frame(slide_index, shape_index)
        if text_frame is not None:
            old_text = text_frame.TextRange.Text
            old_text = old_text.replace('X', '').replace('x', '')
            new_text = '{0}{1}'.format(text, old_text)
            text_frame.TextRange.Text = new_text


class WorkTable(object):
    def __init__(self, ppt: PowerPoint, projects: [], slide_index=2, shape_index=1,
                 start_row: int = 1, start_column: int = 0, max_column: int = 10, user_fullname=''):
        self.ppt = ppt
        self.slide_index = slide_index
        self.shape_index = shape_index
        self.start_row = start_row
        self.start_column = start_column
        self.column_count = 0
        self.row_count = 0
        self.max_column = max_column
        self._columns = self.parse()
        self._rows = self.pre_process(projects)
        self._columns_num = len(self._columns)
        self._rows_num = len(self._rows)
        self._cached_data = OrderedDict()
        self.user_fullname = user_fullname

    def parse(self):
        columns = []
        for column_index, first_row_cell in enumerate(self.ppt.get_cells(self.slide_index,
                                                                         self.shape_index, row_index=0)):
            second_row_cell = self.ppt.get_cell(column_index=column_index)
            columns.append(ColumnRawData(column_index,
                                         self.ppt.get_attr(first_row_cell, 'cell', 'text'),
                                         self.ppt.get_attr(second_row_cell, 'cell', 'text'),
                                         self.ppt))
        return columns

    def render(self, *args, **context):
        self.column_count = self.start_column
        for column in self._columns:
            can_merge = False
            if column.can_render():
                if column.can_merge():
                    can_merge = True
                data = column.render(*args, **context)
            else:
                data = column.get_text()
            self.write(can_merge, data)
            self.column_count += 1
        self.row_count += 1

    @property
    def current_column(self):
        return self.start_column + self.column_count

    @property
    def current_row(self):
        return self.start_row + self.row_count

    def _has_equal_cached_data(self, data):
        if not self._cached_data:
            return False

        for key, value in self._cached_data.items():
            if value[1] != data:
                return False

        return True

    def merge_cells(self):
        columns = list(self._cached_data.items())
        if self._cached_data and len(columns) >= 2:
            current_column = columns[0][1][0]
            current_row = columns[0][0]
            src_cell = self.ppt.get_cell(current_row, current_column)
            dst_cell = self.ppt.get_cell(columns[-1][0], current_column)
            self.ppt.set_text(columns[0][1][1], row_index=current_row, column_index=current_column)
            self.ppt.merge(src_cell, dst_cell)
            self._cached_data = OrderedDict()
        elif self._cached_data and len(columns) == 1:
            current_column = columns[0][1][0]
            current_row = columns[0][0]
            self.ppt.set_text(columns[0][1][1], row_index=current_row, column_index=current_column)
            self._cached_data = OrderedDict()

    def write(self, can_merge: bool, data: str, **context):
        if can_merge:
            if not self._has_equal_cached_data(data):
                self.merge_cells()
            self._cached_data.update({self.current_row: (self.current_column, data)})
        else:
            self.ppt.set_text(data, row_index=self.current_row, column_index=self.current_column)

    @staticmethod
    def pre_process(projects: [Project]):
        rows = []
        for project in projects:
            for task in project.tasks:
                parent_task = None
                if task.have_child_tasks():
                    parent_task = task
                    for child_task in task.child_tasks:
                        current_task = child_task
                        rows.append((project, parent_task, current_task))
                else:
                    current_task = task
                    rows.append((project, parent_task, current_task))
        return rows

    def process(self):
        error_flag = False
        with self.ppt.context(self._rows_num):
            try:
                for project, parent_task, current_task in self._rows:
                    self.render(project=project, parent_task=parent_task, current_task=current_task)
                self.merge_cells()
                self.ppt.set_user_fullname(self.user_fullname, self.slide_index)
            except Exception as e:
                error_flag = True
                print(e)
        if error_flag:
            raise Exception


class ColumnRawData(object):
    def __init__(self, column_id, field_text, render_text, ppt):
        self.column_id = column_id
        self.field_text = field_text
        self.render_text = render_text
        self.template = Template(render_text)
        self.ppt = ppt

    def can_render(self):
        return '{{' in self.render_text

    def can_merge(self):
        return 'merge' in self.render_text

    def render(self, *args, **context):
        return self.template.render(*args, **context)

    def get_text(self):
        return self.render_text


class RedmineInterface(object):
    def __init__(self, url, key='', start_time='2019-03-01', end_time='2020-05-07', username='', password=''):
        self.url = url or 'http://192.168.67.129:7777/redmine'
        if key == '':
            key = None
        self.key = key
        self.redmine = Redmine(url, key=key, username=username, password=password)
        self.current = self.redmine.user.get('current')
        self.start_time = start_time
        self.end_time = end_time
        self.issues = self.redmine.issue.filter(assigned_to_id=str(self.current.id),
                                                limit=100,
                                                status_id="*",
                                                sort='updated_on:desc',
                                                start_date='><{start_time}|{end_time}'.format(
                                                    start_time=self.start_time,
                                                    end_time=self.end_time
                                                ))

    def _issue_iter(self, issues):
        for issue in issues:
            for _issue in self.issues:
                if issue.id == _issue.id:
                    yield _issue

    def get_projects(self):
        projects = []
        for issue in self.issues:
            project_id = issue.project.id
            if getattr(issue, 'parent', None) is None:
                tasks = [Task(issue.subject, issue.id, issue.done_ratio, issue.status.name,
                              getattr(issue, 'start_date', ''), getattr(issue, 'due_date', ''))]
                if len(issue.children):
                    tasks = [tasks[0].append_task(
                        Task(_issue.subject, _issue.id, _issue.done_ratio, _issue.status.name,
                             getattr(_issue, 'start_date', ''), getattr(_issue, 'due_date', ''))
                    ) for _issue in self._issue_iter(issue.children)]
            else:
                tasks = []
            project = None
            for _project in projects:
                if _project.uid == project_id:
                    project = _project
            if project is None:
                project = Project(issue.project.name, project_id, [])
                projects.append(project)
            project.extend_task(tasks)
        return projects

    def get_current_user_fullname(self):
        return '{0}{1}'.format(self.current.lastname, self.current.firstname)


def generate_projects(*args, **kwargs):
    redmine = RedmineInterface(*args, **kwargs)
    projects = redmine.get_projects()
    return projects


def get_current_user_fullname(*args, **kwargs):
    redmine = RedmineInterface(*args, **kwargs)
    return redmine.get_current_user_fullname()


def process(*args, **kwargs):
    power_point = PowerPoint("template.pptx", "release.pptx")
    work_table = WorkTable(power_point, generate_projects(*args, **kwargs),
                           user_fullname=get_current_user_fullname(*args, **kwargs))
    work_table.process()


@click.command()
@click.option("--url", default='http://spdm/redmine/',
              help="server address example: http://192.168.67.129:7777/redmine")
@click.option("--key", default='', help="SPDM access token")
@click.option("--username", default='', help="SPDM username")
@click.option("--password", default='', help="SPDM password")
@click.option("--start_time", help="task start time example:2019-03-01", required=True)
@click.option("--end_time", help="task end time example:2020-05-07", required=True)
def gen_ppt(url, key, start_time, end_time, username, password):
    """Generate Powerpoint"""
    try:
        process(url=url, key=key, start_time=start_time, end_time=end_time, username=username, password=password)
    except Exception as e:
        click.echo(str(e))


if __name__ == '__main__':
    gen_ppt()
