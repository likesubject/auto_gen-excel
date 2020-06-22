SPDM 数据导出工具
^^^^^^^^^^^^^^^^^^^^^^^^

如何使用
---------
第一步：打开CMD 运行main.exe --help 可以查看所有的选项（如直接运行run.bat可跳过此步骤）
Usage: main.exe [OPTIONS]

  Generate Excel

Options:
  --url TEXT             server address example: http://spdm/redmine/
  --key TEXT             SPDM access token
  --username TEXT        SPDM username
  --password TEXT        SPDM password
  --year INTEGER         Statistical year,the default this year
  --month INTEGER        Statistical month  [required]
  --enable-merge-cells   enable merge cells
  --project TEXT         SPDM project identifier
  --help                 Show this message and exit.

第二步：
用记事本或文本编辑工具打开（在这一步不能双击打开）run.bat，配置SPDM用户名和密码。
bat文件中的内容：

.. code-block:: bash

    main.exe  --username like --password xxxxx --month 5 --project axio

参数 --username 后面是SPDM用户名。例：like；
参数 --password 后面是SPDM用户名。例：xxxxx；
参数 --month 后面是统计月份。例：5；
参数 --year 设置年份，默认统计年份是今年，如有需要可在bat文件中追加参数 --year 2019 修改统计年份为2019年；
参数 --enable-merge-cells 禁止使能单元格合并，默认禁止合并列上相同内容的单元格；
参数 --project(可选的) 如果不存在则获取全部，SPDM项目唯一标识。例：spd。

第三步：双击打开run.bat运行。
运行过程示例
Info: SPDM simulates successful login of the user
Step one: Downloading data from SPDM,please waiting....
  [####################################]  100%
Warming: Ignore project
Step three: Generating Excel,please waiting....
  [####################################]
Successfully generated, please open `2016-06-01--2016-06-30 created on 2020-06-19_10-26-43.xlsx` file under the `work table` dir

运行成功后在work tables目录下会生成工作表，如操作系统中安装有excel将会自动打开工作表。