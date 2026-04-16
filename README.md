# Excel Data Checker

#Project Overview
This is an office automation script based on Python, used to quickly scan Excel files.

#Main Features
* **Automatic Detection**: Automatically detects empty values (None) or negative numbers in the first column.
* **Highlighting**: In a newly created copy file, marks error rows as "Wrong Data" and sets them to **bold red** font.
* **Data Safety**: Does not modify the original file, automatically generates a copy named `correct_file_copy.xlsx`.

# Tools Used
* **Python 3.x**
* **Libraries**: `openpyxl`, `pathlib`

# Learning Outcomes
This is my first Python automation project. Through this project, I learned how to:
1. Use `pathlib` to handle file paths.
2. Use `openpyxl` to read, write, and style Excel files.
3. Use `isinstance` for strict data type validation.

# Excel 数据自动校验工具 (Excel Data Checker)

# 项目简介
这是一个基于 Python 的办公自动化脚本，用于快速扫描 Excel 文件。

# 主要功能
* **自动识别**：自动检测第一列中的空值（None）或负数。
* **高亮标记**：在新建的副本文件中，将错误行标记为“Wrong Data”并设置为**红色加粗**字体。
* **数据安全**：不修改原文件，自动生成一个 `correct_file_copy.xlsx` 副本。

# 使用工具
* **Python 3.x**
* **库**: `openpyxl`, `pathlib`

# 学习收获
这是我的第一个 Python 自动化项目，通过这个项目我掌握了：
1. 如何使用 `pathlib` 处理文件路径。
2. 使用 `openpyxl` 进行 Excel 的读写与样式设置。
3. 运用 `isinstance` 进行严格的数据类型校验。
