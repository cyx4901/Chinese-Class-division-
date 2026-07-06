# Balanced Random Class Allocation Tool

This tool can read an Excel student list, perform balanced random class allocation based on the allocation criteria entered by the user, and write the results back into a new worksheet in the original Excel file.

## Double-Click Usage

1. Double-click `启动分班工具.bat`.
2. The program will not automatically scan Excel files on your computer or in any folder. Please click “Select” to manually choose the student list to be used for class allocation.
3. Click “Generate Template” to create `分班名单模板.xlsx`. You may also directly select an existing registration export form.
4. In the “Class allocation criteria” field, enter `A`, `AB`, `ABC`, or `ABCD`.
5. If directed placement is required, add or fill in the columns `是否定向` and `指定班级` in the student list, for example, `Class 2`.
6. Click “Start Class Allocation”. The tool will add or update a `分班结果` worksheet in the original Excel file. It will not save a separate new file.

The tool will only add or update the `分班结果` worksheet. It will not write any content into the original student list worksheet, nor will it modify the cells, formatting, filters, or formulas in the original student list worksheet.

Class names should be separated by Chinese commas `，`, and the number of class names must match the “Number of classes”. For example, if the number of classes is set to `5`, the class names should be entered as:

```text
Class 1，Class 2，Class 3，Class 4，Class 5
```

If the number of classes does not match the number of class names, a pop-up reminder will appear after clicking “Start Class Allocation”, asking the user to add or remove class names. The program will not directly start reading the student list or allocating classes.

Class allocation criteria:

* `A`: Gender
* `B`: Household registration category, namely Shanghai household registration and non-Shanghai household registration
* `C`: Detailed household registration category, namely local household registration with residence consistency, local household registration with residence separation, non-Shanghai household registration with sufficient points, and non-Shanghai household registration with insufficient points
* `D`: Child personnel type

For example, entering `A` will balance only male and female students; entering `AB` will balance gender and household registration category; entering `ABC` will balance gender and detailed household registration category at the same time; entering `ABCD` will additionally include child personnel type.

## Recommended Class Allocation Criteria: ABC

If the class allocation requirement is to “balance students according to gender and household registration information, with household registration further divided into local household registration with residence consistency, local household registration with residence separation, non-Shanghai household registration with sufficient points, and non-Shanghai household registration with insufficient points”, it is recommended to enter the following in the “Class allocation criteria” field:

```text
ABC
```

`ABC` means:

* `A = Gender`
* `B = Household registration category`
* `C = Detailed household registration category`

Where `B = Household registration category` includes:

* Shanghai household registration
* Non-Shanghai household registration

Where `C = Detailed household registration category` includes:

* Local household registration with residence consistency
* Local household registration with residence separation
* Non-Shanghai household registration with sufficient points
* Non-Shanghai household registration with insufficient points

After entering `ABC`, the program will try to perform balanced random class allocation based on combinations of “gender + household registration category + detailed household registration category”. For example:

* Male + Shanghai household registration + local household registration with residence consistency
* Female + Shanghai household registration + local household registration with residence separation
* Male + non-Shanghai household registration + non-Shanghai household registration with sufficient points
* Female + non-Shanghai household registration + non-Shanghai household registration with insufficient points

Specific class allocation method:

1. Directed placement students are processed first and directly assigned to their specified classes.
2. Other students are grouped by `gender + household registration category + detailed household registration category`.
3. Students within each group are randomly shuffled.
4. Students from the same group are then distributed to each class as evenly as possible.
5. If a random seed is entered, the same class allocation result can be reproduced when the student list and settings remain unchanged.

Differences between different inputs:

* Entering `A`: Balances only gender, which does not meet the full household registration balancing requirement.
* Entering `AB`: Balances gender and household registration category, but does not further divide students into local household registration with residence consistency, local household registration with residence separation, non-Shanghai household registration with sufficient points, and non-Shanghai household registration with insufficient points.
* Entering `ABC`: Exactly meets the full balancing requirement of “gender + household registration category + detailed household registration category”.
* Entering `ABCD`: Adds “child personnel type” balancing on top of `ABC`.

The original Excel file will add or update a `分班结果` worksheet, which includes:

* `Class allocation results`: the final class assigned to each student.
* A student list sorted by class.
* Rows for students who did not participate in class allocation at the bottom of the worksheet.
* Class statistics, stratification details, running information, exception records, and skipped-record explanations.
* Class lists displayed by class sections: `Class 1` is displayed first with its students listed below, followed by `Class 2`, and so on.
* At the end of each class list, proportion pie charts will be generated based on the selected class allocation criteria. For example, selecting `A` will generate a gender ratio chart, while selecting `ABC` will generate gender, household registration category, and detailed household registration category ratio charts.
* All proportion charts for the same class will be arranged horizontally in one row.
* Chart information will be displayed directly on the chart, without using extra rows for “category / number of students” tables.
* Pie chart size and spacing will be automatically adjusted based on the number of categories and the number of charts in the same row, so that labels are as clear as possible while limiting the maximum chart size to avoid oversized charts.
* Class statistics, class allocation notes, running information, exception records, and skipped-record explanations will all be kept as clear tables.
* If `maximum students per class × number of classes` cannot accommodate all participating students, additional classes will be automatically created, and the reason will be written in the first row of the worksheet.

Note: Since the results need to be written back into the original Excel file, please close the student list file before clicking “Start Class Allocation”.

Data worksheet protection rule: The original student list worksheet is used only for reading, not for writing. When class allocation is run again, the program will only replace the old `分班结果` worksheet and will not overwrite the original student list.

Privacy rule: The program will not automatically scan existing Excel files on the local computer, nor will it read historical files. It will only read the selected file after the user manually selects a student list and clicks “Start Class Allocation”.

## Dependency Installation and Common Errors

`requirements.txt` is the list of Python dependencies required to run the program. It is not a student list and does not store any student information. The main dependency currently used is `openpyxl`, which is used to read and write Excel files and generate charts.

If the dependencies are not installed, the following errors may occur:

* `ModuleNotFoundError: No module named 'openpyxl'`
* `ImportError: No module named openpyxl`
* `Missing openpyxl. Please install it first: python -m pip install openpyxl`
* After double-clicking `启动分班工具.bat`, the window flashes and closes immediately, or it reports that a Python module is missing.

Solution:

1. In the folder where this tool is located, hold down `Shift`, right-click a blank area, and select “Open PowerShell window here”.
2. Run:

```powershell
py -m pip install -r requirements.txt
```

If the computer does not support the `py` command, use:

```powershell
python -m pip install -r requirements.txt
```

If it says `pip` does not exist, run the following first:

```powershell
py -m ensurepip --upgrade
py -m pip install -r requirements.txt
```

If it says `py` or `python` is not recognized as an internal or external command, it means Python may not be installed on the computer, or Python was not added to PATH during installation. Python 3 needs to be installed first, and then the installation commands above should be run again.

## Command Line Usage

Generate a template:

```powershell
python class_splitter.py --create-template 分班名单模板.xlsx
```

Perform class allocation and write the result back to the `分班结果` worksheet in the original Excel file by default:

```powershell
python class_splitter.py 分班名单模板.xlsx --class-count 4
```

Perform class allocation based on specified criteria:

```powershell
python class_splitter.py 学生名单.xlsx --balance ABCD
```

Specify class names:

```powershell
python class_splitter.py 学生名单.xlsx --classes "Class 1,Class 2,Class 3,Class 4"
```

Set a fixed random seed so that the result can be reproduced:

```powershell
python class_splitter.py 学生名单.xlsx --seed 20260627
```

Set the maximum number of students per class:

```powershell
python class_splitter.py 学生名单.xlsx --capacity 36
```

## Student List Fields

Basic required columns:

* `姓名`
* `性别`: Male, Female

If `B` or `C` is selected, the tool will automatically determine household registration information from the following fields first:

* `户籍地址(省)`
* `儿童人户分离登记情况`
* `儿童人员类型`
* `居住证类型`
* `报名情况`
* `分配验证点`
* `备注1` to `备注4`

The following columns can also be added or filled in directly:

* `户籍大类`: Shanghai household registration, non-Shanghai household registration
* `户籍细类`: local household registration with residence consistency, local household registration with residence separation, non-Shanghai household registration with sufficient points, non-Shanghai household registration with insufficient points

Optional columns:

* `是否定向`: Yes, No
* `指定班级`
* `备注`

## Self-Test

```powershell
python class_splitter.py --self-test
```
