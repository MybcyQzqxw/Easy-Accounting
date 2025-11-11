@echo off
:: Check if students_data.txt exists
if not exist "students_data.txt" (
    echo.
    echo ========================================
    echo 警告：未找到 students_data.txt 文件！
    echo ========================================
    echo.
    echo 该文件包含学生数据，是程序正常运行所必需的。
    echo 请参考 students_data.txt.example 创建该文件。
    echo.
    echo 文件格式：每行一个学生，姓名和学号用制表符分隔
    echo 示例：
    echo 张三	3120230001
    echo 李四	3120230002
    echo.
    echo 按任意键退出打包过程...
    pause >nul
    exit /b 1
)

:: Clean old build files and directories
echo Cleaning old build files...

if exist "build" (
    echo Deleting build directory...
    rmdir /s /q "build"
)

if exist "dist" (
    echo Deleting dist directory...
    rmdir /s /q "dist"
)

if exist "myenv" (
    echo Deleting myenv directory...
    rmdir /s /q "myenv"
)

if exist "Easy Accounting.spec" (
    echo Deleting Easy Accounting.spec file...
    del /f /q "Easy Accounting.spec"
)

echo Cleanup complete!
echo.

:: Create virtual environment
echo Creating virtual environment...
python -m venv myenv

:: Activate virtual environment
call myenv\Scripts\activate

:: Upgrade pip to the latest version
python.exe -m pip install --upgrade pip

:: Install dependencies
pip install -r requirements.txt

:: Install pyinstaller
pip install pyinstaller

:: Use pyinstaller to package the program
pyinstaller --noconsole --onefile --icon="favicon.ico" --add-data "favicon.ico;." --add-data "style.qss;." --add-data "students_data.txt;." "Easy Accounting.py"