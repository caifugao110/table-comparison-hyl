@echo off
REM Update script for table-comparison-hyl

SET GIT_REPO=https://gitee.com/caifugao110/table-comparison-hyl.git
SET TEMP_DIR=%TEMP%\table_update_temp
SET DEST_DIR=D:\tobin
SET BAT_NAME=%~nx0

REM Create destination directory if not exists
IF NOT EXIST "%DEST_DIR%" MKDIR "%DEST_DIR%"

REM Clean up temp directory
IF EXIST "%TEMP_DIR%" RMDIR /S /Q "%TEMP_DIR%"
MKDIR "%TEMP_DIR%"

echo Testing connection to Gitee...

REM Test if git is installed
where git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Git is not installed! Please install git first.
    pause
    exit /b 1
)

echo Connected to Gitee. Cloning latest version using git...

REM Clone the repository to temp directory
git clone "%GIT_REPO%" "%TEMP_DIR%\table-comparison-hyl-master" --depth 1

IF %ERRORLEVEL% NEQ 0 (
    echo Git clone failed! Please check the repository URL or network connection.
    pause
    exit /b 1
)

echo Git clone completed successfully!

echo Extraction completed. Checking folder structure...

REM Check if the expected folder exists
IF EXIST "%TEMP_DIR%\table-comparison-hyl-master" (
    REM Create target subdirectory if not exists
    IF NOT EXIST "%DEST_DIR%\table-comparison-hyl-master" MKDIR "%DEST_DIR%\table-comparison-hyl-master"
    
    REM Copy files, excluding the bat file itself and .git directory
    ROBOCOPY "%TEMP_DIR%\table-comparison-hyl-master" "%DEST_DIR%\table-comparison-hyl-master" /E /XF "%BAT_NAME%" /XD ".git" /XX
    
    IF %ERRORLEVEL% LEQ 8 (
        echo Files copied successfully!
    ) ELSE (
        echo File copy failed!
        pause
        exit /b 1
    )
) ELSE (
    echo Expected folder 'table-comparison-hyl-master' not found in zip!
    echo Available folders:
    DIR /AD "%TEMP_DIR%"
    pause
    exit /b 1
)

echo Cleaning up temp files...
REM Clean up temporary files
RMDIR /S /Q "%TEMP_DIR%"

echo Update completed successfully!
echo The latest version has been extracted to %DEST_DIR%\table-comparison-hyl-master
pause
