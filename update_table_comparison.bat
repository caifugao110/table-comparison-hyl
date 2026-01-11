@echo off
REM table-comparison-hyl 更新工具
chcp 936 >nul
TITLE table-comparison-hyl 更新工具
COLOR 0A
CLS

ECHO ============= table-comparison-hyl 更新工具 =============
ECHO.

SET GIT_REPO=https://gitee.com/caifugao110/table-comparison-hyl.git
SET TEMP_DIR=%TEMP%\table_update_temp
SET DEST_DIR=D:\tobin
SET BAT_NAME=%~nx0

IF NOT EXIST "%DEST_DIR%" MKDIR "%DEST_DIR%"

IF EXIST "%TEMP_DIR%" RMDIR /S /Q "%TEMP_DIR%"
MKDIR "%TEMP_DIR%"

ECHO 1/5 测试连接...
where git >nul 2>nul || (
    COLOR 44
    ECHO Git未安装！
    pause
    exit /b 1
)

ECHO 2/5 克隆仓库...
git clone "%GIT_REPO%" "%TEMP_DIR%\table-comparison-hyl-master" --depth 1 >nul 2>nul || (
    COLOR 44
    ECHO 克隆失败！
    pause
    exit /b 1
)

ECHO 3/5 检查结构...
IF EXIST "%TEMP_DIR%\table-comparison-hyl-master" (
    REM 清空目标文件夹
    IF EXIST "%DEST_DIR%\table-comparison-hyl-master" RMDIR /S /Q "%DEST_DIR%\table-comparison-hyl-master"
    MKDIR "%DEST_DIR%\table-comparison-hyl-master"
    
    ECHO 4/5 复制文件...
    ROBOCOPY "%TEMP_DIR%\table-comparison-hyl-master" "%DEST_DIR%\table-comparison-hyl-master" /E /XF ".gitignore" "from\.gitkeep" /XD ".git" /XX /NFL /NDL /NJH /NJS >nul
    
    IF %ERRORLEVEL% GTR 8 (
        COLOR 44
        ECHO 复制失败！
        pause
        exit /b 1
    )
) ELSE (
    COLOR 44
    ECHO 结构错误！
    pause
    exit /b 1
)

ECHO 5/5 清理文件...
RMDIR /S /Q "%TEMP_DIR%"

ECHO ============= 更新完成！ =============
ECHO 已更新到：
ECHO %DEST_DIR%\table-comparison-hyl-master
ECHO =======================================

pause

