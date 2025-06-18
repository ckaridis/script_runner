@echo off
@REM Just a small script that opens powershell and stop all Docker containers.

@REM Author: https://github.com/ckaridis

call powershell docker stop $(docker ps -aq)

pause;
