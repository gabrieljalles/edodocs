@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM Todo o trabalho fica no build.py. Exemplos (veja o topo do build.py):
REM   build.bat                  pacote com todos os processos
REM   build.bat Compensação      pacote so com a Compensação
REM   build.bat --separado       um pacote para cada processo
REM   build.bat --sem-planilhas  sem as planilhas (cada pessoa escolhe a sua)
REM O pacote fica em pacotes\ : descompacte e abra Edodocs\Edodocs.exe
python build.py %*
pause
