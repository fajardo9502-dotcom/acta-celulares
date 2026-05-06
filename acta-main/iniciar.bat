@echo off
echo Compilando Servidor.java...
javac -cp "libs/*" Servidor.java
if %ERRORLEVEL% NEQ 0 (
    echo ERROR al compilar. Verifica que tengas Java JDK instalado.
    pause
    exit /b 1
)
echo Iniciando servidor...
java -cp ".;libs/*" Servidor
pause
