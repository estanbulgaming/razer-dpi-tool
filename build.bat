@echo off
echo Building RazerDPI.exe...

:: Try .NET Framework 4.x first (most common)
if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" (
    "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" /target:winexe /out:RazerDPI.exe RazerDPI.cs
    goto done
)

:: Try .NET Framework 3.5
if exist "%WINDIR%\Microsoft.NET\Framework\v3.5\csc.exe" (
    "%WINDIR%\Microsoft.NET\Framework\v3.5\csc.exe" /target:winexe /out:RazerDPI.exe RazerDPI.cs
    goto done
)

echo ERROR: .NET Framework not found!
pause
exit /b 1

:done
if exist RazerDPI.exe (
    echo.
    echo Success! RazerDPI.exe created.
    echo You can now run RazerDPI.exe
) else (
    echo.
    echo Build failed!
)
pause
