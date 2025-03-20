
:: Define the bold PSG
set "P1=	__________"
set "P2=	\______   \"
set "P3=	 |     ___/"
set "P4=	 |    |    "
set "P5=	 |____|    "
set "P6=	           "

set "S1=  _________"
set "S2= /   _____/"
set "S3= \_____  \ "
set "S4= /        \"
set "S5=/_______  /"
set "S6=        \/ "

set "G1=  ________ "
set "G2= /  _____/ "
set "G3=/   \  ___ "
set "G4=\    \_\  \"
set "G5= \______  /"
set "G6=        \/ "


:: Slide the text
for /l %%i in (10,-1,0) do (
    cls
    echo  ^> ^> ^> ^> ^> ^> ^> ^> 
    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
    echo                 !P1:~%%i! !S1:~%%i! !G1:~%%i!
    echo                 !P2:~%%i! !S2:~%%i! !G2:~%%i!
    echo                 !P3:~%%i! !S3:~%%i! !G3:~%%i!
    echo                 !P4:~%%i! !S4:~%%i! !G4:~%%i!
    echo                 !P5:~%%i! !S5:~%%i! !G5:~%%i!
    echo                 !P6:~%%i! !S6:~%%i! !G6:~%%i!
    ping 127.0.0.1 -n 1 -w 200 >nul
)

    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
    echo.
echo					                               ^< ^< ^< ^< ^< ^< ^< ^<



