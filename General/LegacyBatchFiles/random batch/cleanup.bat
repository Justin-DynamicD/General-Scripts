@echo off
FOR /f %%A in (Irvine_YN.txt) DO netuser -c %%A /s:IrvineR
:End
