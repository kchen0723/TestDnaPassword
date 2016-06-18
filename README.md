# TestDnaPassword
Test Excel DNA issue

Please download this project, compile it and then you can reproduce this bug.

Reproduce with TestPasswordWorkBook.xlsm  
1) open TestPasswordWorkBook.xlsm  
2) drug the TestDnaPassword.xll to excel to load the addin  
3) click the button "Button1", a form will be popup  
4) close the form  
5) close excel.  
6) a password dialog will be popup  

Reproduce with TestPassword.xlsm  
1) open TestPassword.xlsm  
2) drug the TestDnaPassword.xll to excel to load the addin  
3) click the "TestPassword" tab, then click "click to test". a form will be popup  
4) close the form  
5) close excel.  
6) a password dialog will be popup  

Both of the TestPassword.xlsm and TestPasswordWorkBook.xlsm are almost empty, only have a user form. 
Their password is 123456. you can open it if needed.