Prerequisites:
* This application will not create SQL server database by itself, so kindly create or use already crated database
* Update connection string under “ProcessDataDal” class. 
* Update Server Name, Port address, Username, Password details in “SendEmail” method of “frmProcessExcel’ class.
*  Uncomment below line in “SendEmail” method of “frmProcessExcel’ class to successfully send and email. 
o //sendEmail.Send();  
o Also enable “Less secure app” settings in Gmail account.
* Compile the application with X64 architecture.

Note:
* Log file and error file generates at same location from where application launched. 

