**# BuildingConnected
The Source Code is written in Python and scrapes basic details like Name, Designation, Email Address and Phone Numbers from BuildingConnected Website.

I have Implemented a different approach then the one suggested in the form:

1) The Form Suggested to login first and then extract the information. 
2) But i have used another method. I created a session, and logged into the website manually. In the code, i used that chrome session to avoid login. I have used this method because during the login method, by inputting email id and password, it also sends verification code to mail id, and to avoid 2-step verification, i used existing session. 
  
3) To open the session on port 9222, paste this command in cmd:
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome_debug"

4) Chrome tab will appear, log in to website, and keep this session open while running the code.

5) I have used this techinique during my upwork's freelancing tasks, some sites are lowes.com, Websummit.qatar were login is mandatory and code is not always feasible solution. You can check this sites on my job history.

6) Another the code checks via xpaths provided and also by CSS selectors. The script will break if there is any change in website's structure.

7) Lastly, the code saves the extracted data to excel file and also displays on output section in code editor.

8) Note: This fields are temporary selected as by using disposable emails, company verfication is not done, so tender details are not visible.

9) The loom link shows that code is fully fuctional.
    Link: https://www.loom.com/share/eacc8b07c841464ab6fd20d580272e77?sid=3cc2561e-a756-4478-97cd-10b0b9deb6de**
