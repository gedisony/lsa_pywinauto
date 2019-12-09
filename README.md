# lsa_pywinauto

One more example to using pywinauto to automate stuff in windows. Hope this helps someone.

We have Lawson HR at work. I wrote this tool to make life easier. 

What it does:
  * Launch Lawson Security Administrator(LSA) and login with my credentials
    * open to the user maintenance screen
    * show several users' credentials
  * Automate running security audit
    * Run report job in LSA for all user roles
    * Launch Lawson Interface Desktop export sso information
