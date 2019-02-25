# Connect Microsoft 365 MFA
#### Create a PowerShell profile that allows you to connect to all Microsoft services in PowerShell using MFA.
---
Copy all files from the repo into the following path on your Windows 10 PC: 

`%userprofile%\Documents\WindowsPowerShell`

Then open a new PowerShell window and run the following function to install all Microsoft 365 PowerShell modules:

**`install-m365`**

Now you can just type **`m365`** to connect to the following services. Or use the shortcuts below to connect to a single service
* Azure **`azure`**
* Azure AD **`mso`**
* Azure Rights Management **`m365`**
* Exchange Online **`exo`**
* Microsoft Online (MSOL) **`mso`**
* Power BI **`m365`**
* Security and Compliance Center **`m365`**
* SharePoint Online **`m365`**
* Skype **`sbo`**
* Teams **`teams`**

Run `NOM365` to end all sessions before closing your PowerShell window.

***
  
MIT License
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the ""Software""), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Copyright (c) Dan Chemistruck 2019. All rights reserved.
