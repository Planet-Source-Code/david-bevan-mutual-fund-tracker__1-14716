Mutual Fund Tracker

Updated 3/16/2001
** IMPORTANT **

I was previously using cbs.marketwatch.com but unfortunately they change the format in which the diplayed there data. This of course caused the application not to work.

I was lucky enough to stumble upon zacks.com. There are using basically the exact format that cbs.marketwatch.com was using.

Because of this occurence I am now including a saved sample page from the zacks site. That way if they ever change there site you can still see the application work and it would be easier to modifiy and understand the code

***************

1. WHAT IT DOES

The Mutual Fund Tracker application enables you to save lists of mutual funds into profiles.
You can have up to 10 profiles. Each profile can hold as many mutual fund symbols as you want - actually I have only tested about 90 funds at once and it was ok.

You add/delete funds through the manage profiles button. Here you can change the profile name, add funds, and delete funds. It is pretty straight forward.

Once you have entered your mutual fund symbols you will be able to get performance information by clicking the Get Results button. Assuming you have entered a valid mutualfund symbol, the results will be displayed in the results grid at the bottom. The performance information includes such items as 1 Week performance, 5 Year performance, etc. 

With a result set returned you can sort by clicking on the column headers. Each time you click a header it will switch from an ascending sort to a descending sort and vice versa.

Once you have results you can save the results to an excel spreadsheet. The created spreadsheet is ready to print. It should print in landscape with all the columns fitting on the page.

That is basically it.

Quick Run down

1. 10 profiles with multiple funds in each profile
2. Add/Delete funds and rename profiles through Manage profiles button
3. Get performance results through Get Results button
4. Sort Results Ascending/Descending by clicking on headers
5. Create an excel spreadsheet of the results through Save as Excel button

To Start

There is a demo profile which is load the first time you run the exe. You can just click the Get Results button to get the performance information for those two funds. I included a sample sreadsheet which is the results of the save, Demo Profile Thursday Jan 25 2001.xls.



2. WHAT IT DOES NOT DO


A. Doesn't connect you to the internet. Assumes you are already connect and won't retrieve any performance information if you are not.
B. Mutual funds only - no stocks
c. Doesn't validate the symbols you enter. If its wrong the results will show all zeros
D. It's not fastest thing around. Works pretty good on an ISDN but on a 28.8 modem get a sun dial. It also depend on the number of funds you are getting results for. Each fund is basically another web page that has to be retrieved and parsed behind the scenes.



3. SOME NERD STUFF

A. Created in VB6 SP4
B. Stores all information HKEY_LOCAL_MACHINE\SOFTWARE\MutualFundTracker
   ( Probably won't work if you do not have rights to read/write to the registry )
C. Reads/Writes to registry from a bas module - I found it some where on the internet but I can't remember where. Thanks to whoever created it
D. Uses the INet object to get the performance information.
   I am getting the information from cbs.marketwatch.com
E. Has some EXCEL VBA



4. MISC

Should be fairly bug free. If it isn't, fix it, your getting the code anyway.


5. Thanks

I can't remember where I got all the snippets of code that I tweaked, but thank you all for your generosity. I did get the code for the inet object from vbMarketwatch, which is available at wwww.planet-source-code.com.
