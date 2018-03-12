
![](https://content.screencast.com/media/5b47d1de-0000-433b-a902-2b8052720ada_9d700cb2-87df-433c-8403-c813c6a51c87_static_0_0_2017-12-29_01-55-33.png)  
# PQGoogleAnalytics - Google Analytics connector for Power BI and Power Query for Excel

This program is written on M# language and it is used for retrieving data from Google Analytics api directly to Power BI Excel using [Power Query Addon](https://www.microsoft.com/en-us/download/details.aspx?id=39379). 

A lot of code \[in this application\] is taken from [article](http://kohera.be/blog/detail/how-to-get-google-analytics-data-in-power-query) on kohera.be. 

You can download Xls and PBIX file from [here](https://github.com/40-02/GoogleAnalyticsViaPowerQueryForExcel/releases)

## Quick setup guide

0. Open the settings of Power BI, go to "Privacy" tab, choose the setting "Always ignore the Privacy Level settings".

1. Open the link below in your browser, hit "Allow" button and copy the given token.

https://accounts.google.com/o/oauth2/auth?scope=https://www.googleapis.com/auth/analytics.readonly&response_type=code&access_type=offline&redirect_uri=urn:ietf:wg:oauth:2.0:oob&approval_prompt=force&client_id=155297956885-088obc06926s8kolll6kdqkd9u842n56

2. Open Power Query by hitting "Edit Queries" button on "Home" tab on the ribbon of Power BI.

3. Paste the token from the step 1 into "authToken" parameter.

4. Go to "getRefreshToken" query and Copy the token from your preview window.

5. Paste the token from the step 5 into "refreshToken" parameter.

6. Your Power BI is ready to get a data from Google Analytics.

## Requirements ##

1. MS Excel 2010, 2013, 2016 or Power BI
2. [Power Query Addon](https://www.microsoft.com/en-us/download/details.aspx?id=39379)

## Ignore privacy levels ##

To make this file work you need to enable the setting "ignore privacy levels" on privacy tab in Power BI or Power Query for Excel settings. Here is the screenshot:

![](http://content.screencast.com/media/9eac1f74-8980-4a7c-9042-4d189fd08a99_9d700cb2-87df-433c-8403-c813c6a51c87_static_0_0_2016-12-08_11-19-28.png)

## Video demonstration  ##

Here is the short video which describes how this workbook works. 

https://vimeo.com/maximuvarov/googleanalyticsviapowerqueryformsexcel

![](https://www.evernote.com/l/AAnq3Tra0TNMGrEb8ouN4BqL-ACyIbHeeJgB/image.png)
