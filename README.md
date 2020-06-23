# Robotics-Process-Automation
In this repository I have shared some notebooks which I have used to prepare RPA BOT in my workplace to reduce the manual day to day effort.

Below are the details of the RPA BOT:

1)ISF2A:- In this notebook the code is trying to find duplicate of a given sentence. We have a excel which has been prepareped by scraping a internal website containing 1000 of ideas and description/details  of the ideas. We need to find if a given idea is a duplicate of any given ideas.

Thus what I did here is used nlp to toekenize the idea and description and using cosine similarity I tried to find the the 10 most similar idea to a given idea also highlighting the similarity percentage.

2)IBS DAILY:- In this RPA I am extracting data from multiple excel based on some specific inputs like data for specific dates, the site for which data is required, max. or min. data volume reqd. etc and pasting into my final calcation. With the help of this BOT the time required to perform the activity was reduced to 15 mins from 2hrs as it has reduced a tons of calcuation involved.

3)IBS FEB:- This RPA also perform similar kind of function but in additional it involves taking screenshot and pasting in excel, picking the data for the dates for which traffic was highest and copy paste of some specific data from notepad to excel.

4)SRS_AUTOMATE:- Handling multiple excel to prepare one consolidated excel which will contain huge amount of data selected on basis of site its azimuth, its lat and long etc.

