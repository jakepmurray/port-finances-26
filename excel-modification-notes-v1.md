this directory contains an app that's intended to be used as a team frisbee finance website. it has the following goals:

1. be intuitive and user-friendly for everyone on the team
2. have different access levels for players and the treasurer (treasurer should get a helpful dashboard summary view along with their own personal page, everyone else should only see their own finances)
3. be easy to edit
4. be easy to pass of in the future

The excel file will live on google sheets. I'll download it into the project folder and upload it to github, where the html dashboard will be hosted on github sites. 

There are some changes I want to make off the bat. 

1. I want the roster page of the spreasheet to not have jersey # or date joined. Access codes should not exist on any sheet except the access codes sheet. Also remove the summary stats at the top of the page. 
2. Events sheet shouldn't have # expected on it and shouldn't have summary stats at top. it should have start and end date instead of just date
3. Attendance sheet should read off of roster and events to fill in rows and columns. total attendees and total cost is fine but the rest of the stuff at the top should go
4. transactions "linked events" should be a dropdown that directly links to the "events" page if possible
5. player balance tab shouldnt distinguish between events and dues. they should all be the same. dues owed shouldnt have to be manually filled. it should be rules-based with the attendance and events tab. 

I don't know the best way to address this yet, but i want an easy way for players to add items they paid for without editing the master sheet. Maybe a "pending approval" tab that is easy to migrate into the master transactions sheet? 

