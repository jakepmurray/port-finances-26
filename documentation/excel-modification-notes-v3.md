# Iter 1

We are getting very close here. Now I have a few more changes. 

1. The biggest change will be to events and attendance. I'll let you chose the best way to handle this. I like the events view, but not all events have the same participation for all transactions within that event. So we need logic similar to what exist, but there should be line items linked to bigger events where attendance can be by line item instead of by event. For example, if I add line item "Van 1" linked to the "PEC" event, that should create an attendance column for "Van 1" where I can select who was in that van. I still want to be able to see that Van 1 is linked to PEC in attendance, but it should have its own column, and the document should be flexible so that when I add line items they populate in attendance and get linked to the appropriate event. I would think adding a line items sheet is the best way to do this but i'm not 100% sure.

2. I want logic in the xlsx to toggle which events should appear in the website "my balance" breakdown for players. This will probably involve a change to both the html and the xlsx. For now, let's do it by grouping into categories of "Pre-Season 2026" and "In-Season 2026" and only showing the events that are in-season. However, this should be flexible to add more events or group differently in the future. The treasurer should be able to toggle these into their views, but no one else should have that option.

3. With change (1) enacted, players should see their line item breakdowns in their "my dashboard" view under each event. 

-------------------

# Iter 2

This wasn't your best effort. Let's go over everything.

- I think i messed up with the line items sheet. I think "transactions" is exactly what i wanted line items to be

- Attendance should read off of the line items instead of events (transactions unless you can think of a reason not to do this) but still include what the parent event is. the "description" is the line item name 

- All events are still showing up on the html, not just in-season

- The html now says that eveyone is settled up even though they aren't. 

Let's give it another go!

-------------------

# Iter 3

The html site is looking great!

For the xlsx attendance, It seems like what you're doing is close to what I want but not quite there. Instead of "Attendance cols 52–81 now pull from Transactions rows 7–36 instead of the Line Items sheet", I want the old attendance sheet to be completely replaced by this new logic instead of the new logic being tacked on at the end. 

Can you make that change and any appropriate changes to the html based on that update?


--------------------

# Iter 4

This is awesome!

One more changes for now. Coaches, captains, and treasurer should get the same dues view as players as their default view in addition to the other tabs they have access to. Support should only have the player view.