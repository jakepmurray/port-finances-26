# Iter 1

App is working pretty well at this point. Now trying to clean up the excel a little bit. 

There is manual redundancy between tabs on the xlsx sheet. 

1. Nothing except notes on the roster tab should be manual. It should either reference the access codes tab for player names, roles, statuses or use calculated fields from transactions to know if dues were paid or not. 

2. Events should read off of transactions for actual cost by the sum of all matching events. Who paid upfront may be multiple people so isnt necessary on this sheet. The primary purpose of this tab is to establish the events we had, their total cost, and their pay split method. 

3. Can you confirm the calculated fields for "Total Charged" on player balances are working properly? The logic should be:
- (sum of all "equal split" events) / (# people with "player" status on roster) + sum (Indicator(attended) * (cost of attendance event)) / (# attended event)

Can you also add a "support" status for non-captains and coaches helping with the team and remove "manager"?

-------------

# Iter 2

Follow-up notes:

1. On roster tab, I am not seeing any formulas from Dues Amount ($). This should read from "Total Charged" on player balances tab. 

2. I think "pending approval" should be its own spreadsheet that can be copy and pasted directly into this one. It should be exactly the same as the transactions sheet just with an "approved" tab on it. It will be the only sheet that players interact with.

3. The "events" section of the dashboard tab is messed up. Please make sure this calculates correctly. Dont change the other sheets, just dashboard sheet.

4. I reiterate note #3 above: 3. "Total Charged" on player balances is not working properly. The logic should be:
- (sum of all "equal split" events) / (# people on roster) + sum (Indicator(attended) * (cost of attendance event)) / (# attended event)

Please completely redo these formulas and make sure they're correct. 

-------------

# Iter 3

1. The "events" section is still wrong. Please completely redo this section with accurate values. 

2. The redone logic for "total charged" now is not recording amounts owed for anyone. Remember -- the total amount owed should be an equal split among everyone for "equal split" and equal split among those attending for "attendance"

-------------

# Iter 3

I may have confused you. 
