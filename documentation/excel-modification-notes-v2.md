App is working pretty well at this point. Now trying to clean up the excel a little bit. 

There is manual redundancy between tabs on the xlsx sheet. 

1. Nothing except notes on the roster tab should be manual. It should either reference the access codes tab for player names, roles, statuses or use calculated fields from transactions to know if dues were paid or not. 

2. Events should read off of transactions for actual cost by the sum of all matching events. Who paid upfront may be multiple people so isnt necessary on this sheet. The primary purpose of this tab is to establish the events we had, their total cost, and their pay split method. 

3. Can you confirm the calculated fields for "Total Charged" on player balances are working properly? The logic should be:
- (sum of all "equal split" events) / (# people with "player" status on roster) + sum (Indicator(attended) * (cost of attendance event)) / (# attended event)

Can you also add a "support" status for non-captains and coaches helping with the team and remove "manager"?

-------------
