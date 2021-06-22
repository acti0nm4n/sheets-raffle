# Google Sheets raffle selector

Simple Google Sheets for creating unique tickets per user and selecting winners.

Input sheet requires first name, last name and number of tickets for that user. Columns can be redefined by the col* vars at top.

## Usage

- Create new sheet or open existing
- Select Tools > Script editor
- Copy in contents of Code.js
- Save
- Refresh Sheet, new menu item 'Raffle' will be added.
- Select 'Create tickets' to create unique tickets for each user
- Select 'Pick winner' to highlight winner, can be used multiple times, count will increase.
- Select 'Reset winnners' to clear selections and reset count.
