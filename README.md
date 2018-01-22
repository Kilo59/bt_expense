# bt_expense
Send "expense entries" held in an Excel Workbook to BigTime via REST API

[![Build Status](https://travis-ci.org/Kilo59/bt_expense.svg?branch=master)](https://travis-ci.org/Kilo59/bt_expense)
[![Coverage Status](https://coveralls.io/repos/github/Kilo59/bt_expense/badge.svg)](https://coveralls.io/github/Kilo59/bt_expense)

## BigTime REST API
http://iq.bigtime.net/BigtimeData/api/v2/help/Expense

### Expense Entry

```
HEADERS:  X-Auth-Token:{YourAPIToken}, X-Auth-Realm:{YourFirmId}
HTTP Post:  /expense/detail
POST CONTENT:  {staffsid: 123, 
		projectsid: 123, 
		catsid: 123, 
		dt: "2013-01-01", 
		CostIN: 1.25, 
		notes: "These are my expense entry notes..."}
		
		You can include ANY of the AddUpdate fields below, 
		but you MUST include the required (*) fields.

HTTP RESPONSE:  (updated ExpenseEntry object -- see below for details)
```
