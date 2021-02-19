# Functions
# This respository contains functions to help complete everyday tasks. A list of the functions is provided below
# 
# Function Send-OutlookEmail
#     Use: This function will send an email from your outlook on you local system as the currently logged in user. 
#          This can be used to send messages with the following options:
#               Request Read Receipts
#               Message Priority Low,Normal,High
#               BodyIsHTML - Will send the message as HTML 
#     Syntax: 
#     Send-OutlookEmail -To someone@outlook.com -Subject "Testing" -Body "Just a test this time" -MessagePriority high -RequestReadReceipt
