# Python Bulk Message Sender for Reddit

##### *For educational purposes only.*

This is a Python script written for the purpose of sending messages to
users in bulk. Reddit imposes limitations on sending messages
*(for obvious reasons)*, and therefore this script pauses for two hours
after every 100 messages are sent out.

The script uses an .xlsx file as a source for users to send the message to.
It then uses results.xlsx to specify which accounts had successfully had
a message sent to them.

There are also two .txt files used as sources for both the message body
and the message headline. **All of these files have been provided as samples.**

#### Sample Message

![Sample message in Reddit inbox](sample-message.png "Sample message in Reddit inbox")

*This message is provided in sample-message.txt - You can use Reddit markup in the .txt files.*

#### Credits
- [praw](https://praw.readthedocs.io/en/latest/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [python-dotenv](https://github.com/theskumar/python-dotenv)
