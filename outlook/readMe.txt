This code implements on the premise that

1. Your macrobook has the worksheet named "body".
2. The sheet has at least 1 cell contains some character at "A" column.
(Without any character at A column,you'll show a draft mail written none.)
3.In VBA references window,you checked "Microsoft Outlook **.* Object Library".

In the first place,This code doesn't imprement if you didn't install "Outlook Application".

If you do it,a mail is created in your draft folder of outlookApp. 
This mail body is taken from A column and pasted from the top.

