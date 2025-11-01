That is a smart auto-working sheduler that must help you to improve and make the job easier for the most of the people.
If you have any ideas please share with me. DM to me(Georgiy Kraynik)
That part necessary and important be sure after importing the code add a triggers (it makes it work, later I will attach the link with a screen recorded instructions)
It's necessary to deploy triggers and set them up correctly to make a scheduler works
RebuildAndApplyDuty - time based - minutes timer - every minute
handleEdit - from spreadsheet - on edit
updateDutyStatus - time based - minutes timer - every minute
autoResetLogs - time driven - day timer - midnight to 1 am(mostly depends on your shift day-or-night)
autoInsertBreaks - time driven - minutes timer - every minute
sendDailyLogByMail - time driven - day timer - 9am to 10 am(set it right after your work finishes)
exportMonthStats - from spreadsheet - on edit
