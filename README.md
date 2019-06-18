Powershell script to clear MSMQ of old messages
===================================

Iterates over your private queues and removes "old" messages.

Installation
------------

Right click on clearMSMQ.ps1 and choose Run with Powershell


Usage
-----
Open clearMSMQ.ps1 and edit the options
```
# options 

# remove messages older than
# $old = $now.AddDays(-1)
$old = $now.AddMinutes(-10)

# show details of message to be removed
$showMsg = 1

# run without removing
$dryRun = 1

```
