###
# remove old messages for all local msmq queues
# msmq = microsoft messaging queuing
#
# 20190530 122800
#   add options $showMsg and $dryRun
# 20190521 122800
#   initial
#

#
# config
#
[System.Reflection.Assembly]::LoadWithPartialName("System.Messaging") | Out-Null
$utf8 = new-object System.Text.UTF8Encoding
$now = Get-Date
# remove messages older than
# $old = $now.AddDays(-1)
$old = $now.AddMinutes(-10)
# show details of message to be removed
$showMsg = 1
# run without removing
$dryRun = 1

# 
# run
#

# get queues
$queuePaths = Get-MsmqQueue -QueueType Private | Select -Property QueueName;
if ($dryRun)
{
    echo "dry run; would be "
}
echo "removing old messages for all local msmq queues"
echo ""
echo "nbr queues: $($queuePaths.Length); checking messages older than $($old)"
$queueCounts = Get-MsmqQueue -QueueType Private | Format-Table -Property QueueName,MessageCount;
echo $queueCounts
echo ""
pause
foreach ($queuePath in $queuePaths)
{
    # for proper permissions, prepend .\
    $localQueuePath = ".\$($queuePath.QueueName)"
    echo "queue: $localQueuePath"
	$queue = New-Object System.Messaging.MessageQueue $localQueuePath
    
    # to read ArrivedTime property
    $queue.MessageReadPropertyFilter.SetAll()
    
    # get snapshot of all messages, but uses memory, and slower
	# $msgs = $queue.GetAllMessages()
    # echo "  $($msgs.Length) messages"
    # get cursor to messages in queue
    $msgs = $queue.GetMessageEnumerator2()
    
    # add a message so can test
    # $queue.Send("<test body>", "test label")
    # pause
    $removed = 0
	# foreach ($msg in $msgs)
    while ($msgs.MoveNext([timespan]::FromSeconds(1)))
	{               
        $msg = $msgs.Current
		if ($msg.ArrivedTime -and $msg.ArrivedTime -lt $old)        
		{
            if ($showMsg)
            {
                echo "--------------------------------------------------"
                echo "ArrivedTime: $($msg.ArrivedTime)"
                echo "BodyStream:"
                if ($msg.BodyStream) 
                {
                    echo $utf8.GetString($msg.BodyStream.ToArray())
                }
                echo "Properties:"
                echo $msg | Select-Object
                echo ""
            }
            
            try {
                if (!$dryRun)
                {
                    # receive ie remove message by id from queue
                    $queue.ReceiveById($msg.Id, [timespan]::FromSeconds(1))
                }
                $removed++
            } catch [System.Messaging.MessageQueueException] {
                $errorMessage = $_.Exception.Message
                # ignore timeouts                
                if (!$errorMessage.ToLower().Contains("timeout"))
                {
                    throw $errorMessage
                }
            }
		}
	} 
    if ($dryRun)
    {
        echo "dry run; would have "
    }
    echo "removed $removed messages from $localQueuePath"
    echo ""
}

pause
#
###
