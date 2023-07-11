# PS-sysmon-html
This is a powershell script that monitors a Windows System and generates an HTML file with a table of the system parameters and their values.

## System Parameters

The script monitors the following system parameters:

- Event viewer errors from the last 24 hours
- Disk(s) usage percentage
- CPUs usage percentage
- Memory utilization percentage
- CPU and GPU temperature in Celsius

## Output File

The script generates an HTML file with a table that includes all the system parameters and their respective values. The output file name and path can be modified in the script. The default output file is `C:\Temp\SystemReport.html`.

## How to Run

To run the script, open a powershell window and navigate to the folder where the script is located. Then, run the following command:

```powershell
.\SystemMonitor.ps1
```

The script will run once and generate the HTML file. To run the script automatically every hour, you can use the Task Scheduler to create a scheduled task that runs the script every hour. For more information on how to use the Task Scheduler, see [this article](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page).

To send the HTML file as an email attachment, you can use the `Send-MailMessage` cmdlet in powershell. For more information on how to use the `Send-MailMessage` cmdlet, see [this article](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7.2).

## License

This script is licensed under the MIT License. See [LICENSE](LICENSE) for more details.
