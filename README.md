# About WUT (Windows Update Tool)

This is a tool that allows you to install Windows updates using the PowerShell console interactively.

# Features

By default, this tool will search for new updates available, download and install them all, but it will not automatically restart the computer so that users do not lose their unsaved work.

| Option | Description |
| --- | --- |
| `-Reboot` | Force a reboot to complete the installation of any update, if needed. |
| `-SearchOnly` | Search for new updates and stop execution after showing the list of available updates. |
| `-DownloadOnly` | Search and download new available updates and stop execution. |
| `-ResetWindowsUpdate` | Reset all Windows Update components before trying to install any new available updates. |
| `-ShowUpdateHistory` | Displays a list of all updates previously installed. |

>[!NOTE]
>This tool does not reboot the computer by default, even if it is needed to complete the installation of new updates. If you want it to reboot automatically after installing the updates, use the **-Reboot** option. 

# Running remotely with PSExec
[PSExec](https://learn.microsoft.com/en-us/sysinternals/downloads/psexec) is a nice tool from Microsoft to allow execution of commands remotely.

```
PsExec.exe \\<RemoteComputername> -u <Username> -p <Password> -s powershell \\<FileServername>\<SharedFoldername>\wut.ps1 -reboot
```

[More about PSExec](https://petri.com/psexec/)


## Support

I develop most of my code under open licenses, free of charge. 
You can express your support by [making a donation](https://tppay.me/lr9flucg).

