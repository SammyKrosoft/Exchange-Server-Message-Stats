# Exchange-Server-Message-Stats
This is a script that dumps the message stats from an OnPrem Exchange environment. It's taken from Microsoft's Risk Assessment bundle.

There is another updated version of this script here on Marco Roth's Github :
[Message Stats script from Marco Roth](https://github.com/msftmroth/MessageStats)

- Simple execution but limited to dates in the past from today

- Quick sample: ``` .\MessageStats.ps1 1000 ```

Alternatively, you can also use the script from the Microsoft Exchange Team Blog:
[Exchange Team Blog - Generating User Message Profile](https://techcommunity.microsoft.com/t5/exchange-team-blog/generating-user-message-profiles-for-use-with-the-exchange/ba-p/610916)

- For convenience, I also forked a copy of this script in this repository

- Better granularity than the MessageStats.ps1 script

- Quick Sample: ``` Generate-MessageProfile.ps1 -ADSites "EastDC1","West*" -StartOnDate 01/15/2024 -EndBeforeDate 01/31/2024 -OutCSVFile MultiSites.CSV -ExcludeHealthData ```


# Usage
```powershell
.\MessageStats.ps1 <Num days back>
```

With ```<Num days back>``` being the number of days we want to go back to check the message statistics

Example to cover the 1000 last days:
```powershell
.\MessageStats.ps1 1000
```

![image](https://user-images.githubusercontent.com/33433229/167988095-843c50db-18df-4d91-82aa-ff9c5ffa4a84.png)

# Download
Download either from this repository'(see the list of files above), or right-click "Save As" from [this link](https://raw.githubusercontent.com/SammyKrosoft/Exchange-Server-Message-Stats/main/messagestats.ps1)
