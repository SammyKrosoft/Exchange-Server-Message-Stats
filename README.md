# Exchange-Server-Message-Stats
This is a script that dumps the message stats from an OnPrem Exchange environment. It's taken from Microsoft's Risk Assessment bundle.

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
