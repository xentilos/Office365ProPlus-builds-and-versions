# How to use

Edit first line and replace builds:
```powershell
$build = @("16.0.11929.20752", "16.0.11929.20776", "16.0.12527.20612", "16.0.12730.20236", "16.0.12730.20250", "16.0.12730.20270")
```
You can also use a csv file as a input:
```powershell
$build = import-csv "C:\path\to\file\with\builds.csv"
```

The result will be show in console, but if you woudl like to save it to file(prefered) edit list line and replace it with:
```powershell
$tab | export-csv "C:\path\to\file\output.csv" -NoTypeInformation -Delimiter ";"
```