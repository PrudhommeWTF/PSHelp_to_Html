---
external help file: HelpToHtml-help.xml
Module Name: HelpToHtml
online version:
schema: 2.0.0
---

# New-HtmlHelp

## SYNOPSIS
New-HtmlHelp will create an HTML file of helps written for specified PowerShell Cmdlets, Scripts or Modules.

## SYNTAX

### Module (Default)
```
New-HtmlHelp -Module <String> [-OutputFolder <String>] [-WithModulePage] [<CommonParameters>]
```

### Command
```
New-HtmlHelp -Command <String[]> [-OutputFolder <String>] [-WithModulePage] [<CommonParameters>]
```

### Script
```
New-HtmlHelp -Script <String> [-OutputFolder <String>] [-WithModulePage] [<CommonParameters>]
```

## DESCRIPTION
This PowerShell Function will create HTML files based on items provided Help.
The Funcion use Comment based help written by Functions, Cmdlets, Scripts or Modules provider to create a proper HTML file.

## EXAMPLES

### EXAMPLE 1
```
New-HtmlHelp -Module Value
```

Describe what this call does

### EXAMPLE 2
```
New-HtmlHelp -Command Value
```

Describe what this call does

### EXAMPLE 3
```
New-HtmlHelp -Script Value
```

Describe what this call does

### EXAMPLE 4
```
New-HtmlHelp -OutputFolder Value -WithModulePage
```

Describe what this call does

## PARAMETERS

### -Module
Module Name or Path to the .

```yaml
Type: String
Parameter Sets: Module
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Command
Describe parameter -Command.

```yaml
Type: String[]
Parameter Sets: Command
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Script
Describe parameter -Script.

```yaml
Type: String
Parameter Sets: Script
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFolder
Describe parameter -OutputFolder.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: .\docs
Accept pipeline input: False
Accept wildcard characters: False
```

### -WithModulePage
Describe parameter -WithModulePage.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: True
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### List of input types that are accepted by this function.
## OUTPUTS

### List of output types produced by this function.
## NOTES
Place additional notes here.

## RELATED LINKS

[URLs to related sites
The first link is opened by Get-Help -Online New-HtmlHelp]()

