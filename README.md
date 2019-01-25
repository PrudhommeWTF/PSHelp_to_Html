---
external help file: HelpToHtml-help.xml
Module Name: HelpToHtml
online version: https://github.com/PrudhommeWTF/PSHelp_to_Html
schema: 2.0.0
---

# New-HtmlHelp

## SYNOPSIS
New-HtmlHelp will create an HTML file of Comment Based Help written for specified PowerShell Cmdlets, Scripts or Modules.

## SYNTAX

### Module (Default)
```
New-HtmlHelp -Module <String> [-RelatedHelp <ArrayList>] [-OutputFolder <String>] [<CommonParameters>]
```

### Command
```
New-HtmlHelp -Command <String[]> [-RelatedHelp <ArrayList>] [-OutputFolder <String>] [<CommonParameters>]
```

### Script
```
New-HtmlHelp -Script <String> [-RelatedHelp <ArrayList>] [-OutputFolder <String>] [<CommonParameters>]
```

## DESCRIPTION
This PowerShell Function will create HTML files based on items provided Help.
The Funcion use Comment based help written by Functions, Cmdlets, Scripts or Modules provider to create a proper HTML file.

## EXAMPLES

### EXAMPLE 1
```
New-HtmlHelp -Module ActiveDirectory
```

Will create a local HTML file containing all the Cmdlet in ActiveDirectory module and their Comment Based Help content.

### EXAMPLE 2
```
New-HtmlHelp -Module ActiveDirectory -RelatedHelp @(
```

@{
        Title = 'Title'
        UriOrFilePath = 'UriOrFilePath'
    }
)

Will create a local HTML file containing all the Cmdlet in ActiveDirectory module and their Comment Based Help content, with a link with the specified title pointing to the specified path ou Uri.

### EXAMPLE 3
```
New-HtmlHelp -Command 'New-Item','Remove-Item'
```

Will create a local HTML file containing Comment Based Help of Cmdlets New-Item and Remove-Item.

### EXAMPLE 4
```
New-HtmlHelp -Command 'New-Item','Remove-Item' -RelatedHelp @(
```

@{
        Title = 'Title'
        UriOrFilePath = 'UriOrFilePath'
    }
)

Will create a local HTML file containing Comment Based Help of Cmdlets New-Item and Remove-Item, with a link with the specified title pointing to the specified path ou Uri.

### EXAMPLE 5
```
New-HtmlHelp -Script .\test.ps1
```

Will create a local HTML file containing Comment Based Help of the PowerShell Script 'test.ps1'.

### EXAMPLE 6
```
New-HtmlHelp -Script .\test.ps1 -RelatedHelp @(
```

@{
        Title = 'Title'
        UriOrFilePath = 'UriOrFilePath'
    }
)

Will create a local HTML file containing Comment Based Help of the PowerShell Script 'test.ps1', with a link with the specified title pointing to the specified path ou Uri.

## PARAMETERS

### -Module
Module Name or Path to the PSM1 file.

```yaml
Type: String
Parameter Sets: Module
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Command
Names of the PowerShell Cmdlets or Functions to create HTML Help for.

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
Path to the script to create HTML Help for.

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

### -RelatedHelp
Allows you to add links to outside folders ou URL in the HTML output.
Parameter value needs to be an ArrayList of Hastables composed as follow:
    @{
        Title = 'Title'
        UriOrFilePath = 'UriOrFilePath'
    }

```yaml
Type: ArrayList
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFolder
Specify the path to write the HTML file in.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: .
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String
## OUTPUTS

### System.IO.File
## NOTES
Author: Thomas Prud'homme (Blog: https://blog.prudhomme.wtf Tw: @Prudhomme_WTF).

## RELATED LINKS

[https://github.com/PrudhommeWTF/PSHelp_to_Html](https://github.com/PrudhommeWTF/PSHelp_to_Html)

