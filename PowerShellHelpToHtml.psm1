Function New-HtmlHelp {
    <#
        .SYNOPSIS
        New-HtmlHelp will create an HTML file of Comment Based Help written for specified PowerShell Cmdlets, Scripts or Modules.

        .DESCRIPTION
        This PowerShell Function will create HTML files based on items provided Help.
        The Funcion use Comment based help written by Functions, Cmdlets, Scripts or Modules provider to create a proper HTML file.

        .PARAMETER Module
        Module Name or Path to the PSM1 file.

        .PARAMETER Command
        Names of the PowerShell Cmdlets or Functions to create HTML Help for.

        .PARAMETER Script
        Path to the script to create HTML Help for.

        .PARAMETER RelatedHelp
        Allows you to add links to outside folders ou URL in the HTML output. Parameter value needs to be an ArrayList of Hastables composed as follow:
            @{
                Title = 'Title'
                UriOrFilePath = 'UriOrFilePath'
            }

        .PARAMETER OutputFolder
        Specify the path to write the HTML file in.

        .EXAMPLE
        New-HtmlHelp -Module ActiveDirectory

        Will create a local HTML file containing all the Cmdlet in ActiveDirectory module and their Comment Based Help content.

        .EXAMPLE
        New-HtmlHelp -Module ActiveDirectory -RelatedHelp @(
            @{
                Title = 'Title'
                UriOrFilePath = 'UriOrFilePath'
            }
        )

        Will create a local HTML file containing all the Cmdlet in ActiveDirectory module and their Comment Based Help content, with a link with the specified title pointing to the specified path ou Uri.

        .EXAMPLE
        New-HtmlHelp -Command 'New-Item','Remove-Item'

        Will create a local HTML file containing Comment Based Help of Cmdlets New-Item and Remove-Item.

        .EXAMPLE
        New-HtmlHelp -Command 'New-Item','Remove-Item' -RelatedHelp @(
            @{
                Title = 'Title'
                UriOrFilePath = 'UriOrFilePath'
            }
        )

        Will create a local HTML file containing Comment Based Help of Cmdlets New-Item and Remove-Item, with a link with the specified title pointing to the specified path ou Uri.

        .EXAMPLE
        New-HtmlHelp -Script .\test.ps1

        Will create a local HTML file containing Comment Based Help of the PowerShell Script 'test.ps1'.

        .EXAMPLE
        New-HtmlHelp -Script .\test.ps1 -RelatedHelp @(
            @{
                Title = 'Title'
                UriOrFilePath = 'UriOrFilePath'
            }
        )

        Will create a local HTML file containing Comment Based Help of the PowerShell Script 'test.ps1', with a link with the specified title pointing to the specified path ou Uri.

        .NOTES
        Author: Thomas Prud'homme (Blog: https://blog.prudhomme.wtf Tw: @Prudhomme_WTF).

        .LINK
        https://github.com/PrudhommeWTF/PSHelp_to_Html

        .INPUTS
        System.String

        .OUTPUTS
        System.IO.File
    #>
    [CmdletBinding(
        DefaultParameterSetName = 'Module'
    )]
    [OutputType([IO.FileInfo[]])]
    Param(
        [Parameter(
            Mandatory         = $true,
            ValueFromPipeline = $true,
            HelpMessage       = 'Module name or path the PSM1 file of the module to write HTML help for.',
            ParameterSetName  = 'Module'
        )]
        [String]$Module,

        [Parameter(
            Mandatory = $true,
            HelpMessage = 'Command to write HTML help for.',
            ParameterSetName = 'Command',
            ValueFromPipeline = $true
        )]
        [String[]]$Command,

        [Parameter(
            Mandatory = $true,
            HelpMessage = 'Path to the script to write HTML help for.',
            ParameterSetName = 'Script'
        )]
        [ValidateScript({Test-Path -Path $_})]
        [String]$Script,

        [Collections.ArrayList]$RelatedHelp,

        [String]$OutputFolder = '.'
    )
    Begin {
        Function Convert-HelpToHtml {
            Param(
                [Parameter(
                    Mandatory = $true,
                    HelpMessage = 'Help object to convert to HTML.',
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName = $true
                )]
                [Object]$HelpObject
            )
            Begin {
                Function New-HelpParameterToHtml {
                    Param(
                        [Parameter(
                            Mandatory = $true,
                            HelpMessage = 'Help Parameter to convert to HTML.',
                            ValueFromPipeline = $true
                        )]
                        [Object]$Parameters
                    )
                    Process {
                        @"
<h5>-$($_.name)</h5>
<div class="parameterInfo">
    <p>$($_.description.Text)</p>
    <table class="table">
        <tr>
            <th scope="row">Type</th>
            <td>$($_.type.name)</td>
        </tr>
        <tr>
            <th scope="row">Position</td>
            <td>$($_.position)</th>
        </tr>
        <tr>
            <th scope="row">Default value</th>
            <td>$($_.default)</td>
        </tr>
        <tr>
            <th scope="row">Accept pipeline input</th>
            <td>$($_.pipelineInput)</td>
        </tr>
        <tr>
            <th scope="row">Accept wild card characters</th>
            <td>$($_.globbing)</td>
        </tr>
    </table>
</div>
"@
                    }
                }
                Function New-HelpExampleToHtml {
                    Param(
                        [Parameter(
                            Mandatory = $true,
                            HelpMessage = 'Help Examples to convert to HTML.',
                            ValueFromPipeline = $true
                        )]
                        [Object]$Example
                    )
                    Process {
                        @"
<strong>$($_.title)</strong>
<div class="card">
    <div class="card-header">
        PowerShell
    </div>
    <div class="card-body">
        <pre class="card-text">
$($_.introduction.Text) $($_.code)
        </pre>
    </div>
</div>
<p>$($_.remarks.text)</p>
"@
                    }
                }
                Function New-HelpSyntaxToHtml {
                    Param(
                        [Parameter(
                            Mandatory = $true,
                            HelpMessage = 'Help Syntax to convert to HTML.',
                            ValueFromPipeline = $true
                        )]
                        [Object]$Syntax
                    )
                    Process {
                        $Output = @"
<strong>$($_.title)</strong>
<div class="card">
    <div class="card-header">
        PowerShell
    </div>
    <div class="card-body">
        <pre class="card-text">
$($_.name)
"@
                        $_.parameter | ForEach-Object -Process {
                            if ($_.required -eq $true) {
                                if ($_.position -ne $null) {
                                    $Calc = '[-'
                                } else {
                                    $Calc = '-'
                                }

                                #Parameter Name
                                $Calc += $_.name

                                if ($_.position -ne $null) {
                                    $Calc += ']'
                                }

                                #Type
                                if ($_.parameterValue) {
                                    $Calc += " &lt;$($_.parameterValue)&gt;"
                                }
                            } else {
                                if ($_.position -ne $null) {
                                    $Calc = '[[-'
                                } else {
                                    $Calc = '[-'
                                }

                                #Parameter Name
                                $Calc += $_.name

                                if ($_.position -ne $null) {
                                    $Calc += "]$(if ($_.parameterValue -ne '') {" &lt;$($_.parameterValue)&gt;"})]"
                                } else {
                                    $Calc += ']'
                                }
                            }
                            $Output += @"

    $Calc
"@
                        }
                        $Output += @"
        </pre>
    </div>
</div>
<p>$($_.remarks.text)</p>
"@

                        Write-Output -InputObject $Output
                    }
                }
            }
            Process {    
                #Menu
                $ContentId = $HelpObject.Name -replace '\*|\W',''

                if (Test-Path -Path $HelpObject.Name) {
                    $Name = Get-ItemProperty -Path $HelpObject.Name | Select-Object -ExpandProperty Name
                } else {
                    $Name = $HelpObject.Name
                }

                $MenuEntry = "<li><button type=`"button`" class=`"btn btn-link`" OnClick=`"displayOnly('#$ContentId')`">$($Name)</button></li>"

                #Help Required Parameters
                $RequiredParametersOutput = $HelpObject.parameters.parameter | Where-Object -FilterScript {$_.required -eq $true} | New-HelpParameterToHtml

                #Help Optional Parameters
                $OptionalParametersOutput = $HelpObject.parameters.parameter | Where-Object -FilterScript {$_.required -eq $false} | New-HelpParameterToHtml

                #Help Examples
                if ($HelpObject.examples -ne $null) {
                    $ExampleOutput = $HelpObject.examples.example | New-HelpExampleToHtml
                } else {
                    $ExampleOutput = $null
                }

                #Help Syntax
                if ($HelpObject.syntax -ne $null) {
                    $SyntaxOutput = $HelpObject.syntax.syntaxItem | New-HelpSyntaxToHtml
                } else {
                    $SyntaxOutput = $null
                }

                #Content
                $HtmlEntry = @"
<div id="$($ContentId)" class="row">
    <div class="col-md-9">
        <h1>$Name</h1>
        <p>$($HelpObject.Synopsis.text)</p>
        $(if ($SyntaxOutput) {
            @"
        <h2 id="Syntax_$ContentId">Syntax</h2>
        $SyntaxOutput
"@
        })
        $(if ($HelpObject.Description.text) {
            @"
        <h2 id="Description_$ContentId">Description</h2>
        <p>$($HelpObject.Description.text)</p>
"@
        })
        $(if ($ExampleOutput) {
            @"
        <h2 id="Examples_$ContentId">Examples</h2>
        $ExampleOutput
"@
        })
        $(if ($RequiredParametersOutput) {
            @"
        <h2 id="RequiredParameters_$ContentId">Required Parameters</h2>
        $RequiredParametersOutput
"@
        })
        $(if ($OptionalParametersOutput) {
            @"
        <h2 id="OptionalParameters_$ContentId">Optional Parameters</h2>
        $OptionalParametersOutput
"@
        })
        $(if ($_.inputTypes) {
            @"
        <h2 id="Inputs_$ContentId">Inputs</h2>
        <p>$($_.inputTypes)</p>
"@
        })
        $(if ($_.returnValues) {
            @"
        <h2 id="Outputs_$ContentId">Outputs</h2>
        <p>$($_.returnValues)</p>
"@
        })
        $(if ($_.returnValues) {
            @"
        <h2 id="RelatedLinks_$ContentId">Related Links</h2>
        <p>$($_.relatedLinks)</p>
"@
        })
    </div>
    <div class="col-md-3">
        <h5 class="text-secondary">Quick links</h5>
        <ul class="list-unstyled right-menu">
            <li><a href="#Syntax_$ContentId" class="btn btn-link$(if (!$SyntaxOutput) {' disabled'})">Syntax</a></li>
            <li><a href="#Description_$ContentId" class="btn btn-link$(if (!$HelpObject.Description.text) {' class="disabled'})">Description</a></li>
            <li><a href="#Examples_$ContentId" class="btn btn-link$(if (!$ExampleOutput) {' disabled'})">Examples</a></li>
            <li><a href="#RequiredParameters_$ContentId" class="btn btn-link$(if (!$RequiredParametersOutput) {' disabled'})">Required Parameters</a></li>
            <li><a href="#OptionalParameters_$ContentId" class="btn btn-link$(if (!$OptionalParametersOutput) {' disabled'})">Optional Parameters</a></li>
            <li><a href="#Inputs_$ContentId" class="btn btn-link$(if (!$_.inputTypes) {' disabled'})">Inputs</a></li>
            <li><a href="#Outputs_$ContentId" class="btn btn-link$(if (!$_.returnValues) {' disabled'})">Outputs</a></li>
            <li><a href="#RelatedLinks_$ContentId" class="btn btn-link$(if (!$_.relatedLinks) {' disabled'})">Related Links</a></li>
        </ul>
    </div>
</div>
"@

                #Output
                New-Object -TypeName PSObject -Property @{
                    Menu    = $MenuEntry
                    Content = $HtmlEntry
                }
            }
        }
        Function New-ModuleHelpToHtml {
            Param(
                [Parameter(
                    Mandatory = $true,
                    HelpMessage = 'Module name to convert Help to HTML.'
                )]
                [String]$ModuleName,

                [Collections.ArrayList]$RelatedHelpFiles
            )
            $Commands = Get-Command -Module $ModuleName

            #Start with Content
            $Commands | ForEach-Object -Begin {
                $LeftMenu = @'
<div class="col-3">
    <h5 class="text-secondary">Available Cmdlets</h5>
    <ul class="list-unstyled left-menu">
'@
                $MainContent = @'
<div class="col-9" id="content">
'@
            } -Process {
                $HtmlEntry = Convert-HelpToHtml -HelpObject (Get-Help -Name $_.Name -Full)
                $LeftMenu    += $HtmlEntry.Menu
                $MainContent += $HtmlEntry.Content

            } -End {
                #Closing Menu & MainContent Valriables
                if ($RelatedHelpFiles) {
                    $RelatedHelpFiles | ForEach-Object -Begin {
                        $LeftMenu += @'
    </ul>
    <h5 class="text-secondary">Related helps</h5>
    <ul class="list-unstyled left-menu">
'@
                    } -Process {
                        $LeftMenu += @"
        <li><a href="$($_.UriOrFilePath)">$($_.Title)</a></li>
"@
                    }
                }
                $LeftMenu += @'
    </ul>
</div>
'@
                $MainContent += @'
</div>
'@
            }

            $LeftMenu, $MainContent
        }
        Function New-CommandHelpToHtml {
            [CmdletBinding()]
            Param(
                [Parameter(
                    Mandatory = $true,
                    HelpMessage = 'Command to convet Help to HTML for.',
                    ValueFromPipeline = $true
                )]
                [String]$Command,

                [Collections.ArrayList]$RelatedHelpFiles
            )
            Begin {
                #Open Menu & MainContent Valriables
                $LeftMenu = @'
<div class="col-3">
    <h5 class="text-secondary">Available Cmdlets</h5>
    <ul class="list-unstyled left-menu">
'@
                $MainContent = @'
<div class="col-9" id="content">
'@
            }
            Process {
                $HtmlEntry = Convert-HelpToHtml -HelpObject (Get-Help -Name $Command -Full)
                $LeftMenu    += $HtmlEntry.Menu
                $MainContent += $HtmlEntry.Content
            }
            End {
                #Closing Menu & MainContent Valriables
                if ($RelatedHelpFiles) {
                    $RelatedHelpFiles | ForEach-Object -Begin {
                        $LeftMenu += @'
    </ul>
    <h5 class="text-secondary">Related helps</h5>
    <ul class="list-unstyled left-menu">
'@
                    } -Process {
                        $LeftMenu += @"
        <li><a href="$($_.UriOrFilePath)">$($_.Title)</a></li>
"@
                    }
                }
                $LeftMenu += @'
    </ul>
</div>
'@
                $MainContent += @'
</div>
'@

                #Return Variables
                $LeftMenu
                $MainContent
            }
        }
    }
    Process {
        if ($PSCmdlet.ParameterSetName -eq 'Module') {
            if (Test-Path -Path $Module -ErrorAction SilentlyContinue) {
                $ModuleInfo = Get-ItemProperty -Path $Module
                try {
                    Import-Module -Name $ModuleInfo.FullName
                    Write-Verbose -Message "Imported module $($ModuleInfo.FullName)"
                }
                catch {
                    Write-Error -Message $_
                }
                $Module = $ModuleInfo.BaseName
            } else {
                if (Get-Module -Name $Module -ListAvailable) {
                    try {
                        Import-Module -Name $Module
                        Write-Verbose -Message "Imported module $($ModuleInfo.FullName)"
                    }
                    catch {
                        Write-Error -Message $_
                    }
                } else {
                    Throw 'Module not available on this computer.'
                }
            }

            $HtmlTitle   = "$Module - PowerShell Help"
            $FileName    = "$OutputFolder\$Module.html"
            $MainContent = New-ModuleHelpToHtml -ModuleName $Module -RelatedHelpFiles $RelatedHelpFiles

            Remove-Module -Name $ModuleInfo.BaseName
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Command') {
            $HtmlTitle   = if (($Command | Measure-Object).count -gt 1) {
                'PowerShell Help'
            } else {
                "$Command - PowerShell Help"
            }
            $FileName    = if (($Command | Measure-Object).count -gt 1) {
                "$OutputFolder\HelpToHtml_$((Get-Date).ToString('yyyyMMdd')).html"
            } else {
                "$OutputFolder\$Command.html"
            }
            $MainContent = $Command | New-CommandHelpToHtml -RelatedHelpFiles $RelatedHelpFiles
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Script') {
            $ScriptProperty = Get-ItemProperty -Path $Script
            $HtmlTitle      = "$($ScriptProperty.Name) - PowerShell Help"
            $FileName       = "$OutputFolder\$($ScriptProperty.BaseName)-Help.html"
            $MainContent    = New-CommandHelpToHtml -Command $Script -RelatedHelpFiles $RelatedHelpFiles
        }

        $Header = @"
<!doctype html>
<html lang="en">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

        <!-- Bootstrap CSS -->
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO" crossorigin="anonymous">

        <!-- Font Awesome -->
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.6.3/css/all.css" integrity="sha384-UHRtZLI+pbxtHCWp1t77Bi1L4ZtiqrqD80Kn4Z8NTSRyMA2Fd33n5dQ8lWUE00s/" crossorigin="anonymous">

        <style>
.left-menu {
    border-right: 1px solid #eee;
}
.right-menu {
    border-left: 1px solid #eee;
}
#content {
    margin-top: 10px;
}
#myBtn {
  display: none;
  position: fixed;
  bottom: 20px;
  right: 30px;
  z-index: 99;
  border: none;
  outline: none;
  padding: 15px;
}
        </style>

        <script type="text/javascript">
function displayOnly(name) {
	`$('#content').children().hide();
	`$(name).show();
}

// When the user scrolls down 20px from the top of the document, show the button
window.onscroll = function() {scrollFunction()};

function scrollFunction() {
  if (document.body.scrollTop > 20 || document.documentElement.scrollTop > 20) {
    document.getElementById("myBtn").style.display = "block";
  } else {
    document.getElementById("myBtn").style.display = "none";
  }
}

// When the user clicks on the button, scroll to the top of the document
function topFunction() {
  document.body.scrollTop = 0; // For Safari
  document.documentElement.scrollTop = 0; // For Chrome, Firefox, IE and Opera
}
        </script>

        <title>$HtmlTitle</title>
    </head>
    <body OnLoad="displayOnly('#')">
        <div>
            <nav class="navbar navbar-expand-lg navbar-light bg-light">
                <span class="navbar-brand mb-0 h1">$HtmlTitle</span>
                <div class="collapse navbar-collapse"></div>
                <span class="navbar-text">
                    <div class="btn-group dropleft">
                        <button type="button" class="btn btn-danger dropdown-toggle" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                            Informations
                        </button>
                        <div class="dropdown-menu">
                            <a class="dropdown-item"><strong>Generated on:</strong> $(Get-Date)</a>
                            <div class="dropdown-divider"></div>
                            <a class="dropdown-item" href="https://github.com/PrudhommeWTF/PSHelp_to_Html" target="_blank"><i class="fab fa-github-square"></i> PSHelp_to_Html on GitHub</a>
                        </div>
                    </div>
                </span>
            </nav>
            <div class="container-fluid">
                <div class="row">
"@
        $Footer = @'
                </div>
            </div>
        </div>
        <button id="myBtn" class="btn btn-light" onclick="topFunction()"><i class="fas fa-chevron-up"></i> Back to top</button>

        <!-- Optional JavaScript -->
        <!-- jQuery first, then Popper.js, then Bootstrap JS -->
        <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js" integrity="sha384-ZMP7rVo3mIykV+2+9J3UJ46jBk0WLaUAdn689aCwoqbBJiSnjAK/l8WvCWPIPm49" crossorigin="anonymous"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js" integrity="sha384-ChfqqxuZUCnJSK3+MXmPNIyE6ZbWh2IMqE241rYiqJxyMiZ6OW/JmZQ5stwEULTy" crossorigin="anonymous"></script>

    </body>
</html>
'@
    
        $Header, $MainContent, $Footer |  Out-File -FilePath $FileName
    }
}