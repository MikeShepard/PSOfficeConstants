add-type -path C:\temp\HtmlAgilityPack.dll

function ParseWebValue {
    Param([string]$value)
    $val1 = $value.Split(' ')[0]
    if (($out = $val1 -as [int]) -is [int]) {
        [int]$out 
    } elseif ($val1.StartsWith('&amp;')) {
        if (($out = $val1.Replace('&amp;', '0x') -as [int]) -is [int]) {
            $val1.Replace('&amp;', '0x')
        } elseif (($out = $val1.Replace('&amp;H', '0x') -as [int]) -is [int]) {
            $val1.Replace('&amp;H', '0x')
        }
    } 
}
$products = @{
    'Excel' = 'xl'
    'Word'       = 'wd'
    'Visio'      = 'vis'
    'Outlook'    = 'ol'
    'Publisher'  = 'pb'
    'PowerPoint' = 'pp'
    'Access'     = 'ac'
    'Project'    = 'pj'
}
$exceptions = 0
$docsToSkip = @('word.xlpieslicelocation', 'outlook.olxdefaultfolde', 'powerpoint.xlpieslicelocation', 'word.xlpieslicelocation')
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
foreach ($product in $products.Keys) {
    write-host "Generating script for $product"

    $url = 'https://docs.microsoft.com/en-us/office/vba/api/{0}(enumerations)' -f $product

    if ($product -in 'Outlook', 'Access') {
        #Access and Outlook have a slightly different base URL
        $url = 'https://docs.microsoft.com/en-us/office/vba/api/{0}' -f $product
    }
    if ($product -eq 'Access') {
        #and the page layout is different in Access, too.
        $xpath = '//tbody/tr/td[1]/a'
    } else {
        $xpath = '/html[1]/body[1]/div[2]/div[1]/section[1]/div[1]/div[2]/main[1]/ul[2]/li/a'
    }
    $doc = New-Object HtmlAgilityPack.HtmlWeb
    $result = $doc.Load($url) 
    $links = $result.DocumentNode.SelectNodes($xpath)
    $productStringBuilder = new-object System.Text.StringBuilder -ArgumentList 10000
    [void]$productStringBuilder.AppendLine("#constants for $product based on $url")
    [void]$productStringBuilder.Append('$')
    [void]$productStringBuilder.Append($products[$product])
    [void]$productStringBuilder.AppendLine('=[Ordered]@{')
    $linksDone = 0
    foreach ($link in $links) {
        $pct = $linksDone * 100.0 / $links.Count
        write-progress -Activity "Enums for $product" -status $link.Attributes['href'].Value -PercentComplete $pct
        $docName = $link.Attributes['href'].Value
        if ($docName -in $docsToSkip) { break }
        $innerURL = 'https://docs.microsoft.com/en-us/office/vba/api/{0}' -f $docname

        [void]$productStringBuilder.AppendLine()
        [void]$productStringBuilder.AppendLine("#values from $docname")
        [void]$productStringBuilder.AppendLine("#****************************")
        $innerdoc = New-Object HtmlAgilityPack.HtmlWeb
        $innerresult = $innerdoc.Load($innerurl) 
        $innerRows = $innerresult.DocumentNode.SelectNodes('//table[1]/tbody/tr/td')
        $count = 0
        foreach ($Row in $innerRows) {
            #write-verbose "$count- $($row.Innertext)"
            if ($row.InnerText.Replace(' ', '') -eq '') { break}
            $column = $row.Xpath.split('[')[-1].Replace(']', '')
            if ([int]$column -gt 3) {
                break
            }
            switch ($column) {
                1 {
                    if ($count -gt 0) {
                        [void]$productStringBuilder.AppendLine()
                    }
                    $count += 1
                    $name = $row.InnerText.split(' ')[0]; break
                }
                2 {
                    $value = ParseWebValue($row.innertext)
                    if (-not($null -eq $Value )) {
                        if ("$docname.$name" -ne 'excel.constants.xlManual') {
                            [void]$productStringBuilder.Append($name)
                            [void]$productStringBuilder.Append("`t=`t")
                            [void]$productStringBuilder.Append($Value)
                        }
                    } else {
                        write-host "Skipping $docName - $name because it isn't an INT ($($row.InnerText))"
                        $exceptions += 1
                    }
                    break
                }
                3 {
                    if ("$docname.$name" -ne 'excel.constants.xlManual') {
                        [void]$productStringBuilder.Append("`t`t# $($row.InnerText)"); break
                    }
                }

            }

        } #rows
        [void]$productStringBuilder.AppendLine()
        $linksDone += 1

    } #
    [void]$productStringBuilder.AppendLine('}')
    [void]$productStringBuilder.AppendLine('#End Enum')
    [void]$productStringBuilder.AppendLine('')
    [void]$productStringBuilder.AppendLine('${0}=new-object PSCustomObject -Property ${0}' -f $products[$product])

    Set-Content -path "$product.ps1" -value $productStringBuilder.ToString() -Force
    $productStringBuilder = $null

}#product
write-progress -Activity "Enums for $product" -Completed
write-host "$exceptions exceptions found"


<#
$xl = New-Object -ComObject visio.Application
$constants = $xl.gettype().assembly.getexportedtypes() | where-object {$_.IsEnum -and $_.name -eq 'constants'}

$pso = new-object psobject
measure-command {
[enum]::getNames($constants) | foreach { if ($_ -and $constants::$__){ $pso | Add-Member -MemberType NoteProperty $_ ($constants::$_) }}
$xlConstants = $pso
}

$hash=@{}
measure-command {
[enum]::getNames($constants) | foreach { if ($_ -and $constants::$__){$hash[$_]= ($constants::$_) }}

}
#>

