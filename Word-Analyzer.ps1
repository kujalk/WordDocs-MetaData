<#
Purpose - This script is used to get author, modified time, word count .. from Word files
Execution - ./Word-Analyzer.ps1 -Folder C:\Sample
Developer - K.Janarthanan
Date - 20/11/2020
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String]
    $Folder
)

try 
{
    if(-not(Test-Path -Path $Folder))
    {
        throw "Folder is not found"
    }

    $FinalFolder = "$Folder\*"
    $WordFiles = Get-ChildItem -Path $FinalFolder -Include *doc,*docx -Recurse -Force

    #Properties 
    $Properties = @("Author","Last author","Creation date","Last save time","Number of pages",
    "Number of words","Number of characters","Number of characters (with spaces)")

    if($WordFiles.Count -gt 0)
    {
        Write-Host "Going to process files" -ForegroundColor Green

        foreach($SingleFile in $WordFiles)
        {
            $CSVObject = New-Object -Type PSObject

            Write-Host "`nWorking on File :  $($SingleFile.Name)" -ForegroundColor Green
            try 
            {
                $CSVObject | Add-Member -MemberType NoteProperty -Name "File Name" -Value $SingleFile.Name
                $CSVObject | Add-Member -MemberType NoteProperty -Name "File Path" -Value $SingleFile.DirectoryName
                
                $Application = New-Object -ComObject word.application
                $Application.Visible = $false

                $Document = $Application.documents.open($SingleFile.fullname,$false,$true)
                $Document.Repaginate()
                $Binding = "System.Reflection.BindingFlags" -as [type]

                Foreach($Property in $Document.BuiltInDocumentProperties) 
                {
                    try 
                    {
                        $Key= [System.__ComObject].invokemember("name",$Binding::GetProperty,$null,$property,$null)
                        $Val = [System.__ComObject].invokemember("value",$Binding::GetProperty,$null,$property,$null)
                        
                        if($Key -in $Properties)
                        {
                            $CSVObject | Add-Member -MemberType NoteProperty -Name $Key -Value $Val
                        }
                    }
                    catch
                    {
                        #These errors will be ignored
                        if($Key -in $Properties)
                        {
                            $CSVObject | Add-Member -MemberType NoteProperty -Name $Key -Value "Not Available"
                        }
                    }
                }

                $CSVObject | Add-Member -MemberType NoteProperty -Name "No of Images" -Value $document.inlineshapes.count

                #$application.documents.close($false) 
                $Document.Close([ref] [Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)  
                
                $CSVObject | Export-CSV -NoTypeInformation -Append -Path "./Details_WordDocs.csv"
            }
            catch 
            {
                Write-Host "Error occured while processing file $($SingleFile.FullName) - $_"  -ForegroundColor Yellow  
            }
        }
    }
    else 
    {
        Write-Host "No any files found with .doc or .docx format" -ForegroundColor Green
    }
}
catch 
{
    Write-Host "Error occured : $_" -ForegroundColor Red
}