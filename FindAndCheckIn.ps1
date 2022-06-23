#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
  
#PowerShell to Bulk check-in all documents
Function CheckIn-AllDocuments([String]$SiteURL)
{
    Try{
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Credentials
 
        #Get the Web
        $Web = $Ctx.Web
        $Ctx.Load($Web)
        $Ctx.Load($Web.Webs)
        $Ctx.ExecuteQuery()
 
        #Get All Lists from the web
        $Lists = $Web.Lists
        $Ctx.Load($Lists)
        $Ctx.ExecuteQuery()
  
        #Prepare the CAML query
        $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
        $Query.ViewXml = "@
        <View Scope='RecursiveAll'>
            <Query>
                <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
            </Query>
            <RowLimit Paged='TRUE'>2000</RowLimit>
        </View>"
 
        #Array to hold Checked out files
        $CheckedOutFiles = @()
        Write-host -f Yellow "Processing Web:"$Web.Url
         
        #Iterate through each document library on the web
        ForEach($List in ($Lists | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $False -and $_.ItemCount -gt 0}) )
        {
            Write-host -f Yellow "`t Processing Document Library:"$List.Title
 
                $Counter=1
                #Batch Process List items
                Do {
                    $ListItems = $List.GetItems($Query)
                    $Ctx.Load($ListItems)
                    $Ctx.ExecuteQuery()
 
                    $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition
 
                    #Get All Checked out files
                    ForEach($Item in $ListItems | Where {$_.FileSystemObjectType -eq "File"})
                    {
                        #Display a Progress bar
                        Write-Progress -Activity "Scanning Files in the Library" -Status "Testing if the file is Checked-Out '$($Item.FieldValues.FileRef)' ($Counter of $($List.ItemCount))" -PercentComplete (($Counter / $List.ItemCount) * 100)
 
                        Try{
                            #Get the Checked out File data
                            $File = $Web.GetFileByServerRelativeUrl($Item["FileRef"])
                            $Ctx.Load($File)
                            $CheckedOutByUser = $File.CheckedOutByUser
                            $Ctx.Load($CheckedOutByUser)
                            $Ctx.ExecuteQuery()
 
                            If($File.Level -eq "Checkout")
                            {
                                Write-Host -f Green "`t`t Found a Checked out File '$($File.Name)' at $($Item['FileRef']), Checked Out By: $($CheckedOutByUser.LoginName)"
 
                                #Check in the document
                                $File.CheckIn("Checked-in By Administrator through PowerShell!", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)     
                                $Ctx.ExecuteQuery()
                                Write-Host -f Green "`t`t File '$($File.Name)' Checked-In Successfully!"
                            }
                        }
                        Catch {
                            write-host -f Red "Error Check In: $($Item['FileRef'])" $_.Exception.Message
                        }
                        $Counter++
                    }
                }While($Query.ListItemCollectionPosition -ne $Null)
        }
 
        #Iterate through each subsite of the current web and call the function recursively
        ForEach($Subweb in $Web.Webs)
        {
            #Call the function recursively to process all subsites underneath the current web
            CheckIn-AllDocuments -SiteURL $Subweb.URL
        }
    }
    Catch {
        write-host -f Red "Error Bulk Check In Files!" $_.Exception.Message
    }
}
 
#Config Parameters
#Edit for each site collection url to reiterate through subsites
$SiteURL="https://sitecollectionurl.sharepoint.com/sites/"
  
#Setup Credentials to connect
$Cred= Get-Credential
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
  
#Call the function: sharepoint online powershell to check in all documents in a Site Collection
CheckIn-AllDocuments -SiteURL $SiteURL
