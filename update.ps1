Function Auto_Image {
    $out = "D:\test"
    $images="$env:USERPROFILE\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets"
    $DstPic = "D:\Time\AutoPic.jpg"

    New-Item -ItemType Directory -Path $out -Force

    If (!(Test-Path $images)) {
        # if the location doesn't exist, Spotlight might not have been enabled
        Write-Host "Spotlight image cache not found - is Spotlight enabled?"
        Write-Host "Check manual instructions at https://www.howtogeek.com/247643/how-to-save-windows-10s-lock-screen-spotlight-images-to-your-hard-drive/"
        Exit -2
    }
    
    Function output_image {
        Param(
            [Parameter(Mandatory = $True)]
            [string]
            $out,
            [Parameter(Mandatory = $True)]
            [System.Object]
            $imageFile
        )
        # Construct the new image file name and copy the file to the output location, adding ".jpg" to the end
        $newname = $imageFile.BaseName + ".jpg"
        $dest = Join-Path -Path $out -ChildPath $newname
        Copy-Item -Path $imageFile.FullName -Destination $dest -Force
    }
    
    Get-ChildItem $images | Foreach-Object {
        Try {
            $image = New-Object -ComObject Wia.ImageFile
            $image.LoadFile($_.FullName)
    
            # Spotlight images are 1980 x 1080 for desktops, 1080 x 1980 for phones
            If($image.Height -eq 1080) {
                # aspect ratio of desktops make images 1080 in height
                output_image -out $out -imageFile $_
            }
        } Catch [System.ArgumentException] {
            # Probably wasn't an image, so skip.
        }
    }
    Get-ChildItem -Path $out | Sort-Object LastAccessTime -Descending | Select-Object -First 1 | Copy-Item -Destination $DstPic
    Remove-Item -Path $out -Confirm $true
}

Function Auto_Omer {
    if (($ws.Cells.Item(15 , 10).text) -ne "0") {$pp.ActivePresentation.Slides(4).SlideShowTransition.Hidden = $false}
    else {$pp.ActivePresentation.Slides(4).SlideShowTransition.Hidden = $true}
}

# update zal excel



#open excel
$TimePath = "D:\Time"
$x1 = New-Object -ComObject "Excel.Application"
$x1.Visible = $true


$ZalFile = "$TimePath\ex\Zal.xlsx"
$ZalWB = $x1.workbooks.Open($ZalFile)

$EXfile = "$TimePath\ex\time2.xlsm"
$wb = $x1.workbooks.Open($EXfile)
#update internet links
if (Test-Connection -ComputerName 8.8.8.8) {
    $wb.refreshall()
    Start-Sleep -Seconds 5
}
$ws = $wb.Sheets.Item(5)

#check if its holiday today by checking the I4 cell in excel
if (($ws.Cells.Item(4 , 9).text) -ne "0") {
    $DayTime = $ws.Cells.Item(4 , 9).text + ".pptm"
} elseif ((Get-Date).DayOfWeek -eq "Saturday") {
    $DayTime = "shabat.pptm"
} elseif ((Get-Date).DayOfWeek -eq "Friday") {
    $DayTime = "shabat.pptm"
} else {
    $DayTime = "hol.pptm"
    Auto_Image
}


#open powerpoint
$ppFile = "$TimePath\pp\$DayTime"

$pp = New-Object -ComObject "powerpoint.application"

#closing current presentation
# $pp.ActivePresentation.Save()
# $pp.ActivePresentation.Close()

#opening the new presentation
$prp = $pp.Presentations.Open($ppFile)
$prp.UpdateLinks()

#set the background for hol.pptm
if ($DayTime -eq "hol.pptm") {
    $osld = $pp.ActivePresentation.Slides(1)
    $osld.FollowMasterBackground = $false
    $osld.Background.Fill.UserPicture($DstPic)
}

# decide if today it's the Omer time and show the hide or unhide the slide of it
if (($ws.Cells.Item(4 , 9).text) -eq "0") {
    Auto_Omer
}

# $prp.Save()
$pp.ActivePresentation.Save()
$pp.ActivePresentation.Close()
$pp.Quit()
Stop-Process -Name POWERPNT

Start-Process powerpnt -ArgumentList ("/s " + "$ppFile")

$wb.Save()
$wb.Close()
$ZalWB.Save()
$ZalWB.Close()


$x1.Quit()