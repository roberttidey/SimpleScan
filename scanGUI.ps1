# Simple scan
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$deviceManager = new-object -ComObject WIA.DeviceManager

$wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"

$outPath = $Env:USERPROFILE + "\Downloads\"
$outImageNames = @()

#Get all scanners
$infos = $deviceManager.DeviceInfos
$wiaScannerType = 1
$ScannerIx = 0
$ScannerNames = @()
$ScannerNumbers = @()
foreach($info in $infos) {
	if($info.Type -eq $wiaScannerType) {
		$ScannerIx = $ScannerIx + 1
		foreach($p in $info.Properties) {
			if($p.Name -eq 'Name') {
				$ScannerNumbers += $ScannerIx
				$ScannerNames += $p.Value
			}
		}
	}
}

function ConvertToPdf {
	param ($inFiles)
	$outFile = $outPath + $inFiles[0].Substring(0, $inFiles[0].LastIndexOf(".")) + ".pdf"
    Add-Type -AssemblyName System.Drawing

    try {
        $doc = [System.Drawing.Printing.PrintDocument]::new()
        $opt = $doc.PrinterSettings = [System.Drawing.Printing.PrinterSettings]::new()
        $opt.PrinterName = "Microsoft Print to PDF"
        $opt.PrintToFile = $true
        $opt.PrintFileName = $outFile

        $script:_pageIndex = 0
        $doc.add_PrintPage({
            param([object]$Sender, [System.Drawing.Printing.PrintPageEventArgs] $a)
            $file = $outPath + $inFiles[$script:_pageIndex]

            $script:_pageIndex = $script:_pageIndex + 1
            try {
                $image = [System.Drawing.Image]::FromFile($file)
                $a.Graphics.DrawImage($image, $a.PageBounds)
                $a.HasMorePages = $script:_pageIndex -lt $inFiles.Count
            }
            finally {
                $image.Dispose()
            }
        })

        $doc.PrintController = [System.Drawing.Printing.StandardPrintController]::new()
        $doc.Print()
	}
    finally {
        if ($doc) { $doc.Dispose() }
    }
	return [System.IO.File]::Exists($outFile)

}

#Set up GUI controls
#Scanner select
	$ScannerLabel = New-Object Windows.Forms.Label
	$ScannerLabel.Text = "Scanner"
	$ScannerLabel.AutoSize = $true
	$ScannerLabel.Location = New-Object Drawing.Point(20,10)
	$ScannerLabel.ForeColor = [System.Drawing.Color]::Black

	$ScannerComboBox = New-Object system.Windows.Forms.ComboBox
	$ScannerComboBox.text = ""
	$ScannerComboBox.width = 140
	$ScannerComboBox.autosize = $true
	$ScannerComboBox.location = New-Object System.Drawing.Point(20,30)
	# Add the items in the dropdown list
	$ScannerNames | ForEach-Object {[void] $ScannerComboBox.Items.Add($_)}
	# Select the default value
	$ScannerComboBox.SelectedIndex = 0

#Format select
	$FormatLabel = New-Object Windows.Forms.Label
	$FormatLabel.Text = "Image Format"
	$FormatLabel.AutoSize = $true
	$FormatLabel.Location = New-Object Drawing.Point(170,10)
	$FormatLabel.ForeColor = [System.Drawing.Color]::Black

	$FormatComboBox = New-Object system.Windows.Forms.ComboBox
	$FormatComboBox.text = ""
	$FormatComboBox.width = 60
	$FormatComboBox.autosize = $true
	$FormatComboBox.location = New-Object System.Drawing.Point(170,30)
	# Add the items in the dropdown list
	@('JPG','PNG','PDF') | ForEach-Object {[void] $FormatComboBox.Items.Add($_)}
	# Select the default value
	$FormatComboBox.SelectedIndex = 1

#Mode select
	$ModeLabel = New-Object Windows.Forms.Label
	$ModeLabel.Text = "Scan Mode"
	$ModeLabel.AutoSize = $true
	$ModeLabel.Location = New-Object Drawing.Point(20,70)
	$ModeLabel.ForeColor = [System.Drawing.Color]::Black

	$ModeComboBox = New-Object system.Windows.Forms.ComboBox
	$ModeComboBox.text = ""
	$ModeComboBox.width = 100
	$ModeComboBox.autosize = $true
	$ModeComboBox.location = New-Object System.Drawing.Point(20,90)
	# Add the items in the dropdown list
	@('A4 Colour','A4 Grey','A4 Black-White','Letter Colour','Letter Grey','Letter Black-White') | ForEach-Object {[void] $ModeComboBox.Items.Add($_)}
	# Select the default value
	$ModeComboBox.SelectedIndex = 1

#Resolution select
	$ResolutionLabel = New-Object Windows.Forms.Label
	$ResolutionLabel.Text = "Resolution"
	$ResolutionLabel.AutoSize = $true
	$ResolutionLabel.Location = New-Object Drawing.Point(170,70)
	$ResolutionLabel.ForeColor = [System.Drawing.Color]::Black

	$ResolutionComboBox = New-Object system.Windows.Forms.ComboBox
	$ResolutionComboBox.text = ""
	$ResolutionComboBox.width = 60
	$ResolutionComboBox.autosize = $true
	$ResolutionComboBox.location = New-Object System.Drawing.Point(170,90)
	# Add the items in the dropdown list
	@(75,100,150,200,300,400,600,1200) | ForEach-Object {[void] $ResolutionComboBox.Items.Add($_)}
	# Select the default value
	$ResolutionComboBox.SelectedIndex = 4

#Primary scan button and action
	$ScanButton = New-Object System.Windows.Forms.Button
	$ScanButton.Location = New-Object System.Drawing.Size (20,170)
	$ScanButton.Size = New-Object System.Drawing.Size(95,40)
	$ScanButton.Font=New-Object System.Drawing.Font("Lucida Console",18,[System.Drawing.FontStyle]::Regular)
	$ScanButton.BackColor = "LightGray"
	$ScanButton.Text = "Scan"
	$ScanButton.Add_Click({
		#Do the scan
		$SelectedScanner = $ScannerComboBox.SelectedIndex
		$device = $deviceManager.DeviceInfos.Item($ScannerNumbers[$SelectedScanner]).Connect()    
		$item = $device.Items.Item(1)
		
		#Get Mode
		$mode = @(1,2,4,1,2,4)[$ModeComboBox.SelectedIndex]
		$xSize = @(2490,2490,2490,2550,2550,2550)[$ModeComboBox.SelectedIndex]
		$ySize = @(3510,3510,3510,3300,3300,3300)[$ModeComboBox.SelectedIndex]

		#Get Format index
		$FormatIx = $FormatComboBox.SelectedIndex
		if($FormatIx -ne 0) {
			$wiaFormat = $wiaFormatPNG
			$fExt = '.png'
		} else {
			$wiaFormat = $wiaFormatJPEG
			$fExt = '.jpg'
		}
		
		#Get Resolution
		$Resolution=$ResolutionComboBox.Items[$ResolutionComboBox.SelectedIndex]
		$item.properties("6146").Value = $mode
		$item.properties("6147").Value = $resolution
		$item.properties("6148").Value = $resolution
		$item.properties("6151").Value = $xSize
		$item.properties("6152").Value = $ySize
		$image = $item.Transfer($wiaFormat) 
		if($image.FormatID -ne $wiaFormat) {
			$imageProcess = new-object -ComObject WIA.ImageProcess
			$imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
			$imageProcess.Filters.Item(1).Properties.Item("FormatID") = $wiaFormat
			$image = $imageProcess.Apply($image)
		}
		$scandate = (Get-Date).ToString("yyyyMMdd_HHmmss")
		$fname = $scandate + "-scan"
		$fname = $fname + $fExt

		$image.SaveFile($outPath + $fname)
		
		#if pdf then convert image and delete original scan
		if($FormatIx -eq 2) {
			if(ConvertToPdf -inFiles @($fname)) {
				try {
					Remove-Item -path ($outPath + $fname)
				}
				finally {write-Output 'Done'}
			}
		} else {
			$script:outImageNames += $fname
		}
	})

#Convert image collection to pdf button
	$ConvertButton = New-Object System.Windows.Forms.Button
	$ConvertButton.Location = New-Object System.Drawing.Size (130,170)
	$ConvertButton.Size = New-Object System.Drawing.Size(130,40)
	$ConvertButton.Font=New-Object System.Drawing.Font("Lucida Console",16,[System.Drawing.FontStyle]::Regular)
	$ConvertButton.BackColor = "LightGray"
	$ConvertButton.Text = "Make PDF"
	$ConvertButton.Add_Click({
		if($script:outImageNames.Count -gt 0) {
			if(ConvertToPdf -inFiles $script:outImageNames) {
				foreach($outImageName in $script:outImageNames) {
					Remove-Item -path ($outPath + $outImageName)
				}
				$outImageNames = @()
			}
		}
	})
		
#Set up form

	$Form = New-Object Windows.Forms.Form
	$Form.Text = "Simple Scan"
	$Form.Width = 300
	$Form.Height = 270
	$Form.BackColor="LightBlue"

<#	$debugTextBox = New-Object system.Windows.Forms.TextBox
	$debugTextBox.text = ""
	$debugTextBox.width = 200
	$debugTextBox.autosize = $true
	$debugTextBox.location = New-Object System.Drawing.Point(20,140)
	$Form.Controls.add($debugTextBox)#>
	$Form.Controls.add($ScannerLabel)
	$Form.Controls.add($ScannerComboBox)
	$Form.Controls.add($ModeLabel)
	$Form.Controls.add($ModeComboBox)
	$Form.Controls.add($FormatLabel)
	$Form.Controls.add($FormatComboBox)
	$Form.Controls.add($ResolutionLabel)
	$Form.Controls.add($ResolutionComboBox)
	$Form.Controls.add($ScanButton)
	$Form.Controls.add($ConvertButton)
	$Form.Add_Shown({$Form.Activate()})
	$Form.ShowDialog()  | Out-Null
	$Form.Dispose()