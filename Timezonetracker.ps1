# טעינת Assemblies עבור Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# יצירת Form (חלון) עם עיצוב מודרני כהה
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Logon & Time Zone Tracker by Jonathan Brahmi"
$form.Size = New-Object System.Drawing.Size(430, 450)
$form.BackColor = [System.Drawing.Color]::FromArgb(38, 50, 56) # צבע רקע כהה כמו ב-Ubuntu
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.StartPosition = 'CenterScreen'

# פונקציה לעיצוב כפתור
function Style-Button {
    param ($button)
    $button.BackColor = [System.Drawing.Color]::FromArgb(33, 150, 243) # צבע כחול מודרני
    $button.ForeColor = [System.Drawing.Color]::White
    $button.FlatStyle = 'Flat'
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
}

# יצירת כפתור להתחלת התהליך
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start the Process"
$startButton.Location = New-Object System.Drawing.Point(10, 20)
$startButton.Size = New-Object System.Drawing.Size(120, 30)
Style-Button $startButton

# יצירת כפתור לחיפוש מחשב ב-AD
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Find computer"
$searchButton.Location = New-Object System.Drawing.Point(280, 20)
$searchButton.Size = New-Object System.Drawing.Size(120, 30)
Style-Button $searchButton

# יצירת כפתור לשמירת הנתונים ב-CSV
$buttonSaveCsv = New-Object System.Windows.Forms.Button
$buttonSaveCsv.Text = 'Export to CSV'
$buttonSaveCsv.Location = New-Object System.Drawing.Point(147, 20)
$buttonSaveCsv.Size = New-Object System.Drawing.Size(120, 30)
Style-Button $buttonSaveCsv

# יצירת שדה טקסט לחיפוש מחשב
$computerTextBox = New-Object System.Windows.Forms.TextBox
$computerTextBox.Location = New-Object System.Drawing.Point(10, 80)
$computerTextBox.Size = New-Object System.Drawing.Size(395, 30)
$computerTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 12)

# יצירת ListBox להצגת תוצאות
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 120)
$listBox.Size = New-Object System.Drawing.Size(395, 250)
$listBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# יצירת Label להצגת הודעות
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 500)
$statusLabel.Size = New-Object System.Drawing.Size(600, 40)
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)

# יצירת רשימה לשמירת נתונים לפני שמירה ל-CSV
$computersList = @()  # רשימה חדשה ריקה
$tempCsvPath = Join-Path $env:TEMP "computersInfo.csv"  # נתיב קובץ TEMP

# פונקציה לחיפוש מחשבים ב-AD
function Get-ADComputers {
    Write-Host "התחלת חיפוש ב-AD..."
    $computers = Get-ADComputer -Filter * -Properties Name, LastLogonDate
    $result = @()

    foreach ($comp in $computers) {
        $lastLogon = $comp.LastLogonDate
        $currentTime = Get-Date
        $isDST = [System.TimeZoneInfo]::Local.IsDaylightSavingTime($currentTime)

        $dstStatus = if ($isDST) { "Summer" } else { "Winter" }

        $resultObj = New-Object PSObject -property @{
            ComputerName = $comp.Name
            LastLogon = $lastLogon
            DstChange = $dstStatus
            LocalTime = $currentTime.ToString("yyyy-MM-dd HH:mm:ss")
        }

        $result += $resultObj
    }

    # הודעה לדיבוג: הצגת תוצאות חיפוש
    Write-Host "Wait.. : $($result.Count)"
    return $result
}

# פעולה על כפתור "התחל תהליך"
$startButton.Add_Click({
    $statusLabel.Text = "Start to scan"
    
    # חיפוש המחשבים
    $computersInfo = Get-ADComputers
    
    if ($computersInfo.Count -gt 0) {
        $computersList = $computersInfo  # עדכון רשימה עם תוצאות
        $statusLabel.Text = "הסקריפט הסתיים. הנתונים מוכנים לשמירה."
        
        # שמירת הנתונים בקובץ TEMP
        $computersList | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "DATA will be saved in TEMP: $tempCsvPath"
    } else {
        $statusLabel.Text = ")-: sorry, i do not found computer"
    }

    # הצגת תוצאות ב-Out-GridView
    $computersInfo | Out-GridView
})

# פעולה על כפתור "חפש מחשב"
$searchButton.Add_Click({
    $computerName = $computerTextBox.Text
    if ($computerName) {
        $compInfo = Get-ADComputerByName -computerName $computerName
        
        # אם נמצא מחשב, הצג את המידע בתיבת הטקסט
        if ($compInfo) {
            $listBox.Items.Clear()
            $listBox.Items.Add("$($compInfo.ComputerName) - $($compInfo.DstChange) - זמן מקומי: $($compInfo.LocalTime)")
            $statusLabel.Text = "המחשב נמצא: $($compInfo.ComputerName)"
        } else {
            $statusLabel.Text = ")-: do not find computer"
            $listBox.Items.Clear()
        }
    } else {
        $statusLabel.Text = "Please enter host name to search."
    }
})
# פעולה על כפתור "שמור לקובץ CSV"
$buttonSaveCsv.Add_Click({
    Write-Host "כפתור שמירה ל-CSV נלחץ."
    
    # אם יש נתונים בקובץ TEMP
    if (Test-Path $tempCsvPath) {
        # פתיחת תיבת דיאלוג לשמירת קובץ
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $saveFileDialog.Title = "Save copy as csv file"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $csvPath = $saveFileDialog.FileName
            
            # העתקת הנתונים מקובץ TEMP לקובץ שנבחר
            Copy-Item -Path $tempCsvPath -Destination $csvPath
            [System.Windows.Forms.MessageBox]::Show("Progress success -$csvPath", "הצלחה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show(")-: No DATA to save", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# הוספת רכיבי GUI לחלון
$form.Controls.Add($startButton)
$form.Controls.Add($computerTextBox)
$form.Controls.Add($buttonSaveCsv)
$form.Controls.Add($searchButton)
$form.Controls.Add($listBox)
$form.Controls.Add($statusLabel)

# הצגת ה-Form
$form.ShowDialog()