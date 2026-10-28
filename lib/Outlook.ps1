# Classic Outlook: finding the mail an index row stands for, and the command
# line that shows it. The index's mapi16:// URL cannot be opened directly
# (a shell namespace, not a command), so the plugin finds the item over COM
# and hands WinShun `OUTLOOK.EXE /select outlook:<EntryID>` to start: WinShun
# is in front when Enter is pressed, so Outlook may come to the front too.

function Get-OutlookExe {
    $key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
    $exe = $null
    try { $exe = (Get-ItemProperty $key -ErrorAction Stop).'(default)' } catch { }
    if ($exe -and (Test-Path $exe)) { $exe } else { 'outlook.exe' }
}

function ConvertTo-OutlookFolderPath([string] $IndexFolder) {
    # The index's /store/folder/sub → Outlook's \\store\folder\sub.
    '\\' + ($IndexFolder.Trim('/') -replace '/', '\')
}

function Get-OutlookSelectArgs([string] $Target) {
    # /select takes outlook:<EntryID> or outlook:\\store\folder. A double quote
    # (or a final backslash, which escapes the closing quote) would let the
    # target end the argument and add others, so such a target is refused.
    if ($Target.Contains('"') -or $Target.EndsWith('\')) { throw "Cannot hand this folder name to Outlook: $Target" }
    "/select `"outlook:$Target`""
}

function Get-OutlookApp {
    # The running Outlook, or $null. Never CoCreate: when the running Outlook
    # cannot be reached yet, that would start a second one, which fails.
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { }
    $null
}

function Test-OutlookRunning { [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) }

function Get-OutlookStartArgs {
    # How COM starts Outlook: no window, and it stays up. A following /select
    # then opens only the mail, not the main window.
    '-Embedding'
}

function ConvertFrom-OutlookIndexUrl([string] $Url) {
    <#
    The mail's entry ID from its index address: the last segment of
    mapi16://…/store/folder/<id> (an attachment's /at=… dropped) holds it one
    byte per character, as U+AC00 + byte. That PST/OST entry ID (flags, store
    UID, node) is accepted by GetItemFromID and by OUTLOOK.EXE /select,
    without asking Outlook first. '' when the address has none.
    #>
    $segments = @(($Url -replace '/at=[^/]*$', '').Split('/'))
    $last = $segments[-1]
    if (-not $last) { return '' }
    $hex = New-Object Text.StringBuilder
    foreach ($c in $last.ToCharArray()) {
        $b = [int]$c - 0xAC00
        if ($b -lt 0 -or $b -gt 0xFF) { return '' }
        [void]$hex.Append($b.ToString('X2'))
    }
    $hex.ToString()
}

function Get-OutlookFolder($App, [string] $IndexFolder) {
    $parts = @($IndexFolder.Trim('/').Split('/') | Where-Object { $_ })
    if ($parts.Count -eq 0) { throw "No folder in '$IndexFolder'" }
    try {
        $folder = $App.Session.Folders.Item($parts[0])
        foreach ($p in ($parts | Select-Object -Skip 1)) { $folder = $folder.Folders.Item($p) }
    } catch {
        throw "Folder not found in Outlook: $IndexFolder"
    }
    $folder
}

function Find-OutlookMail {
    <#
    The mail in the index folder received at DateUtc (±1 minute; DASL
    compares in UTC) whose subject is Subject. When the subject differs
    but only one mail has that time, that one. $null when none.
    #>
    param($App, [string] $IndexFolder, [datetime] $DateUtc, [string] $Subject)
    $folder = Get-OutlookFolder $App $IndexFolder
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $lo = $DateUtc.AddMinutes(-1).ToString('yyyy-MM-dd HH:mm', $inv)
    $hi = $DateUtc.AddMinutes(1).ToString('yyyy-MM-dd HH:mm', $inv)
    $filter = "@SQL=`"urn:schemas:httpmail:datereceived`" >= '$lo' AND `"urn:schemas:httpmail:datereceived`" <= '$hi'"
    $hits = $folder.Items.Restrict($filter)
    $first = $null
    foreach ($item in $hits) {
        if ($null -eq $first) { $first = $item }
        if ([string]$item.Subject -eq $Subject) { return $item }
    }
    if ($hits.Count -eq 1) { return $first }
    $null
}

function Get-OutlookMailSubject([string] $AttachmentSubject, [string] $Name) {
    # The index gives an attachment the subject "<file name> (<mail subject>)".
    $prefix = "$Name ("
    if ($AttachmentSubject.StartsWith($prefix) -and $AttachmentSubject.EndsWith(')')) {
        return $AttachmentSubject.Substring($prefix.Length, $AttachmentSubject.Length - $prefix.Length - 1)
    }
    ''
}

function Get-OutlookAttachmentRoot { Join-Path $env:TEMP 'harafuhub.outlook' }

function Clear-OutlookAttachments {
    <# Removes the attachments saved more than a day ago (one folder per save). #>
    param([string] $Root = (Get-OutlookAttachmentRoot))
    if (-not (Test-Path $Root)) { return }
    $cutoff = (Get-Date).AddDays(-1)
    Get-ChildItem $Root -Directory | Where-Object { $_.CreationTime -lt $cutoff } |
        ForEach-Object { try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop } catch { } }
}

function ConvertTo-SafeFileName([string] $Name) {
    <#
    An attachment's name as a plain file name: only the part after the last
    \ or /, no characters Windows forbids, no leading or trailing dots and
    spaces, no device name (CON, NUL, COM1 …); 'attachment' when nothing is left.
    #>
    $leaf = ($Name -split '[\\/]')[-1]
    $bad = [regex]::Escape(-join [IO.Path]::GetInvalidFileNameChars())
    $leaf = ($leaf -replace "[$bad]", '_').Trim(' ', '.')
    if (-not $leaf) { return 'attachment' }
    if ($leaf -match '^(?i)(con|prn|aux|nul|com\d|lpt\d)(\.|$)') { $leaf = "_$leaf" }
    if ($leaf.Length -gt 150) { $ext = [IO.Path]::GetExtension($leaf); if ($ext.Length -gt 20) { $ext = '' }; $leaf = $leaf.Substring(0, 150 - $ext.Length) + $ext }
    $leaf
}

function Set-DownloadedMark([string] $Path) {
    # Mark of the Web (Internet zone), as a browser sets it: Office opens the
    # file in Protected View and blocks its macros, as for a mail attachment.
    Set-Content -LiteralPath $Path -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3" -Encoding Ascii
}

function Save-OutlookAttachment {
    <#
    Saves the attachment Name of the mail found by Find-OutlookMail into a
    new folder under Root (the file keeps its cleaned name, so the program
    that opens it shows it), marks it as from the Internet and returns its
    path. Throws when the mail or the attachment is not there, or when the
    path would lie outside Root.
    #>
    param($App, [string] $IndexFolder, [datetime] $DateUtc, [string] $MailSubject, [string] $Name,
        [string] $Root = (Get-OutlookAttachmentRoot), [string] $EntryId)
    $mail = $null
    if ($EntryId) { try { $mail = $App.Session.GetItemFromID($EntryId) } catch { } }
    if ($null -eq $mail) { $mail = Find-OutlookMail $App $IndexFolder $DateUtc $MailSubject }
    if ($null -eq $mail) { throw "Mail not found in Outlook (moved or deleted?)" }
    $attachment = $null
    foreach ($a in $mail.Attachments) {
        if ([string]$a.FileName -eq $Name -or [string]$a.DisplayName -eq $Name) { $attachment = $a; break }
    }
    if ($null -eq $attachment) { throw "Attachment $Name not found in the mail" }
    Clear-OutlookAttachments -Root $Root
    $dir = Join-Path $Root ((Get-Date).ToString('yyyyMMdd-HHmmss-') + [guid]::NewGuid().ToString('N').Substring(0, 6))
    New-Item -ItemType Directory -Force $dir | Out-Null
    $path = [IO.Path]::GetFullPath((Join-Path $dir (ConvertTo-SafeFileName $Name)))
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    if (-not $path.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing to save outside $Root" }
    $attachment.SaveAsFile($path)
    Set-DownloadedMark $path
    $path
}
