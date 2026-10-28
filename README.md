# Outlook mail · harafuhub.outlook

A resident WinShun plugin in Windows PowerShell. `ol <words>` (or `mail <words>`) searches the
mail classic Outlook keeps in its PST/OST files, through the Windows Search index that Outlook
already maintains: results in milliseconds, Outlook need not be running, nothing goes online.

| Type | Finds |
|---|---|
| `ol` | the 20 newest mails |
| `ol rechnung müller` | mails with words starting `rechnung` and `müller` (subject, body, sender, attachment names) |
| `ol "booking confirmation"` | the exact phrase |
| `ol from:anna` · `ol to:bob` | sender, or a To/Cc recipient |
| `ol in:done` | folder path contains a word starting `done` |
| `ol has:att` | only mails with attachments (and matching attachments) |
| `ol type:pdf rechnung` | only PDF attachments containing `rechnung*` (`type:xlsx,docx`, `doc`, `xls` too) |
| `ol … sort:old` | order: `new` (default), `old`, `size`, `from`, `subject` |
| `ol … date:7d` | received: `today`, `yesterday`, `week`, `month`, `year`, `<n>d`, or a year / month / day: `2026`, `2026-10`, `2026-10-21`, `21.10.2026` |
| `ol … after:2026-10-01 before:15.10.2026` | received from / before a date (same forms) |
| `ol … size:>1mb` | size: `>1mb`, `<500kb`, `1mb-100mb` (no unit: MB) |

**WinShun 0.15+ search panel (Ctrl+E):** every row carries its received date (Date modified and
Created columns) and size, so the columns show, sort and filter mail and attachments. The panel's
Date modified, Size and file type / Documents filters go straight into the index query (with
`date:`/`size:` typed as well, the narrower bound wins). Older WinShun versions ignore all this;
there `date:`, `size:` and `type:` do the filtering.

With words to search for, matching **attachments** (PDF, DOCX, XLSX, DOC, XLS: name and
content) are listed beside the mails, each as its own row with a file icon (Settings: off).

On a mail: **Enter** opens it in Outlook · **Ctrl+Enter** copies subject and sender ·
**Shift+Enter** shows its folder in Outlook. On an attachment: **Enter** saves it to
`%TEMP%\harafuhub.outlook` (cleared after a day) and opens it · **Ctrl+Enter** copies the file name ·
**Shift+Enter** opens its mail. Hovering a row shows the start of the text.

A saved attachment keeps only its file name (any `..\` or other path in it is dropped, and the
plugin checks the file lands inside that folder) and is marked as downloaded from the Internet,
so Office opens it in Protected View and blocks its macros, as it would from Outlook.

`sort:from` and `sort:subject` order the 200 newest matches (the index takes seconds to sort a
whole mailbox by them); the other orders cover every match.

How opening works: the index's `mapi16://` address cannot be started directly, so the plugin
finds the mail over Outlook COM (folder from the index, received time and subject) and asks
WinShun to start `OUTLOOK.EXE /select outlook:<EntryID>`. When Outlook is closed, Enter has
WinShun start it and says so, without waiting for it; press Enter again once it is open.
Folders whose name contains a `"` are not handed to Outlook's command line (Shift+Enter says so).

**AI:** with Settings → AI assistant → "AI may use plugins" on, the AI can run this search, and
what it finds goes to the AI: mail subjects, senders, dates, folders, attachment names and the
start of each mail's or attachment's text.
**给 AI 用**：打开"AI 可以使用插件"后，AI 可以调用这个搜索，邮件主题、发件人和部分正文会交给 AI。

**Permissions: none.** What it touches, all on this PC as you:

- reads the Windows Search index (OLE DB, `Search.CollatorDSO`), read-only;
- talks to a running classic Outlook over COM to find a mail, and asks WinShun to start
  `OUTLOOK.EXE /select …` to show it (or plain `OUTLOOK.EXE` when Outlook is closed); the plugin
  starts no programs itself;
- writes an attachment you open to `%TEMP%\harafuhub.outlook\` (with the Internet-zone mark) and
  deletes those copies after a day;
- never goes online itself; only with WinShun's AI option on do results reach the AI (see above).

Needs: classic Outlook (Microsoft 365 / 2016+) with Windows Search indexing its mail
(Outlook → File → Options → Search → Indexing Options). The new Outlook keeps no PST/OST and is
not searched.

Settings: attachments (on); default order (newest first); rows at most (50); also calendar,
contacts and tasks (off; these only open their folder).
