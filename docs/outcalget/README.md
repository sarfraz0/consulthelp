# outcalget

Exports configured Outlook Classic calendars, with named configurations and coordinated concurrent runs.

Implementation: [`outcalget.ps1`](../../outcalget/outcalget.ps1), version 0.3.5.

## Documentation

- [Parameters, examples, results and exit codes](usage.md)
- [Scheduling, configuration, synchronization and recovery](operations.md)
- [Calendar Git update helper](update-calendar.md)
- [Implementation rationale and existing exceptions](implementation-notes.md)
- [Historical verification and this extraction](verification.md)
- [Native PowerShell help](outcalget-help.xml)

## Behavior and prerequisites

outcalget 0.3.5. One Windows PowerShell 5.1 file; no administrator rights, runtime
modules/downloads, Graph permissions, endpoint exclusions or policy bypass.
The native Outlook exporter preserves Full details, today through today + 29 LOCAL
calendar dates inclusive, unrestricted working hours and saved privacy/attachment
choices. Private details and attachments default to off during new setup.
Existing schema-1 configurations from 0.1.x and 0.2.0 are accepted unchanged.

No -Profile means the existing default config.json. -Profile default selects that
same file. -Profile team selects profiles\team.json beneath the existing local state
root. A missing named profile fails with CONFIG_MISSING, exit 2: no default fallback
and no setup prompt. -Configure -Profile team creates it after a successful test
export. Named setup refuses an existing profile; use a new name to create another
binding. -Config remains available for existing explicit configuration paths and
cannot be combined with -Profile. Export profile names are NOT Outlook MAPI profile
names; the configuration's profileName still binds the actual Outlook session.

Calendar setup offers (1) the original Outlook folder picker, (2) another mailbox's
shared default Exchange calendar by full owner email, and (3) accessible calendars
already added to Outlook's Calendar navigation list. Option 3 addresses shared/linked
calendars absent from PickFolder. Only option 2 uses GetSharedDefaultFolder; it does
not select arbitrary secondary shared calendars. Both paths use the same native
GetCalendarExporter/SaveAsICal operation. Outlook sharing/export permissions apply.
No email, invitation or sharing request is sent.
For validated option-2 shared-default bindings, Outlook can return Folder.Store=null
by design. This is accepted after Exchange owner identity has been revalidated;
native Application Folders synchronization still runs before an owned-session export.
A null Store on an ordinary folder binding or an explicit non-Exchange store fails
cleanly instead of claiming synchronization. Each export additionally inspects the
native CalendarSharing.Folder and uses NameSpace.CompareEntryIDs to verify that the
exporter targets the resolved calendar before SaveAsICal. Shared mailbox identity
pinning also uses CompareEntryIDs because different entry-ID representations can
refer to the same object. Native folder bindings additionally retain store validation.

Internet availability is required, using Windows Network List Manager. When Outlook
Classic is already open by the user, the helper attaches, performs no explicit sync
and leaves it open. Otherwise it starts the machine-registered Classic installation
in the saved MAPI profile for configured exports; interactive setup uses normal Outlook
startup/default-profile behavior and can display its usual profile/sign-in dialogs.
The helper waits for native readiness/online Exchange/group sync and requests normal
Quit only for the exact process it started. Other exports can join
that process. The creator waits for all joined usage leases before closing it; joined
runs synchronize their own selected calendars and never request Quit themselves.
Your ordinary configured Windows user must be logged on in the same desktop session.
-NonInteractive cannot suppress Office sign-in/security dialogs. Logged-off/server
Office automation and new Outlook are unsupported.

## Deployment and native help

All executable export logic remains in one standalone script. Deploy `outcalget.ps1` alone with its declared Windows/Outlook dependencies and runtime configuration; neither the XML nor Markdown is required for export or `-Version`. XML is optional and supplies detailed native help. Without it, `Get-Help`, `-?` and the existing `-Help` switch fall back to generated syntax. To retain detailed help, also deploy matching documentation with the repository layout (`outcalget/` beside `docs/`), so the native `.EXTERNALHELP` directive resolves `../docs/outcalget/outcalget-help.xml`.

Markdown is the editable documentation; update the corresponding external XML help whenever public help changes. The XML contains the extracted synopsis, description, parameters, examples, inputs, outputs, notes and links. It is documentation, not executable code or a runtime module.

## Vendor references

- https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.exchangeconnectionmode
- https://learn.microsoft.com/en-us/office/vba/api/outlook.olexchangeconnectionmode
- https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.compareentryids
- https://learn.microsoft.com/en-us/office/vba/api/outlook.calendarsharing.folder
- https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.store
- https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getshareddefaultfolder
- https://learn.microsoft.com/en-us/office/vba/api/outlook.navigationfolder.folder
- https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.logon
- https://learn.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-logontype-principaltype-element
