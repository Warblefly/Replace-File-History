# Replace-File-History
Powershell script (with Robocopy) to replace flaky File History on Windows 11

The Windows "File History" backup facility has become unsupported and increasingly unreliable in recent Windows builds. Furthermore, its error reporting is rather basic, and some of its Windows Registry keys, affecting its functionality, could kindly be described as "obscure".

Additionally, as the years have gone by, the user's ability to control the source of backed-up files has diminished into confusion. At present, it will only copy libraries, the desktop, favourites and contacts.

This applies when it works. As of August 1st 2026, the facility began silently failing.

A replacement was necessary, and this is it.

You can run this Powershell script straight from your Powershell command line, and the only parameter it MUST have is -BackupRoot.

john@johnwarburton.net
