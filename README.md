# mailhero-pst-writer

The Outlook PST writer that ships inside MailHero
for macOS. MailHero runs this program as a separate helper process
(`MailHero.app/Contents/MacOS/mailhero-pst-writer`) and talks to it over
standard input, so the helper is a self-contained LGPL program and its complete
source is published here.

## License

GNU Lesser General Public License, version 3 or (at your option) any later
version. See `COPYING.LESSER`; the GNU General Public License v3 that it
supplements is in `COPYING`.

- `PSTFileFormat/` and `Utilities/` are a modified copy of
  **PSTFileFormat / ContinuMail**, Copyright (C) 2012-2017 ROM Knowledgeware,
  maintained by Tal Aloni, LGPL-3.0-or-later. Upstream:
  https://github.com/ROM-Knowledgeware/PSTFileFormat  
  Local changes are summarised in `PSTFileFormat/PSTFileFormat-MODIFICATIONS.md`
  and tracked in this repository's git history.
- `Program.cs` and the project file are MailHero's own code, Copyright (C) 2026
  Drake Allegrini, released under the same LGPL-3.0-or-later terms.

## What it does

```
mailhero-pst-writer OUTPUT.pst [HELPER_RAM_BYTES]
```

Reads newline-delimited JSON from stdin, one message per line, and writes a
Unicode PST (Outlook 2007+ compatible) at `OUTPUT.pst`. A non-zero exit status
and a message on stderr indicate failure. `HELPER_RAM_BYTES` is an advisory
memory budget passed by MailHero and may be omitted.

Each input line is an object with these fields (all optional unless noted):

| Field | Type | Notes |
|---|---|---|
| `subject` | string | |
| `sender_name`, `sender_email` | string | |
| `to`, `cc`, `bcc` | string[] | RFC 5322 mailbox strings |
| `body` | string | plain text |
| `html_body` | string | HTML alternative |
| `body_rtf_base64` | string | compressed RTF, Base64 |
| `message_id` | string | |
| `delivery_time` | integer | Windows FILETIME (UTC) |
| `folder_path` | string[] | folder hierarchy; defaults to `Inbox` |
| `transport_headers` | string | raw Internet headers |
| `attachments` | object[] | `filename` (required), `mime_type`, `content_id`, `content_location`, `data_base64`, `is_inline` |

## Building

Requires the .NET 8 SDK.

```sh
dotnet publish mailhero-pst-writer.csproj -c Release -r osx-arm64 \
  --self-contained true -p:PublishSingleFile=true -p:UseAppHost=true \
  -p:AssemblyName=mailhero-pst-writer -p:IncludeNativeLibrariesForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true -p:DebugType=None -o out/osx-arm64
```

Repeat with `-r osx-x64` and merge with `lipo -create` for a universal binary.
This is exactly what MailHero's `Tools/PSTWriter/build-macos.sh` does.

## Replacing the helper inside MailHero

Because the helper is a separate executable, you can build your own from this
source and substitute it:

1. Build as above.
2. Replace `MailHero.app/Contents/MacOS/mailhero-pst-writer` with your build.
3. Re-sign the app for local use, e.g. `codesign --force --deep --sign - MailHero.app`.
