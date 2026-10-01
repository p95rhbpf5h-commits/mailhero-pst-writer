# PSTFileFormat — modifications relative to upstream

This directory is a modified copy of the `ROM-Knowledgeware/PSTFileFormat`
library (ContinuMail), LGPL-3.0-or-later. All MailHero modifications remain
under the same license.

The consolidated modification log that earlier builds referenced
(`vendor/PSTFileFormat-MODIFICATIONS.md`) was not carried into this repository.
From the first commit of this repository onward, every change to the library is
recorded in git history: `git log -- PSTFileFormat`.

Known categories of local change, from the code itself:

- Writer-side support used by `Program.cs`: default-store blueprint
  (`DefaultStoreTemplates*.cs`, `PSTFile.Create.cs`) so an empty Unicode PST
  can be created without an existing template file.
- Bare `RecurrencePattern` handling for `PidLidTaskRecurrence`.
- Build adjustments for .NET 8 (`PSTFileFormat.csproj`, compiled into the
  single `mailhero-pst-writer` executable rather than a separate assembly).
