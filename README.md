# Bento Project Packer - BETA

Bento Project Packer finds all audio files referenced in a Bento `project.xml`, copies them next to the project, and updates the XML to use relative `.\filename` paths. This makes projects easy to share and archive.

## Features

- Accept a project XML path or a project directory (uses `project.xml` or the only `*.xml`).
- Search your sample folders recursively. If Everything (es.exe) is available it is used automatically; otherwise a normal filesystem scan runs.
- Copy found samples beside the project with collision safe names like `name (n).ext`.
- Update XML `filename` attributes to `.\<finalName>`.
- Always write a timestamped text report.
- Always create a pre-change backup `project.xml.bak` before saving changes.

## Safety

- Beta software. Run on copies of your projects and keep backups.
- Always run a dry run first to preview changes.

## Requirements

- Windows with PowerShell 5.1 or PowerShell 7+.
- Optional: Everything (Voidtools). The tool will use `es.exe` automatically when found.
  - Quick setup: see `INSTALL-Everything.md` or run `./Install-Everything.ps1`.

## Files

- `Run-BentoProjectPacker.ps1` - runner you execute.
- `Bento-ProjectPacker.ps1` - function library used by the runner.
- `SampleRoots.txt` - optional list of sample roots (one path per line).

Place these files together in the same folder.

## Quick Start

1) Create or edit `SampleRoots.txt` (one path per line) to list your sample folders.
2) Run a dry run to preview the changes.

```
./Run-BentoProjectPacker.ps1 "C:\Bento SD Backup\Projects\TRACK NAME\project.xml" -DryRun
```

Or pass the project directory:

```
./Run-BentoProjectPacker.ps1 "C:\Bento SD Backup\Projects\TRACK NAME" -DryRun
```

Apply changes (a `project.xml.bak` backup is always created first):

```
./Run-BentoProjectPacker.ps1 "C:\Bento SD Backup\Projects\TRACK NAME\project.xml"
```

Override sample search root for a one-off run (bypasses SampleRoots.txt):

```
./Run-BentoProjectPacker.ps1 "C:\Bento SD Backup\Projects\TRACK NAME\project.xml" -SearchRoot "D:\Samples"
```

## Usage

```
./Run-BentoProjectPacker.ps1 <ProjectXmlOrDir> [-DryRun] [-SearchRoot <path>] [-Report <path>]
```

Parameters:

- `-DryRun` - no file copies or XML writes; decisions are printed and a report is generated.
- `-SearchRoot <path>` - override the sample search roots for this run. Otherwise uses `SampleRoots.txt` (or `%USERPROFILE%\Music`).
- `-Report <path>` - write a text report to this file. A timestamped report is always written next to the XML and a copy in the current directory.

## Notes

- Targets all `<params>` XML nodes with a `filename` attribute.
- Path updates are strictly set to `.\<file>`.
- Everything CLI (`es.exe`) speeds up large library searches. If not installed, the script performs a normal recursive scan.

## Troubleshooting

- No candidates found in report:
  - Check `-SearchRoot` or your `SampleRoots.txt` path list.
  - Try a smaller subtree for validation.
- Slow scans:
  - Install Everything and ensure `es.exe` is in PATH, or specify its default install location.
  - Limit the search by pointing `-SearchRoot` to a narrower folder.
- Multiple XML files in folder:
  - Specify the exact XML file instead of the directory.

## TODO

- Group Samples into subfolders as root project folder gets crowded, especially with MultiSample patches

  
## License

MIT License © 2025 The Dark Glove - Logical Perspective Ltd

See LICENSE for full text. Provided "as is" without warranty; always work on copies and keep backups.
