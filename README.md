# FSTool: Folder Shortcut Tool

- Version: 1.0
- Requires: Windows PowerShell 2.0

## Summary

In the specified folder, and optionally all subfolders, finds all shortcuts to folders and replaces them with folder shell links.
Also reports on file & folder paths which exceed the 259-character limit of Windows APIs.

### Usage

To use the script, launch a PowerShell instance, and in the console, type the name of the script, optionally followed by a
directory path (in quotes, if it contains spaces). The folder tree starting at that point, or the current folder if none is
given, will be recursively scanned for shortcut files pointing to folders. The script tells the user what they are, prompts
for confirmation, then converts the links to "folder shortcuts", a.k.a. "folder shell links".

### Notes

A folder shell link is not an ordinary shell link (shortcut file). Instead, it is a special folder which, thanks to a hidden
desktop.ini file declaring it to be a shell link, behaves just like a symbolic link in Explorer. Other apps see it as what it
really is: a folder containing a hidden desktop.ini file, along with a regular shortcut file pointing to the target folder.

At some point this script may be enhanced, e.g. to reverse the process or to add support for converting to/from symlinks.
