# Install or recover the executable

Read the canonical
[installation guide](https://github.com/sjunepark/excel-tovalue/blob/main/README.md#cli-설치)
for prerequisites, platform selection, release downloads, checksum
verification, installation, `PATH`, and upgrades. Installation facts live
there so this package does not maintain a second installer recipe.

The repository and its
[releases](https://github.com/sjunepark/excel-tovalue/releases) are private,
so an unauthenticated fetch of the guide fails. Read it through the user's
existing GitHub CLI login (`gh auth status`), which the downloads also need:

```sh
gh api -H "Accept: application/vnd.github.raw" \
  repos/sjunepark/excel-tovalue/contents/README.md
```

If `gh` is missing or not authorized for the repository, report that and ask the
user for the release files; do not build from source, and do not start a
login flow yourself.

When the user has authorized installation or an update, do the work yourself
with the guide's unattended commands. Supported targets are Windows x64 and
Linux x64 (glibc); on any other platform, report that no executable is
published. Keep the archive, `SHA256SUMS.txt`, and installer from the same
release.

There is no self-upgrade command. Upgrading means running the installer from
a newer release over the existing installation.

Verify `excel-formula-remover --version` first through the executable's full
path, then by command name. File visibility and `PATH` lookup are separate
facts. On Windows, a shell started by a packaged desktop app can see files no
other program sees, so run the guide's `verify-windows-install.ps1` check
rather than trusting your own shell or asking the user to open a terminal.

After verification, resume the original task.
