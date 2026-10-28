# Publishing updates

Teachers use **Check for updates** in the Windows app. GitHub accounts and Python
are not needed. The app checks only this repository's public releases; it never
sends spreadsheets, names, file paths, or logs. GitHub receives normal connection
information such as the IP address and app version.

## Release a change

1. Update the app and set `VERSION` in `app_updates.py` (for example `1.0.1`).
2. Run `python -m unittest discover -s tests -v`.
3. Commit and push only reviewed source changes.
4. Create and push the matching tag: `git tag v1.0.1`, then `git push origin v1.0.1`.
5. Check the **Actions** page. A Windows runner tests and builds a fresh EXE,
   then publishes it to Releases. A plain source push does not publish an update.

The app downloads only the named EXE from this repository and verifies GitHub's
SHA-256 checksum before running it. Installation replaces the current EXE and
restarts it. Keep the EXE in a writable folder, such as Downloads. A protected
installation folder may prevent updating. No administrator privileges are requested.
The app does not retain the current spreadsheet selection across restarts.

The initial updated EXE must be downloaded once; older EXEs have no update button.
Automatic installation is for Windows EXEs. Running the Python source on macOS
offers a link to the Windows release. Wine/Whisky compatibility is not guaranteed.
The EXE is not code-signed; Windows may display a reputation warning.

## Keep private files out

The root `.gitignore` allows only named source files and the blank template.
Do not force-add spreadsheets, output documents, old EXEs, ZIP files, pictures,
build folders, credentials, or caches. Test only with synthetic students.
Never upload the whole working folder. Old EXEs and ZIPs made before spreadsheet
bundling was removed may still contain student data; do not distribute them.

The template has been reviewed and stripped of author/SharePoint metadata. If
replacing it, review all document contents, embedded images and ZIP members first.
