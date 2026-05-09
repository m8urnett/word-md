# Sideload word-md

These instructions cover sideloading the **word-md** add-in for local development or private use.

The usual desktop Word method is to place the add-in manifest in a shared folder, register that folder as a trusted add-in catalog, then load the add-in from Word.

---

## 1. Confirm what kind of add-in you have

Most Word add-ins use an XML manifest file named something like:

```text
manifest.xml
```

The manifest is not the add-in code itself. It tells Word where the add-in web app is hosted.

Inside the manifest, look for something like:

```xml
<SourceLocation DefaultValue="https://m8urnett.github.io/word-md/taskpane.html"/>
```

That URL must be reachable by Word.

For local development, this usually means your add-in dev server must be running.

Example:

```powershell
npm start
```

or:

```powershell
npm run dev-server
```

The exact command depends on how the add-in project was created.

---

## 2. Create a local folder for sideloaded add-ins

Create a folder somewhere easy to find.

Example:

```text
C:\OfficeAddins
```

Copy your manifest file into that folder:

```text
C:\OfficeAddins\manifest.xml
```

You can put multiple add-in manifests in this same folder if you want.

---

## 3. Share the folder on Windows

Right-click the folder:

```text
C:\OfficeAddins
```

Then choose:

```text
Properties > Sharing > Advanced Sharing
```

Enable:

```text
Share this folder
```

Give it a simple share name:

```text
OfficeAddins
```

Click:

```text
Permissions
```

For local development, `Read` permission is enough.

Click:

```text
OK
```

Your network share path will look like this:

```text
\\YOUR-PC-NAME\OfficeAddins
```

To find your PC name, run:

```powershell
hostname
```

Example result:

```text
YOUR-PC
```

Then your share path would be:

```text
\\YOUR-PC\OfficeAddins
```

---

## 4. Add the shared folder as a trusted add-in catalog in Word

Open Microsoft Word.

Go to:

```text
File > Options
```

Then:

```text
Trust Center > Trust Center Settings
```

Then:

```text
Trusted Add-in Catalogs
```

In the catalog URL field, enter your shared folder path.

Example:

```text
\\YOUR-PC\OfficeAddins
```

Click:

```text
Add catalog
```

Check:

```text
Show in Menu
```

Click:

```text
OK
```

Close Word completely.

Reopen Word.

---

## 5. Load the add-in in Word

Open Word.

Go to:

```text
Home > Add-ins
```

Depending on your Word version, choose one of these:

```text
More Add-ins
```

or:

```text
Advanced
```

Open the tab named:

```text
Shared Folder
```

Your add-in should appear there.

Select it and click:

```text
Add
```

The add-in should now load in Word.

---

## 6. Make sure the web app is running

If the add-in appears but does not load correctly, start the web app referenced by the manifest.

For many Office Add-in projects, this is:

```powershell
npm start
```

The manifest may reference something like:

```text
https://localhost:3000/taskpane.html
```

That page must open successfully in your browser.

Test it directly:

```text
https://localhost:3000/taskpane.html
```

If the browser shows a certificate warning, you may need to trust the local development certificate.

For Office add-ins created with Microsoft tooling, this is often handled with:

```powershell
npx office-addin-dev-certs install
```

---

## 7. Common problems

### The add-in does not appear in the Shared Folder tab

Check these:

- The manifest file is directly inside the shared folder.
- The folder is actually shared.
- Word has been restarted after adding the trusted catalog.
- `Show in Menu` is checked.
- The share path is entered as a UNC path, not a local path.

Correct:

```text
\\YOUR-PC\OfficeAddins
```

Wrong:

```text
C:\OfficeAddins
```

---

### The add-in appears but will not load

Check the manifest's `SourceLocation`.

Example:

```xml
<SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
```

Make sure that URL opens in your browser.

Also make sure your dev server is running.

---

### Word blocks the add-in

Check these:

- The manifest is valid XML.
- The add-in web app uses HTTPS.
- The local certificate is trusted.
- The manifest uses valid URLs.
- The add-in ID is a valid GUID.
- You restarted Word after changing Trust Center settings.

---

### Changes to the manifest are not showing up

Word can cache add-ins.

Try:

1. Remove the add-in from Word.
2. Close Word.
3. Restart Word.
4. Add the add-in again from the Shared Folder tab.

If needed, clear the Office add-in cache.

On Windows, try deleting the contents of:

```text
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

You can open that folder by pressing:

```text
Win + R
```

Then entering:

```text
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

Close all Office apps before deleting cache files.

---

## 8. Sideloading in Word on the web

For Word Online, the process is different.

Open Word on the web.

Open a document.

Go to:

```text
Home > Add-ins
```

Then choose:

```text
More Add-ins
```

or:

```text
Advanced
```

Then select:

```text
Upload My Add-in
```

Upload your:

```text
manifest.xml
```

The add-in should load into the document.

---

## 9. Quick checklist

Before debugging anything complicated, confirm these:

- `manifest.xml` exists.
- `manifest.xml` is in a shared folder.
- Word has the shared folder registered under Trusted Add-in Catalogs.
- `Show in Menu` is checked.
- Word was restarted.
- The add-in was added from the `Shared Folder` tab.
- The web app URL in `SourceLocation` works in a browser.
- The web app is running.
- The web app uses HTTPS.
- The local dev certificate is trusted.

---

## 10. Minimal example workflow

Create the folder:

```powershell
mkdir C:\OfficeAddins
```

Copy the manifest:

```powershell
copy .\manifest.xml C:\OfficeAddins\
```

Share the folder manually through Windows Explorer.

Find your computer name:

```powershell
hostname
```

Add this shared path to Word's Trusted Add-in Catalogs:

```text
\\YOUR-PC-NAME\OfficeAddins
```

Restart Word.

Start your add-in web server:

```powershell
npm start
```

In Word, go to:

```text
Home > Add-ins > Advanced > Shared Folder
```

Select the add-in.

Click:

```text
Add
```