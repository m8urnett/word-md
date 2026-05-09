# Deployment

## Preflight

1. Run `npm run check`.
2. Confirm `taskpane.html`, `taskpane.js`, `taskpane.css`, `vendor/`, `icons/`,
   and `manifest.xml` are published to the same HTTPS origin.
3. Open `taskpane.html` from the published origin and confirm the browser can
   load `vendor/marked.min.js`.
4. Verify the manifest URLs point to the published origin.

## Pilot sideload

Use Word's Upload My Add-in flow with `manifest.xml` for a small pilot group.
Test Word for Windows, Word for Mac, and Word on the web if those platforms are
in scope.

## Tenant deployment

For organization-wide deployment:

1. Open the Microsoft 365 admin center.
2. Go to Settings > Integrated apps.
3. Upload `manifest.xml`.
4. Assign the add-in to a controlled pilot group first.
5. Expand assignment after the pilot validates insertion behavior and support
   expectations.

## Rollback

Remove or unassign the app in Integrated apps. If the hosted package must be
rolled back, publish the previous tagged static files and matching manifest.
