# CRMbook Outlook BCC add-in

Adds the configured CRMbook sync email address to BCC while composing a message. Existing BCC recipients are preserved, and resolved duplicates are skipped.

## Install

Visit https://rosholmde.github.io/crmbook-bcc/ and download the Outlook installation file. Remove the previous CRMbook BCC add-in, then use https://aka.ms/olksideload → My add-ins → Custom Addins → Add a custom add-in → Add from File to install the new manifest.xml. Your organization may require administrator installation.

Compose an email, click Add CRM BCC in the ribbon or Apps menu, and inspect BCC before sending. Live Outlook testing remains necessary.

## Hosting

GitHub Pages is configured to publish the main branch root over HTTPS. Keep commands.html, commands.js, manifest.xml, index.html, and the five PNG icons together at the repository root. The .nojekyll file disables Jekyll processing.

The configured CRMbook inbound email address is publicly readable in commands.js. No mailbox contents are hosted here.
