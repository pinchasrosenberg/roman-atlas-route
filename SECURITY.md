# Security

Please report vulnerabilities privately through GitHub's **Report a vulnerability** (Security → Advisories).

The atlas is a single static page with no backend and no user accounts. Leaflet is pinned to 1.9.4 with Subresource
Integrity. A Content-Security-Policy limits the page to Leaflet, map tiles and Hebrew Wikipedia summaries. Wikipedia text is
inserted as plain text, and links and images are checked against an allow-list of Wikipedia domains.
