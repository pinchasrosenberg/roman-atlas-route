# Security

Please report vulnerabilities privately through GitHub's **Report a vulnerability** (Security → Advisories).

The atlas is a single static page with no backend and no user accounts. Leaflet is pinned to 1.9.4 with Subresource
Integrity. A Content-Security-Policy limits the page to Leaflet and the basemap tiles, and `connect-src 'none'`
means the page cannot send data anywhere.
