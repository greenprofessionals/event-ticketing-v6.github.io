# Event Ticketing Platform V6 — Release Notes

V6 is the shared-platform, strict-client-isolation release.

New in V6:

- `ClientID`-scoped shared database architecture
- Client Portal (`client.html`)
- client roles: CLIENT_ADMIN, CLIENT_VIEWER, CLIENT_FINANCE
- client-scoped event dashboards and CSV exports
- ticket/check-in progress monitoring
- event health: READY / NEEDS ATTENTION / BLOCKED
- full readiness simulator remains required before activation
- global search across authorized events only
- recent activity feed from the audit log
- internal guest notes for Admin and Gate Supervisor
- Reminder Center with Email, WhatsApp, Text, and Call actions
- optional automated daily email reminders
- emergency read-only mode that preserves gate search/check-in
- dedicated Admin, Client, Configuration, Distribution, Voucher Claim, and Gate OG images
- separate V6 GitHub folder structure so V5 remains available for rollback

The backend enforces ClientScope before EventScope. Client users cannot expose another client's data by changing query strings or request payloads.
