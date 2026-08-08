# Event Ticketing Platform V6

V6 is the feature-complete shared ticketing platform with strict ClientID isolation.

Main modules:

- `admin.html` — system/event administration
- `client.html` — client-only dashboards and exports
- `config.html` — secure client configuration, branding, preview, approval
- `distribute.html` — distributor batch management
- `v.html` — individual voucher claim / ticket rendering
- `claim.html` — voucher-required help page
- separate `event-gate-v6/` — gate-only operations

V6 adds ClientID isolation, Client Portal, event health, full readiness simulator, reminder center, optional automated email reminders, internal guest notes, live check-in progress, recent activity, global search, client-scoped exports, and emergency read-only mode.

See `DEPLOYMENT_GUIDE.md` for complete installation and folder setup.
