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

V6.1 Admin/Configuration patch

Changes:
1. Admin header corrected to V6.
2. Create Event Shell now includes Client ID.
3. Added Client Portal Role: CLIENT_ADMIN, CLIENT_VIEWER, CLIENT_FINANCE.
4. Added optional temporary portal passcode and event scope.
5. Access Management now includes client roles and Client scope.
6. Generated configuration links are normalized to the V6 TICKETING_URL from site-config.js, even if the backend returns an old V5 base URL.
7. Client Configuration Link button also normalizes links to V6.
8. Configuration-page diagnostic now references event-ticketing-v6 instead of event-ticketing-v5.

Replace only admin.html and config.html in the event-ticketing-v6 GitHub repository.
No Apps Script redeployment is required for the V5-to-V6 link correction itself.

Important: Creating a client portal login from the Create Event section requires the deployed V6 Apps Script backend to accept clientScope on createAccessUser. If your backend does not yet support client-scoped access users, the event still gets created, but the page will report that the portal login could not be created. You can still create the client user from Access Management after the backend is updated.
