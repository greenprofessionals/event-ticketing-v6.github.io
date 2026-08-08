# Event Ticketing Platform V6 — Deployment Guide

## 1. Recommended folder setup

Keep V5.x untouched as a rollback system. Create two NEW GitHub Pages folders:

```text
repository-root/
├── event-ticketing-v5/        # keep current working version
├── event-gate-v5/             # keep current working gate version
│
├── event-ticketing-v6/        # NEW V6 ticketing/admin/client/distribution
│   ├── index.html
│   ├── admin.html
│   ├── client.html
│   ├── config.html
│   ├── distribute.html
│   ├── v.html
│   ├── claim.html
│   ├── shared.js
│   ├── styles.css
│   ├── site-config.js
│   ├── site-config.example.js
│   └── images/
│       ├── admin-og.jpg
│       ├── client-og.jpg
│       ├── config-og.jpg
│       ├── distribution-og.jpg
│       └── voucher-og.jpg
│
└── event-gate-v6/             # NEW V6 gate-only module
    ├── index.html
    ├── gate-shared.js
    ├── styles.css
    ├── site-config.js
    ├── site-config.example.js
    └── images/
        └── gate-og.jpg
```

Public URLs after GitHub Pages publishes:

```text
https://greenprofessionals.github.io/event-ticketing-v6/
https://greenprofessionals.github.io/event-ticketing-v6/admin.html
https://greenprofessionals.github.io/event-ticketing-v6/client.html
https://greenprofessionals.github.io/event-ticketing-v6/config.html
https://greenprofessionals.github.io/event-ticketing-v6/distribute.html
https://greenprofessionals.github.io/event-ticketing-v6/v.html
https://greenprofessionals.github.io/event-gate-v6/
```

`claim.html` is a help/landing page explaining that a voucher is required. The real recipient claim page is `v.html?voucher=...`.

## 2. Google backend setup

For the safest rollout, create a NEW Google Sheet named something like:

```text
Event Ticketing Platform V6
```

Open **Extensions → Apps Script** and replace the starter code with `event-ticketing-v6/Code.gs`.

Run these functions in order:

```text
setupV6System()
bootstrapOwner()
createClientConfigurationForm()
```

Optional, after you are ready to use automatic email reminders:

```text
installV6DailyReminderTrigger()
```

The reminder trigger processes only Active events where **Automated email reminders** is enabled. WhatsApp, text, and phone actions remain device-initiated.

## 3. Deploy Apps Script

In Apps Script:

```text
Deploy → New deployment → Web app
Execute as: Me
Who has access: Anyone
```

Copy the URL ending in `/exec`.

Put the same `/exec` URL into BOTH:

```text
event-ticketing-v6/site-config.js
event-gate-v6/site-config.js
```

Example:

```javascript
window.EVENT_TICKETING_CONFIG = {
  ENDPOINT: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',
  GATE_URL: 'https://greenprofessionals.github.io/event-gate-v6/'
};
```

For later backend changes, keep the same deployment URL:

```text
Deploy → Manage deployments → Edit → New version → Deploy
```

## 4. Upload GitHub folders

Upload the complete `event-ticketing-v6` folder and the complete `event-gate-v6` folder to the repository root. Do not flatten the images folders.

Wait for GitHub Pages to publish, then open the V6 landing page and Admin page.

## 5. Create clients and client users

Log into:

```text
https://greenprofessionals.github.io/event-ticketing-v6/admin.html
```

As System Owner:

1. Open **Client Management**.
2. Create a unique Client ID, for example `SLPP-NY`.
3. Create client portal users in **Access Management**.
4. Choose one of these roles:
   - `CLIENT_ADMIN`
   - `CLIENT_VIEWER`
   - `CLIENT_FINANCE`
5. Set **Client scope** to the client's ID, for example `SLPP-NY`.
6. Set **Event scope** to `*` if the client may see all of its own events, or list specific Event IDs.

The backend enforces ClientID before EventID. A client cannot access another client's event simply by changing a URL or request parameter.

## 6. Create an event

In Admin → **Create Event Shell & Client Link**:

```text
Client ID: SLPP-NY
Event ID: NY2026EX
Working title: Executive Inauguration Ball
Client name/email/phone: client contact
```

Admin generates a private configuration link such as:

```text
https://greenprofessionals.github.io/event-ticketing-v6/config.html?event=NY2026EX&key=SECURE_TOKEN
```

Send it by Email, WhatsApp, Text, or Call.

## 7. Client configuration workflow

The client uses the secure configuration link to:

1. Complete the embedded Google Form.
2. Upload logo and optional background.
3. Refresh submission status.
4. Load the ticket preview.
5. Review every tier/color/price.
6. Revise if needed.
7. Approve the design.

The client later monitors operations through:

```text
https://greenprofessionals.github.io/event-ticketing-v6/client.html
```

## 8. Event readiness and activation

Admin loads the event and runs:

```text
Full Certification Test
```

V6 tests the current client configuration, every ticket tier, fixed/open vouchers, claims, QR behavior, capacity rules, gate check-in, duplicate blocking, walk-ins, payment states, search, and reporting logic in a non-production simulator.

Event Health displays:

```text
READY
NEEDS ATTENTION
BLOCKED
```

Activation requires:

```text
Client Approved + current Full Certification PASS
```

A System Owner can override the simulator gate, and the override is audited.

## 9. Voucher generation and distribution

After activation, Admin generates a voucher batch. The system creates a distributor link:

```text
https://greenprofessionals.github.io/event-ticketing-v6/distribute.html?batch=SECURE_BATCH_TOKEN
```

The distributor sends individual vouchers by Email, WhatsApp, Text, or contacts the recipient by Call.

Individual recipient links use:

```text
https://greenprofessionals.github.io/event-ticketing-v6/v.html?voucher=SECURE_VOUCHER_TOKEN
```

## 10. Gate deployment and operations

Gate staff use only:

```text
https://greenprofessionals.github.io/event-gate-v6/
```

Create Gate Staff / Gate Supervisor users in Admin Access Management and assign Event scope.

Gate features include QR scan, photo scan, manual lookup, duplicate protection, guest search, walk-ins, payment exceptions, internal guest notes for supervisors, and undo check-in for supervisors.

## 11. Client Portal

Client portal users log into:

```text
https://greenprofessionals.github.io/event-ticketing-v6/client.html
```

They see ONLY data allowed by their ClientID and Event scope:

- current/upcoming events
- event health
- tickets issued
- pending vouchers
- check-in progress
- capacity utilization
- money collected/outstanding
- recent activity
- client-scoped search
- CSV export
- Print / Save PDF report

They never receive the master Google Sheet.

## 12. Lightweight V6 operational safeguards

### Emergency read-only mode
Admin can enable it per event. It pauses configuration, voucher issuance, ticket claims, payments, transfers, QR reissues, and ticket status changes while Gate search/check-in remains available.

### Reminder Center
Admin can load unclaimed vouchers and pending payments, then use Email, WhatsApp, Text, or Call.

### Automated email reminders
If `installV6DailyReminderTrigger()` has been run and the event's automated-reminder toggle is enabled, Apps Script sends daily email reminders for eligible Active events within 14 days of the event.

### Guest notes
Internal notes are available only to System Owner, Event Admin, and Gate Supervisor.

### Recent activity
Uses the audit log, so no separate activity database is required.

## 13. Rollback

Do not delete V5.x. If V6 has a deployment issue, continue using the existing V5 folders and backend until V6 is corrected. Because V6 is deployed in separate GitHub folders and is recommended to use a separate V6 Google Sheet/backend during testing, rollback does not require changing production V5 files.
