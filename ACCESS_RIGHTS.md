# V6 Access Rights

| Role | Client scope | Event configuration | Voucher admin | Client dashboard | Finance | Gate | Sensitive overrides |
|---|---|---|---|---|---|---|---|
| SYSTEM_OWNER | All | Full | Full | All clients | Full | Full | Full |
| EVENT_ADMIN | Assigned clients/events | Assigned events | Assigned events | Operational admin view | Assigned events | Assigned events | Ticket/admin actions |
| FINANCE | Assigned clients/events | No | No | Operational financial view | Assigned events | No | Payment actions |
| CLIENT_ADMIN | Own ClientID only | Via private config link | No | Own client/events | Own client | No | No |
| CLIENT_VIEWER | Own ClientID only | No | No | Read-only own client/events | Visible dashboard totals | No | No |
| CLIENT_FINANCE | Own ClientID only | No | No | Own client/events | Own client financial data | No | No |
| GATE_SUPERVISOR | Assigned events | No | No | No | Gate payment exception | Assigned events | Walk-in, undo, notes |
| GATE_STAFF | Assigned events | No | No | No | Minimal payment status | Assigned events | No |
| Distributor | Batch token only | No | Own batch only | No | No | No | No |
| Recipient | Voucher token only | No | No | No | Own ticket only | Presents ticket | No |

Client isolation is enforced in Apps Script. Browser-supplied ClientID values are never sufficient by themselves; the authenticated user's `ClientScope` must allow the ClientID and the event must also pass EventScope checks.
