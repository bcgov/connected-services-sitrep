# BC Gov Connected Services SitRep Dashboard

Live URL: https://bcgov.github.io/connected-services-sitrep/sitrepdash.html
Repo: https://github.com/bcgov/connected-services-sitrep (public, Apache 2.0, files in `/docs`)

## Stack
Vanilla HTML/CSS/JS, GitHub Pages, SharePoint via Microsoft Graph API, MSAL auth

## File structure
```
docs/
├── sitrepdash.html
├── css/sitrepdash.css
├── js/
│   ├── config.js       (CONFIG, TEAMS, shared state)
│   ├── utils.js        (esc, safeJSON, getWeekLabel, getWeekOptions, showToast)
│   ├── auth.js         (MSAL, getToken, signIn)
│   ├── sharepoint.js   (Graph API calls, page states)
│   ├── render.js       (renderGrid, updateSummary, history view)
│   └── ui.js           (init, modal, coordinator, meetings, filter, CSV)
├── fonts/              (BC Sans woff2)
└── images/             (BC Gov logos)
```

## Azure AD App
Configured for SPA authentication with Microsoft Graph API permissions for SharePoint access.

## SharePoint
Uses SharePoint lists for data storage and retrieval via Microsoft Graph API.


- MSAL auth with full-page sign-in screen when not authenticated
- Team cards read from SharePoint, card edits write back via Graph API (PATCH existing / POST new)
- Coordinator panel (news, meetings, featured highlights/blockers) reads/writes SharePoint with 800ms debounce save
- History view — browse all past weeks, click week to see that week's cards
- Week picker in edit modal — can submit for last/this/next week
- WCAG AA accessibility (skip link, ARIA, focus trap, live regions)
- Teams tab pinned at: same URL
- MS Form → Power Automate → SharePoint for team submissions
- Power Automate Monday reminder flow to Teams channel

## Pending
- SharePoint embed whitelist for bcgov.github.io (WAM email sent, awaiting response)
- Mac glitches reported by one user — under investigation

## Design
BC Gov design system — BC Sans font, `#013366` primary blue, RYG colour tokens as CSS variables
