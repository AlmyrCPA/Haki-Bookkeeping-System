# Haki — Books of Accounts

Philippine bookkeeping app (BIR-style books, VAT/withholding summaries, SLSPI/QAP/SAWT,
alphalist, DAT/PDF generators). Originally a single-file browser artifact; now a
Vite + React app backed by Firebase, deployable to Netlify.

## Architecture

| Piece | Where |
| --- | --- |
| Auth | Firebase Email/Password. `src/App.jsx` gates login vs. dashboard via `onAuthStateChanged`. |
| Data | Firestore. One document per client at `/clients/{clientId}` holding the full Haki `data` object. `/clientIndex/{clientId}` holds `{ name, tin }` and powers the switcher. |
| Persistence | `src/BookkeepingApp.jsx` — a single `onSnapshot` load/live-sync + a debounced `setDoc(..., { merge: true })` save. These replaced the two `window.storage` calls; nothing else in the ~6,500-line component changed. |
| The app | `src/BookkeepingApp.jsx` (the migrated original). Exports a few primitives (`makeInitialData`, `emptyCompany`, `PARTY_TYPES`, `EntryModalShell`, `LabeledField`, `Field`, `HakiLogo`) reused by the Phase 1 shell. |
| Phase 1 shell | `src/Login.jsx`, `src/ClientSwitcher.jsx` (sidebar dropdown, above the nav), `src/AddClientModal.jsx`. |

## Local development

Requires Node 18+ (Node 20 recommended — Netlify uses 20).

```bash
npm install
npm run dev
```

The production Firebase config is baked into `src/firebase.js`, so no `.env` is needed
to run against production. To point a local build at a different Firebase project, copy
`.env.example` to `.env` and fill in the `VITE_FIREBASE_*` vars.

## Firebase setup (one time)

1. In the [Firebase console](https://console.firebase.google.com/) for project
   `haki-bookeeping-system`: enable **Firestore** and **Authentication → Email/Password**.
2. Create the first user under Authentication → Users (this is the Phase 1 shared login).
3. Publish the Firestore rules:
   ```bash
   npm i -g firebase-tools
   firebase login
   firebase deploy --only firestore:rules
   ```
   (Rules are in `firestore.rules` — Phase 1: `allow read, write: if request.auth != null;`.)

## Netlify deployment

1. Push this repo to GitHub.
2. In Netlify: **Add new site → Import from Git**, pick the repo.
3. Build settings are read from `netlify.toml` (build `npm run build`, publish `dist`,
   SPA redirect to `index.html`). No env vars required.
4. Deploy. Open the live URL, sign in with the user from step 2, add a client, done.
5. In the Firebase console → Authentication → Settings → **Authorized domains**, add the
   Netlify domain (`your-site.netlify.app` and any custom domain) so login works.

## Migrating existing client backups

Each existing client has a JSON backup via **Company Details → Export full backup**.
A one-time script can loop those files and `setDoc` each into `/clients/{newClientId}`
plus a matching `/clientIndex/{newClientId}` entry — see the migration brief.

## Phase 2 (later)

Roles (Admin / Bookkeeper / Viewer) via custom claims + a Cloud Function, tightened
Firestore rules, and role-aware UI. Layers on top of Phase 1; see the migration brief.
