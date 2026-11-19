# WVS Field Hub

The rep-facing hub for WV Squared — priorities, brand files, store lists, order
forms, promotions, 2027 call cycles, activations and marketing spend.

Hosted on GitHub Pages, with its data in Firebase Firestore.

## How access works

- **Reps** open the link and enter their 4-digit code. The code decides whose
  route they see. It is a soft lock: it stops reps clicking into each other's
  views, it is not real security.
- **Managers** click *Manager? Sign in with your email* and use a Firebase
  email/password account. Signing in that way unlocks the Overview and
  Marketing tabs and the right to edit anything.
- Everyone else is an anonymous Firebase visitor, which is read-only.

## Setup

1. Paste the project's web config into `firebase-config.js` (Firebase console →
   Project settings → General → Your apps → Web app → Config). These values are
   not secrets; `firestore.rules` is what controls access.
2. In the Firebase console → Authentication → Sign-in method, enable
   **Anonymous** and **Email/Password**, and add a manager account under Users.
3. Paste `firestore.rules` into Firebase console → Firestore Database → Rules →
   Publish.
4. Open the hub, sign in as a manager, go to `#import` and run the one-time
   import. That writes `data/bundle.json` into Firestore.

After the import, `data/bundle.json` is only a snapshot of the starting data —
the live data lives in Firestore.
