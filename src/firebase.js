// Firebase bootstrap — one app instance shared across the whole client.
// Config values are safe to ship in the bundle (they identify the project, they
// don't grant access); Firestore security rules are what actually gate data.
// They're read from Vite env vars when present so Netlify can point a preview
// build at a different project, falling back to the committed defaults.
import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { initializeFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY || "AIzaSyB64Kh-STrlIylTf4ajesA90k3l7vZxH4Q",
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || "haki-bookeeping-system.firebaseapp.com",
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || "haki-bookeeping-system",
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || "haki-bookeeping-system.firebasestorage.app",
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || "722514411806",
  appId: import.meta.env.VITE_FIREBASE_APP_ID || "1:722514411806:web:fbe07ceaf33d9b81f3090f",
};

export const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);

// ignoreUndefinedProperties: Haki row objects occasionally carry `undefined`
// fields; without this a single one would make the whole setDoc save throw.
export const db = initializeFirestore(app, { ignoreUndefinedProperties: true });
