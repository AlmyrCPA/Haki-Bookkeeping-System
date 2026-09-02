import React, { useState } from "react";
import { signInWithEmailAndPassword } from "firebase/auth";
import { auth } from "./firebase";
import { HakiLogo } from "./BookkeepingApp";

// Phase 1: one shared login, no roles. onAuthStateChanged in App.jsx flips the
// app from "show login" to "show dashboard" once this succeeds.
export default function Login() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function handleLogin(e) {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      await signInWithEmailAndPassword(auth, email.trim(), password);
    } catch (err) {
      setError(messageFor(err));
      setBusy(false);
    }
  }

  return (
    <div className="haki-login">
      <form className="haki-login-card" onSubmit={handleLogin}>
        <div className="haki-login-brand">
          <HakiLogo size={40} />
          <div>
            <div className="haki-login-title">Haki</div>
            <div className="haki-login-sub">Books of Accounts</div>
          </div>
        </div>

        <label className="haki-login-field">
          <span>Email</span>
          <input
            type="email"
            autoComplete="username"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
          />
        </label>

        <label className="haki-login-field">
          <span>Password</span>
          <input
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
          />
        </label>

        {error && <div className="haki-login-error">{error}</div>}

        <button className="haki-login-btn" type="submit" disabled={busy}>
          {busy ? "Signing in…" : "Sign in"}
        </button>
      </form>

      <Styles />
    </div>
  );
}

function messageFor(err) {
  switch (err?.code) {
    case "auth/invalid-email":
      return "That doesn't look like a valid email address.";
    case "auth/user-disabled":
      return "This account has been disabled.";
    case "auth/user-not-found":
    case "auth/wrong-password":
    case "auth/invalid-credential":
      return "Wrong email or password.";
    case "auth/too-many-requests":
      return "Too many attempts. Wait a moment and try again.";
    case "auth/network-request-failed":
      return "Network error. Check your connection.";
    default:
      return "Couldn't sign in. Please try again.";
  }
}

function Styles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600&display=swap');
      .haki-login {
        min-height: 100vh; display: flex; flex-direction: column; align-items: center; justify-content: center;
        background: linear-gradient(180deg, #0F2A4D, #1D5FA8); padding: 24px;
        font-family: 'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif;
      }
      .haki-login-card {
        width: 100%; max-width: 360px; background: #FFFFFF; border-radius: 14px;
        padding: 28px 26px 26px; box-shadow: 0 24px 60px rgba(0,0,0,0.28); display: flex; flex-direction: column; gap: 16px;
      }
      .haki-login-brand { display: flex; align-items: center; gap: 12px; margin-bottom: 4px; }
      .haki-login-brand svg { border-radius: 9px; }
      .haki-login-title { font-family: 'Fraunces', serif; font-weight: 600; font-size: 20px; color: #0F2A4D; }
      .haki-login-sub { font-size: 12px; color: #55697F; margin-top: 1px; }
      .haki-login-field { display: flex; flex-direction: column; gap: 5px; font-size: 12.5px; color: #16233A; font-weight: 500; }
      .haki-login-field input {
        border: 1px solid #C7DAF0; border-radius: 8px; padding: 10px 11px; font-size: 14px; font-family: inherit; color: #16233A;
      }
      .haki-login-field input:focus { outline: none; border-color: #1D5FA8; box-shadow: 0 0 0 3px rgba(29,95,168,0.15); }
      .haki-login-error {
        font-size: 12.5px; color: #A13D3D; background: #FBECEC; border: 1px solid #E7C9C9; border-radius: 7px; padding: 8px 10px;
      }
      .haki-login-btn {
        margin-top: 4px; background: #1D5FA8; color: #FFFFFF; border: none; border-radius: 8px; padding: 11px 14px;
        font-size: 14px; font-weight: 600; font-family: inherit; cursor: pointer;
      }
      .haki-login-btn:hover:not(:disabled) { background: #17508f; }
      .haki-login-btn:disabled { opacity: 0.6; cursor: default; }
    `}</style>
  );
}
