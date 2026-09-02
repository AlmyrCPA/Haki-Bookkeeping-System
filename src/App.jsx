import React, { useEffect, useState } from "react";
import { onAuthStateChanged, signOut } from "firebase/auth";
import { collection, onSnapshot, query, orderBy } from "firebase/firestore";
import { auth, db } from "./firebase";
import Login from "./Login";
import BookkeepingApp, { HakiLogo } from "./BookkeepingApp";
import ClientSwitcher from "./ClientSwitcher";
import AddClientModal from "./AddClientModal";

const LAST_CLIENT_KEY = "haki-last-client";

export default function App() {
  const [authState, setAuthState] = useState("checking"); // checking | out | in
  const [user, setUser] = useState(null);

  const [clients, setClients] = useState([]);
  const [clientsLoaded, setClientsLoaded] = useState(false);
  const [clientId, setClientId] = useState(() => {
    try { return localStorage.getItem(LAST_CLIENT_KEY) || null; } catch { return null; }
  });
  const [showAddClient, setShowAddClient] = useState(false);

  // Auth gate — onAuthStateChanged flips the app between login and dashboard.
  useEffect(() => {
    return onAuthStateChanged(auth, (u) => {
      setUser(u);
      setAuthState(u ? "in" : "out");
    });
  }, []);

  // Live list of clients for the switcher, from /clientIndex.
  useEffect(() => {
    if (authState !== "in") { setClients([]); setClientsLoaded(false); return; }
    const q = query(collection(db, "clientIndex"), orderBy("name"));
    const unsub = onSnapshot(
      q,
      (snap) => {
        setClients(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
        setClientsLoaded(true);
      },
      (err) => { console.error("Failed to load client index:", err); setClientsLoaded(true); }
    );
    return unsub;
  }, [authState]);

  // Keep the selected client valid and remembered.
  useEffect(() => {
    if (!clientsLoaded) return;
    if (clients.length === 0) { setClientId(null); return; }
    if (!clientId || !clients.some((c) => c.id === clientId)) {
      setClientId(clients[0].id);
    }
  }, [clients, clientsLoaded, clientId]);

  useEffect(() => {
    try {
      if (clientId) localStorage.setItem(LAST_CLIENT_KEY, clientId);
    } catch { /* ignore */ }
  }, [clientId]);

  const switchClient = (id) => setClientId(id || null);
  const handleClientCreated = (id) => {
    setShowAddClient(false);
    setClientId(id);
  };

  if (authState === "checking") {
    return <SplashScreen label="Loading…" />;
  }
  if (authState === "out") {
    return <Login />;
  }
  if (!clientsLoaded) {
    return <SplashScreen label="Loading clients…" onSignOut={() => signOut(auth)} />;
  }

  // Signed in but no clients exist yet — first-run screen.
  if (clients.length === 0) {
    return (
      <>
        <FirstClientScreen
          email={user?.email}
          onAddClient={() => setShowAddClient(true)}
          onSignOut={() => signOut(auth)}
        />
        {showAddClient && (
          <AddClientModal
            onCancel={() => setShowAddClient(false)}
            onCreate={handleClientCreated}
          />
        )}
      </>
    );
  }

  return (
    <>
      <BookkeepingApp
        clientId={clientId}
        onSignOut={() => signOut(auth)}
        clientSwitcher={
          <ClientSwitcher
            clients={clients}
            currentClientId={clientId}
            onSwitch={switchClient}
            onAddClient={() => setShowAddClient(true)}
          />
        }
      />
      {showAddClient && (
        <AddClientModal
          onCancel={() => setShowAddClient(false)}
          onCreate={handleClientCreated}
        />
      )}
    </>
  );
}

function SplashScreen({ label, onSignOut }) {
  return (
    <div className="haki-splash">
      <HakiLogo size={44} />
      <div className="haki-splash-label">{label}</div>
      {onSignOut && (
        <button className="haki-splash-signout" onClick={onSignOut}>Sign out</button>
      )}
      <SplashStyles />
    </div>
  );
}

function FirstClientScreen({ email, onAddClient, onSignOut }) {
  return (
    <div className="haki-splash">
      <HakiLogo size={44} />
      <div className="haki-splash-title">Welcome to Haki</div>
      <div className="haki-splash-label">
        You're signed in{email ? ` as ${email}` : ""}. Create your first client to get started.
      </div>
      <button className="haki-splash-primary" onClick={onAddClient}>Add your first client</button>
      <button className="haki-splash-signout" onClick={onSignOut}>Sign out</button>
      <SplashStyles />
    </div>
  );
}

function SplashStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600&family=IBM+Plex+Sans:wght@400;500;600&display=swap');
      .haki-splash {
        min-height: 100vh; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 14px;
        background: linear-gradient(180deg, #0F2A4D, #1D5FA8); color: #E8F1FC; padding: 24px; text-align: center;
        font-family: 'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif;
      }
      .haki-splash svg { border-radius: 10px; }
      .haki-splash-title { font-family: 'Fraunces', serif; font-size: 24px; font-weight: 600; color: #fff; }
      .haki-splash-label { font-size: 14px; color: #C6DCF3; max-width: 380px; line-height: 1.5; }
      .haki-splash-primary {
        margin-top: 6px; background: #fff; color: #0F2A4D; border: none; border-radius: 8px; padding: 11px 20px;
        font-size: 14px; font-weight: 600; font-family: inherit; cursor: pointer;
      }
      .haki-splash-primary:hover { background: #EAF2FC; }
      .haki-splash-signout {
        background: none; border: 1px solid rgba(255,255,255,0.35); color: #DCEBFA; border-radius: 7px;
        padding: 7px 14px; font-size: 12.5px; font-family: inherit; cursor: pointer;
      }
      .haki-splash-signout:hover { background: rgba(255,255,255,0.10); }
    `}</style>
  );
}
