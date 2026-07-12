import React from "react";
import { createRoot } from "react-dom/client";
import App from "./peres-daily-dashboard.jsx";

/*
 * הדשבורד משתמש ב-window.storage (ה-API של ארטיפקטים בקלוד) לשמירת
 * תמונות מצב, דלתות וסטטוס קמפיינים. מחוץ לקלוד (ורסל) ה-API הזה לא קיים,
 * אז מספקים מימוש זהה מעל localStorage — הכל ממשיך לעבוד אותו דבר.
 */
if (!window.storage) {
  const PREFIX = "peres-storage:";
  window.storage = {
    async get(key) {
      const value = localStorage.getItem(PREFIX + key);
      if (value === null) throw new Error("Key not found: " + key);
      return { key, value, shared: false };
    },
    async set(key, value) {
      localStorage.setItem(PREFIX + key, value);
      return { key, value, shared: false };
    },
    async delete(key) {
      localStorage.removeItem(PREFIX + key);
      return { key, deleted: true, shared: false };
    },
    async list(prefix = "") {
      const keys = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith(PREFIX + prefix)) keys.push(k.slice(PREFIX.length));
      }
      return { keys, prefix, shared: false };
    },
  };
}

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
