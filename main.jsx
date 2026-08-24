import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import "./supabaseClient"; // sets up window.storage before anything renders
import App from "./App";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
