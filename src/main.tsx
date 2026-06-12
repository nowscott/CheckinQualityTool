import { createRoot } from "react-dom/client";
import App from "./App";
import { applyTheme, initialTheme } from "./hooks/useTheme";
import "./styles.css";

applyTheme(initialTheme());
createRoot(document.getElementById("root")!).render(<App />);
