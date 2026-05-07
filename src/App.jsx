import { useEffect, useState } from "react";
import ActivityPage from "./pages/ActivityPage.jsx";
import HomePage from "./pages/HomePage.jsx";
import NotFoundPage from "./pages/NotFoundPage.jsx";
import UnitPage from "./pages/UnitPage.jsx";
import { parsePath } from "./utils/routing.js";
import "./styles/app.css";

function App() {
  const [route, setRoute] = useState(() => parsePath(window.location.pathname));

  useEffect(() => {
    const onPopState = () => setRoute(parsePath(window.location.pathname));
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  function navigate(path) {
    window.history.pushState({}, "", path);
    setRoute(parsePath(path));
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  if (route.page === "unit") return <UnitPage unitSlug={route.unitSlug} onNavigate={navigate} />;
  if (route.page === "activity") return <ActivityPage activitySlug={route.activitySlug} onNavigate={navigate} />;
  if (route.page === "notFound") return <NotFoundPage onNavigate={navigate} />;

  return <HomePage onNavigate={navigate} />;
}

export default App;
