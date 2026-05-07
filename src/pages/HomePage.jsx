import ActivityHub from "../components/ActivityHub.jsx";

function HomePage({ onNavigate }) {
  return (
    <main className="page-shell">
      <section className="home-hero">
        <div>
          <p className="eyebrow">Classroom Activity Hub</p>
          <h1>Student Practice Engine</h1>
          <p>Choose a unit, open a practice activity, and keep your progress on this device.</p>
        </div>
      </section>
      <ActivityHub onNavigate={onNavigate} />
    </main>
  );
}

export default HomePage;
