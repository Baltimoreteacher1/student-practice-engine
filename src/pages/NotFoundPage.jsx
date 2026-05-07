function NotFoundPage({ onNavigate }) {
  return (
    <main className="page-shell">
      <section className="empty-state">
        <p className="eyebrow">Page Not Found</p>
        <h1>This activity link is not ready yet</h1>
        <p>Check the link or go back to the activity hub.</p>
        <button className="button" onClick={() => onNavigate("/")} type="button">
          Back Home
        </button>
      </section>
    </main>
  );
}

export default NotFoundPage;
