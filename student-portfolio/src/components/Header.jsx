function Header({ name, title, tagline, themeColor = '#2563eb' }) {
  // Theme color prop applied as inline style (Supplementary Problem #3)
  const headerStyle = {
    backgroundColor: themeColor,
    color: '#ffffff',
  };

  return (
    <header className="portfolio-header" style={headerStyle}>
      <div className="header-container">
        <span className="badge">Student Portfolio</span>
        <h1 className="header-name">{name}</h1>
        <h2 className="header-title">{title}</h2>
        {tagline && <p className="header-tagline">{tagline}</p>}
      </div>
    </header>
  );
}

export default Header;
