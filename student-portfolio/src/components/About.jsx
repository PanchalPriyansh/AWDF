function About({ bio, education, interests }) {
  return (
    <section id="about" className="portfolio-section about-section">
      <div className="section-title-wrap">
        <h2 className="section-heading">About Me</h2>
        <div className="section-underline"></div>
      </div>
      <div className="about-content">
        <p className="about-bio">{bio}</p>
        <div className="about-info-cards">
          {education && (
            <div className="info-card">
              <div className="card-icon">🎓</div>
              <div>
                <h3 className="card-title">Education</h3>
                <p className="card-desc">{education}</p>
              </div>
            </div>
          )}
          {interests && (
            <div className="info-card">
              <div className="card-icon">🚀</div>
              <div>
                <h3 className="card-title">Areas of Interest</h3>
                <p className="card-desc">{interests}</p>
              </div>
            </div>
          )}
        </div>
      </div>
    </section>
  );
}

export default About;
