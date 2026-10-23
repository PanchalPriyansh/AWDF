function Skills({ skillList }) {
  return (
    <section id="skills" className="portfolio-section skills-section">
      <div className="section-title-wrap">
        <h2 className="section-heading">Skills</h2>
        <div className="section-underline"></div>
      </div>
      <p className="skills-subheading">Core competencies & technologies</p>
      <ul className="skills-list">
        {skillList.map((s) => (
          <li key={s} className="skill-item">
            <span className="skill-dot"></span>
            <span className="skill-name">{s}</span>
          </li>
        ))}
      </ul>
    </section>
  );
}

export default Skills;
