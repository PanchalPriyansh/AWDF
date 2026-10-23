import { Link } from 'react-router-dom';
import Header from '../components/Header';
import About from '../components/About';
import Skills from '../components/Skills';

function Home({ student }) {
  return (
    <div className="page-home">
      <Header
        name={student.name}
        title={student.title}
        tagline={student.tagline}
        themeColor={student.themeColor}
      />
      <div className="home-actions-bar">
        <Link to="/projects" className="btn-primary">
          Task Manager & Projects 🚀
        </Link>
        <Link to="/contact" className="btn-secondary">
          Get in Touch ✉️
        </Link>
      </div>
      <main className="portfolio-content">
        <About
          bio={student.bio}
          education={student.education}
          interests={student.interests}
        />
        <Skills skillList={student.skills} />
      </main>
    </div>
  );
}

export default Home;
