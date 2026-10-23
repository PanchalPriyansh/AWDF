import { useState, useCallback, lazy, Suspense } from 'react';
import { Routes, Route } from 'react-router-dom';
import NavBar from './components/NavBar';
import Footer from './components/Footer';
import DelayedFallback from './components/DelayedFallback';
import './App.css';

// Route-based code splitting using React.lazy() (Practical 8 - Steps 3 & 4)
const Home = lazy(() => import('./pages/Home'));
const Projects = lazy(() => import('./pages/Projects'));
const Contact = lazy(() => import('./pages/Contact'));
const NotFound = lazy(() => import('./pages/NotFound'));

function App() {
  // Dark/Light Mode reactive state
  const [darkMode, setDarkMode] = useState(false);

  // Student portfolio data passed to reusable components via props
  const student = {
    name: 'Priyansh Pravinbhai Panchal',
    title: 'Computer Science Student at CHARUSAT University',
    tagline: 'Passionate about full-stack web development, software engineering, and building impactful digital solutions.',
    themeColor: '#1d4ed8', // Royal Blue
    bio: 'I am a Computer Science student with a strong interest in software development, programming, and emerging technologies. I have experience working with languages and tools such as Python, C, Flask, MongoDB, and Git, and I enjoy applying my technical knowledge to practical projects. My goal is to continuously improve my problem-solving and development skills, gain hands-on industry experience, and build innovative solutions that solve real-world problems.',
    education: 'B.Tech in Computer Science & Engineering (Semester 5), CHARUSAT University',
    interests: 'MERN Stack, Python & Flask Development, Backend Systems, and Emerging Tech',
    skills: [
      'MERN Stack (MongoDB, Express, React, Node.js)',
      'React.js & Vite Tooling',
      'Python & Flask',
      'C & C++ Programming',
      'HTML5 & Modern CSS3',
      'JavaScript (ES6+)',
      'Git & GitHub Version Control',
      'REST APIs & Database Management'
    ],
    email: 'panchalpriyansh3257@gmail.com',
    github: 'PanchalPriyansh'
  };

  // useCallback prevents re-creating toggle handler, avoiding unnecessary re-renders (Supplementary 3)
  const toggleTheme = useCallback(() => {
    setDarkMode((prev) => !prev);
  }, []);

  return (
    <div className={`portfolio-app ${darkMode ? 'dark-theme' : 'light-theme'}`}>
      <NavBar
        brand={student.name}
        darkMode={darkMode}
        onToggleTheme={toggleTheme}
      />

      <div className="main-content-wrapper">
        {/* Suspense wrapper with 300ms minimum-delay fallback (Step 4 & Supplementary 2) */}
        <Suspense fallback={<DelayedFallback delay={300} label="Loading page resources..." />}>
          <Routes>
            {/* Lazy Route: "/" -> Home.jsx */}
            <Route path="/" element={<Home student={student} />} />

            {/* Lazy Route: "/projects" -> Projects.jsx (Full-Stack Task Manager) */}
            <Route path="/projects" element={<Projects username={student.github} />} />

            {/* Lazy Route: "/contact" -> Contact.jsx */}
            <Route
              path="/contact"
              element={
                <Contact
                  studentEmail={student.email}
                  studentGithub={student.github}
                />
              }
            />

            {/* Lazy Route: "*" -> NotFound.jsx */}
            <Route path="*" element={<NotFound />} />
          </Routes>
        </Suspense>
      </div>

      <Footer
        name={student.name}
        year={2026}
        email={student.email}
        github={student.github}
      />
    </div>
  );
}

export default App;
