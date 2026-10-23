# Advance Web Development Frameworks - Lab Practicals

[![Task Management CI/CD Pipeline](https://github.com/PanchalPriyansh/AWDF/actions/workflows/ci.yml/badge.svg)](https://github.com/PanchalPriyansh/AWDF/actions/workflows/ci.yml)

**Student Name:** Priyansh Pravinbhai Panchal  
**Institution:** CHARUSAT University  
**Department:** Department of Computer Science & Engineering (B.Tech CSE, Semester 5)  
**Academic Year:** 2026–2027  

---

## Project Overview

This repository contains the coursework, code implementations, Docker orchestration, and continuous integration workflows for **Advance Web Development Frameworks**. The project is an enterprise-grade full-stack **Task Management Application** built using the MERN stack (MongoDB, Express.js, React, Node.js).

### Architecture Highlights
- **Frontend (`student-portfolio/`):** React 19, Vite, React Router 7, Chart.js analytics dashboard, and modern responsive UI.
- **Backend (`task-manager-api/`):** Node.js, Express.js, MongoDB via Mongoose, JWT authentication, bcrypt password hashing, structured input validation, node-cache in-memory caching (TTL 60s), and native Node.js `EventEmitter` asynchronous event processing.
- **Containerization (`docker-compose.yml`):** Multi-stage Dockerized frontend (Nginx/Node static server), Alpine Node backend, and persistent MongoDB service on a bridge network.
- **Continuous Integration (`.github/workflows/ci.yml`):** Multi-job GitHub Actions pipeline running code linting with Oxlint and automated integration tests with Jest, Supertest, and a containerized MongoDB service.

---

## Practical Progress & Index

| Practical No. | Topic / Focus Area | Status |
|:---:|:---|:---:|
| 1 | Personal Portfolio Frontend & Routing (React + Vite) | Completed |
| 2 | Advanced Component State Management & Interactive Widgets | Completed |
| 3 | Dynamic Task Filtering, Search & Priority State | Completed |
| 4 | Express RESTful API Foundations & Task Routing | Completed |
| 5 | MongoDB Atlas & Local Connection with Mongoose Schemas | Completed |
| 6 | Full CRUD Endpoints with Request Parameter Handling | Completed |
| 7 | JWT User Authentication, Password Hashing & Auth Middleware | Completed |
| 8 | Input Validation Middleware & Comprehensive Error Handling | Completed |
| 9 | Server-Side Caching with Invalidation using `node-cache` | Completed |
| 10 | Asynchronous Background Processing with Node.js `EventEmitter` | Completed |
| 11 | Full-Stack Multi-Stage Docker Containerization & Compose | Completed |
| 12 | CI/CD Pipeline Automation with GitHub Actions & Jest | Completed |
| 13 | Production Cloud Deployment & Monitoring | Pending |

---

## CI/CD Pipeline Architecture (Practical 12)

The GitHub Actions workflow is defined in [`.github/workflows/ci.yml`](.github/workflows/ci.yml) and consists of two sequential jobs:

```
[Push to main / PR to main]
            │
            ▼
┌───────────────────────────────┐
│     Job 1: Lint (Oxlint)      │ ──► Enforces code standards & catches syntax issues
└──────────────┬────────────────┘
               │ (needs: lint)
               ▼
┌───────────────────────────────┐
│     Job 2: Test (Jest)        │
│  - MongoDB 7.0 Service        │ ──► Runs 7 automated API integration tests
│  - Supertest HTTP assertions  │     ensuring zero regression across endpoints
└───────────────────────────────┘
```

### Running Tests Locally

```bash
cd task-manager-api
npm test
```

### Running Lint Locally

```bash
cd student-portfolio
npm run lint
```
