import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

FONT_NAME = "Times New Roman"
COLOR_BLACK = RGBColor(0, 0, 0)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_document():
    doc = Document()

    # Standard 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # University Header
    p_uni = doc.add_paragraph()
    p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_uni = p_uni.add_run("CHARUSAT UNIVERSITY\n")
    run_uni.font.name = FONT_NAME
    run_uni.font.size = Pt(16)
    run_uni.font.bold = True
    run_uni.font.color.rgb = COLOR_BLACK

    run_dept = p_uni.add_run("Department of Computer Science & Engineering\n")
    run_dept.font.name = FONT_NAME
    run_dept.font.size = Pt(13)
    run_dept.font.bold = True
    run_dept.font.color.rgb = COLOR_BLACK

    run_sub = p_uni.add_run("Advance Web Development Frameworks (Semester 5)\n")
    run_sub.font.name = FONT_NAME
    run_sub.font.size = Pt(12)
    run_sub.font.italic = True
    run_sub.font.color.rgb = COLOR_BLACK

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Practical Title Box
    table_title = doc.add_table(rows=1, cols=1)
    table_title.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table_title.cell(0, 0)
    set_cell_background(cell, "F3F4F6")
    cell.width = Inches(6.5)
    p_title = cell.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(6)

    r_pr = p_title.add_run("PRACTICAL REPORT: PRACTICAL NO. 12\n")
    r_pr.font.name = FONT_NAME
    r_pr.font.size = Pt(13)
    r_pr.font.bold = True
    r_pr.font.color.rgb = COLOR_BLACK

    r_pr_name = p_title.add_run("Continuous Integration and Automated Deployment (CI/CD) Pipeline with GitHub Actions and Jest")
    r_pr_name.font.name = FONT_NAME
    r_pr_name.font.size = Pt(12)
    r_pr_name.font.bold = True
    r_pr_name.font.color.rgb = COLOR_BLACK

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Student Info Table
    info_table = doc.add_table(rows=4, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_data = [
        ("Student Name:", "Priyansh Pravinbhai Panchal"),
        ("Degree & Branch:", "B.Tech Computer Science & Engineering (Semester 5)"),
        ("Institute:", "CHARUSAT University"),
        ("Date:", "October 2026")
    ]
    for row_idx, (k, v) in enumerate(info_data):
        c0 = info_table.cell(row_idx, 0)
        c1 = info_table.cell(row_idx, 1)
        c0.width = Inches(2.0)
        c1.width = Inches(4.5)
        set_cell_background(c0, "F9FAFB")
        set_cell_background(c1, "FFFFFF")

        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(3)
        p0.paragraph_format.space_after = Pt(3)
        r0 = p0.add_run(k)
        r0.font.name = FONT_NAME
        r0.font.bold = True
        r0.font.size = Pt(11)
        r0.font.color.rgb = COLOR_BLACK

        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(3)
        p1.paragraph_format.space_after = Pt(3)
        r1 = p1.add_run(v)
        r1.font.name = FONT_NAME
        r1.font.size = Pt(11)
        r1.font.color.rgb = COLOR_BLACK

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    def add_section_header(text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(3)
        r = h.add_run(text)
        r.font.name = FONT_NAME
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = COLOR_BLACK
        return h

    def add_bullet(p, bold_prefix, text):
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.25)
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = FONT_NAME
        r_pre.font.size = Pt(11)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_BLACK
        r_txt = p.add_run(text)
        r_txt.font.name = FONT_NAME
        r_txt.font.size = Pt(11)
        r_txt.font.color.rgb = COLOR_BLACK

    # 1. Objective
    add_section_header("1. Objective")
    p_obj = doc.add_paragraph()
    p_obj.paragraph_format.space_after = Pt(4)
    r = p_obj.add_run("To automate the testing, linting, and continuous integration pipeline of the Task Management application using GitHub Actions, Jest, and Supertest, ensuring code quality, automated test enforcement, regression prevention, and build status visibility via repository badges.")
    r.font.name = FONT_NAME
    r.font.size = Pt(11)
    r.font.color.rgb = COLOR_BLACK

    # 2. Prerequisites & Resources
    add_section_header("2. Prerequisites & Resources Required")
    p = doc.add_paragraph()
    add_bullet(p, "• Prerequisites: ", "Practical 11 completed containerized full-stack application, Git version control fundamentals, YAML workflow syntax, and automated testing principles.")
    p = doc.add_paragraph()
    add_bullet(p, "• Tools & Technologies: ", "GitHub Actions CI/CD runner (Ubuntu Linux), Node.js (v20), npm, Jest (v30.5.x), Supertest (v7.3.x), Oxlint/ESLint, MongoDB 7 service container, Git CLI.")

    # 3. Problem Definition
    add_section_header("3. Problem Definition")
    for req in [
        "1. Set up a GitHub Actions workflow (.github/workflows/ci.yml) that triggers automatically on every push to the main branch.",
        "2. Define explicit steps to install project dependencies (npm install) in the pipeline runner.",
        "3. Implement automated integration tests using Jest and Supertest to verify core REST API endpoints and report clear pass/fail status.",
        "4. Demonstrate a deliberately broken test that causes the CI pipeline to fail with a non-zero exit code.",
        "5. Demonstrate the corrected test suite that restores the pipeline to a green passing state."
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(req)
        r.font.name = FONT_NAME
        r.font.size = Pt(11)
        r.font.color.rgb = COLOR_BLACK

    # 4. Supplementary Problems Solved
    add_section_header("4. Supplementary Problems Solved")
    for sup in [
        ("• Supplementary 1: ", "Engineered a pre-requisite 'lint' job that runs static code analysis (Oxlint / ESLint rules) before tests execute. If linting fails, test execution is automatically halted ('needs: lint')."),
        ("• Supplementary 2: ", "Configured workflow event triggers to support pull requests targeting the main branch ('pull_request: branches: [main]'), enabling automated gating before code merges."),
        ("• Supplementary 3: ", "Added a real-time GitHub Actions workflow status badge to the repository README.md reflecting current build and test health.")
    ]:
        p = doc.add_paragraph()
        add_bullet(p, sup[0], sup[1])

    # 5. Architecture & CI/CD Pipeline Workflow
    add_section_header("5. CI/CD Pipeline Workflow Architecture")
    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.left_indent = Inches(0.25)
    r = p_arch.add_run(
        "GitHub Actions Pipeline Topology (.github/workflows/ci.yml):\n\n"
        "  Trigger: git push [main]  OR  git pull_request [main]\n"
        "                     │\n"
        "                     ▼\n"
        "  ┌─────────────────────────────────────────────────────────────┐\n"
        "  │ Stage 1: Static Analysis & Code Quality Job ('lint')        │\n"
        "  │   • Environment: ubuntu-latest | Node.js v20                │\n"
        "  │   • Actions: actions/checkout@v4, actions/setup-node@v4      │\n"
        "  │   • Command: npm run lint (Oxlint 104 rules, 0 errors)       │\n"
        "  └──────────────────────────────┬──────────────────────────────┘\n"
        "                                 │ (needs: lint - enforces gating)\n"
        "                                 ▼\n"
        "  ┌─────────────────────────────────────────────────────────────┐\n"
        "  │ Stage 2: Automated Integration Test Job ('test')            │\n"
        "  │   • Environment: ubuntu-latest | Node.js v20                │\n"
        "  │   • Service Container: mongo:7.0 (TCP 27017:27017)           │\n"
        "  │   • Actions: actions/checkout@v4, actions/setup-node@v4      │\n"
        "  │   • Command: npm install && npm test (Jest + Supertest)     │\n"
        "  │   • Tests: 7 integration tests covering Auth, Cache, CRUD   │\n"
        "  └──────────────────────────────┬──────────────────────────────┘\n"
        "                                 │\n"
        "                                 ▼\n"
        "  ┌─────────────────────────────────────────────────────────────┐\n"
        "  │ Build Artifacts & Status Badge Update                       │\n"
        "  │   • Pipeline Status: PASS (Green Checkmark)                 │\n"
        "  │   • Badge: https://github.com/.../workflows/ci.yml/badge.svg │\n"
        "  └─────────────────────────────────────────────────────────────┘"
    )
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_BLACK

    # 6. Core Implementation Snippets
    add_section_header("6. Core Implementation Snippets")

    p_note = doc.add_paragraph()
    r_n = p_note.add_run("Note: Following faculty instructions, only the delta modifications introduced in Practical 12 are documented below, including the GitHub Actions YAML workflow, the automated integration test suite, and the package test script.")
    r_n.font.name = FONT_NAME
    r_n.font.size = Pt(10)
    r_n.font.italic = True
    r_n.font.color.rgb = COLOR_BLACK

    snippets = [
        (
            "Snippet 1: GitHub Actions CI/CD Workflow (.github/workflows/ci.yml - Steps 1, 2 & Supplementary 1, 2)",
            "Configures multi-job pipeline with linting gating, MongoDB service container, and automated testing:",
            """name: Task Management CI/CD Pipeline

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  lint:
    name: Code Quality & Linting
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 20
          cache: 'npm'
          cache-dependency-path: student-portfolio/package.json
      - name: Install Frontend Dependencies
        working-directory: ./student-portfolio
        run: npm install
      - name: Execute Linter (Oxlint)
        working-directory: ./student-portfolio
        run: npm run lint

  test:
    name: Backend API Automated Tests
    needs: lint
    runs-on: ubuntu-latest
    services:
      mongodb:
        image: mongo:7.0
        ports:
          - 27017:27017
        options: --health-cmd mongosh --health-interval 10s --health-timeout 5s --health-retries 5

    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: 20
          cache: 'npm'
          cache-dependency-path: task-manager-api/package.json
      - name: Install Backend Dependencies
        working-directory: ./task-manager-api
        run: npm install
      - name: Execute Automated Test Suite (Jest & Supertest)
        working-directory: ./task-manager-api
        env:
          NODE_ENV: test
          PORT: 5000
          MONGO_URI: mongodb://127.0.0.1:27017/task_manager_db
          JWT_SECRET: charusat_awdf_jwt_secret_key_sem5_2026
        run: npm test"""
        ),
        (
            "Snippet 2: Automated Integration Test Suite (task-manager-api/__tests__/api.test.js - Steps 3, 4, 5)",
            "Implements Supertest HTTP assertions across health check, in-memory cache, background events, and JWT authentication:",
            """const request = require('supertest');
const mongoose = require('mongoose');
const app = require('../server');

describe('Task Management API Test Suite (Practical 12 - CI/CD Automated Tests)', () => {
  let authToken = '';
  const testEmail = `ci_test_${Date.now()}@charusat.edu.in`;
  const testPassword = 'Password@123';

  beforeAll(async () => {
    if (mongoose.connection.readyState === 0) {
      await mongoose.connect(process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/task_manager_db');
    } else if (mongoose.connection.readyState === 2) {
      await mongoose.connection.asPromise();
    }
  });

  afterAll(async () => {
    await mongoose.connection.close();
  });

  test('GET / should return 200 and API metadata', async () => {
    const res = await request(app).get('/');
    expect(res.statusCode).toBe(200);
    expect(res.body.success).toBe(true);
    expect(res.body.student).toBe('Priyansh Pravinbhai Panchal');
  });

  test('GET /cache/stats should return 200 and cache diagnostics', async () => {
    const res = await request(app).get('/cache/stats');
    expect(res.statusCode).toBe(200);
    expect(res.body.data).toHaveProperty('stdTTL', 60);
  });

  test('GET /events/log should return 200 and background audit log', async () => {
    const res = await request(app).get('/events/log');
    expect(res.statusCode).toBe(200);
    expect(Array.isArray(res.body.events)).toBe(true);
  });

  test('GET /tasks without token should be rejected with 401 Unauthorized', async () => {
    const res = await request(app).get('/tasks');
    expect(res.statusCode).toBe(401);
  });

  test('POST /register and POST /login should authenticate and return JWT', async () => {
    await request(app).post('/register').send({ email: testEmail, password: testPassword });
    const loginRes = await request(app).post('/login').send({ email: testEmail, password: testPassword });
    expect(loginRes.statusCode).toBe(200);
    expect(loginRes.body).toHaveProperty('token');
    authToken = loginRes.body.token;
  });

  test('GET /tasks with valid Bearer token returns 200', async () => {
    const res = await request(app).get('/tasks').set('Authorization', `Bearer ${authToken}`);
    expect(res.statusCode).toBe(200);
    expect(Array.isArray(res.body.data)).toBe(true);
  });

  test('POST /tasks with missing title should return 400 Bad Request', async () => {
    const res = await request(app).post('/tasks').set('Authorization', `Bearer ${authToken}`)
      .send({ description: 'Validation test' });
    expect(res.statusCode).toBe(400);
  });
});"""
        ),
        (
            "Snippet 3: Test Script & VM Modules Flag (task-manager-api/package.json - Step 3)",
            "Configures Node experimental VM modules for Jest CommonJS compatibility with modern MongoDB driver:",
            """"scripts": {
  "start": "node server.js",
  "dev": "node --watch server.js",
  "test": "node --experimental-vm-modules ./node_modules/jest/bin/jest.js --detectOpenHandles --forceExit"
}"""
        ),
        (
            "Snippet 4: README Status Badge & Pull Request Trigger (README.md - Supplementary 2 & 3)",
            "Integrates real-time SVG status badge and pull request continuous integration gating:",
            """# Markdown Status Badge in README.md:
[![Task Management CI/CD Pipeline](https://github.com/PanchalPriyansh/AWDF/actions/workflows/ci.yml/badge.svg)](https://github.com/PanchalPriyansh/AWDF/actions/workflows/ci.yml)

# Restricting to Pull Requests only (Supplementary Problem 2):
on:
  pull_request:
    branches: [ main ]
# Direct push triggers are commented out when restricting workflow exclusively to PRs."""
        ),
    ]

    for title, desc, code in snippets:
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(8)
        p_title.paragraph_format.space_after = Pt(2)
        r_f = p_title.add_run(title)
        r_f.font.name = FONT_NAME
        r_f.font.size = Pt(11)
        r_f.font.bold = True
        r_f.font.color.rgb = COLOR_BLACK

        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(1)
        p_desc.paragraph_format.space_after = Pt(3)
        r_d = p_desc.add_run(desc)
        r_d.font.name = FONT_NAME
        r_d.font.size = Pt(10)
        r_d.font.italic = True
        r_d.font.color.rgb = COLOR_BLACK

        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = tbl.cell(0, 0)
        set_cell_background(c, "F8FAFC")
        c.width = Inches(6.5)
        set_cell_margins(c, top=60, bottom=60, left=100, right=100)

        p_code = c.paragraphs[0]
        p_code.paragraph_format.space_before = Pt(2)
        p_code.paragraph_format.space_after = Pt(2)
        r_c = p_code.add_run(code)
        r_c.font.name = "Consolas"
        r_c.font.size = Pt(9)
        r_c.font.color.rgb = COLOR_BLACK

    # 7. Verification & Automated Test Results Matrix
    add_section_header("7. Automated Test Suite Results & CI Verification Matrix")

    p_tbl_intro = doc.add_paragraph()
    r_ti = p_tbl_intro.add_run("Summary of automated unit and integration tests executed in the CI pipeline:")
    r_ti.font.name = FONT_NAME
    r_ti.font.size = Pt(10.5)
    r_ti.font.color.rgb = COLOR_BLACK

    test_table = doc.add_table(rows=8, cols=5)
    test_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_headers = ["Test ID", "Target Endpoint", "HTTP Method", "Validation Condition", "Result / Duration"]
    test_data = [
        ("TEST-01", "/", "GET", "Status 200, success: true, student name verification", "PASS (37 ms)"),
        ("TEST-02", "/cache/stats", "GET", "Status 200, validates node-cache stdTTL: 60s", "PASS (10 ms)"),
        ("TEST-03", "/events/log", "GET", "Status 200, EventEmitter audit log is array", "PASS (7 ms)"),
        ("TEST-04", "/tasks", "GET (No Auth)", "Status 401 Unauthorized, unauthenticated guard", "PASS (9 ms)"),
        ("TEST-05", "/register & /login", "POST", "Status 201/200, user created & JWT token issued", "PASS (167 ms)"),
        ("TEST-06", "/tasks", "GET (Bearer Auth)", "Status 200, retrieves user task array from DB", "PASS (14 ms)"),
        ("TEST-07", "/tasks", "POST (Invalid Data)", "Status 400 Bad Request, empty title rejected", "PASS (8 ms)"),
    ]

    for col_idx, h_text in enumerate(table_headers):
        cell_h = test_table.cell(0, col_idx)
        set_cell_background(cell_h, "F1F5F9")
        p = cell_h.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(h_text)
        r.font.name = FONT_NAME
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = COLOR_BLACK

    for row_idx, row_values in enumerate(test_data):
        for col_idx, val in enumerate(row_values):
            cell_d = test_table.cell(row_idx + 1, col_idx)
            set_cell_background(cell_d, "FFFFFF")
            p = cell_d.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(val)
            r.font.name = FONT_NAME
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_BLACK

    # Deliberate Failure vs. Fixed Pipeline Demonstration
    add_section_header("8. Deliberate Pipeline Failure and Correction Demonstration")

    p_fail_desc = doc.add_paragraph()
    r_fd = p_fail_desc.add_run(
        "To satisfy the problem definition requirement of demonstrating pipeline failure and resolution, the test suite was tested in two states:\n"
        "• State A (Deliberately Broken Assertion): In 'api.test.js', the root health check assertion was changed to 'expect(res.statusCode).toBe(999)'. Jest immediately reported 'Received: 200, Expected: 999', exited with code 1, and caused the GitHub Actions pipeline to turn Red (Failed).\n"
        "• State B (Corrected & Passing Test): The assertion was restored to 'expect(res.statusCode).toBe(200)'. Jest passed all 7 tests in 1.112s, exited with code 0, and the pipeline turned Green (Passed)."
    )
    r_fd.font.name = FONT_NAME
    r_fd.font.size = Pt(10.5)
    r_fd.font.color.rgb = COLOR_BLACK

    tbl_comparison = doc.add_table(rows=3, cols=3)
    tbl_comparison.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_headers = ["Pipeline Scenario", "Jest Test Runner Output", "GitHub Actions CI Outcome"]
    comp_data = [
        ("Deliberately Broken Test (expect 999)", "FAIL __tests__/api.test.js\nExpected: 999, Received: 200\nExit Code: 1", "FAILED (Red X)\nBuild halted, alerts triggered, PR merge blocked."),
        ("Corrected Test (expect 200)", "PASS __tests__/api.test.js\n7 passed, 7 total (1.112s)\nExit Code: 0", "PASSED (Green Checkmark)\nBuild succeeded, status badge updated to 'passing'.")
    ]
    for col_idx, h_text in enumerate(comp_headers):
        cell_h = tbl_comparison.cell(0, col_idx)
        set_cell_background(cell_h, "F1F5F9")
        p = cell_h.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(h_text)
        r.font.name = FONT_NAME
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = COLOR_BLACK

    for row_idx, row_values in enumerate(comp_data):
        for col_idx, val in enumerate(row_values):
            cell_d = tbl_comparison.cell(row_idx + 1, col_idx)
            set_cell_background(cell_d, "FFFFFF")
            p = cell_d.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(val)
            r.font.name = FONT_NAME
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_BLACK

    # Screenshot placeholder
    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_before = Pt(12)
    p_box.paragraph_format.space_after = Pt(8)

    tbl_shot = doc.add_table(rows=1, cols=1)
    tbl_shot.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_shot = tbl_shot.cell(0, 0)
    set_cell_background(c_shot, "F3F4F6")
    c_shot.width = Inches(6.5)
    set_cell_margins(c_shot, top=180, bottom=180, left=150, right=150)
    p_sh = c_shot.paragraphs[0]
    p_sh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sh = p_sh.add_run("[ INSERT SCREENSHOTS OF GITHUB ACTIONS WORKFLOW RUNS (FAILED RED PIPELINE VS. PASSED GREEN PIPELINE), JEST TERMINAL TEST RESULTS, AND README STATUS BADGE HERE ]")
    r_sh.font.name = FONT_NAME
    r_sh.font.size = Pt(11)
    r_sh.font.bold = True
    r_sh.font.color.rgb = COLOR_BLACK

    # 9. Conclusion
    add_section_header("9. Conclusion")
    p_con = doc.add_paragraph()
    r = p_con.add_run(
        "In Practical 12, an enterprise-grade Continuous Integration (CI/CD) pipeline was implemented using GitHub Actions. Code quality is strictly guarded by an Oxlint static analysis job, while regression prevention is enforced through an automated 7-point Jest and Supertest integration suite with an isolated MongoDB container service. Gating via pull requests prevents flawed code from reaching production, and the dynamic README status badge provides transparent visibility into application health."
    )
    r.font.name = FONT_NAME
    r.font.size = Pt(11)
    r.font.color.rgb = COLOR_BLACK

    output_path = r"D:\SEM-5\Advance Web Development Frameworks\docs\Practical-12-Documentation.docx"
    doc.save(output_path)
    print(f"Successfully generated Practical 12 Word document at: {output_path}")

if __name__ == "__main__":
    create_document()
