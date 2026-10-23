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

    r_pr = p_title.add_run("PRACTICAL REPORT: PRACTICAL NO. 7\n")
    r_pr.font.name = FONT_NAME
    r_pr.font.size = Pt(13)
    r_pr.font.bold = True
    r_pr.font.color.rgb = COLOR_BLACK

    r_pr_name = p_title.add_run("JWT-Based Authentication and Input Validation Middleware Pipeline")
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
    r = p_obj.add_run("To implement secure JWT-based authentication and input validation as part of the Express middleware pipeline, protecting full-stack Task Management API endpoints with token verification, password hashing, and payload validation.")
    r.font.name = FONT_NAME
    r.font.size = Pt(11)
    r.font.color.rgb = COLOR_BLACK

    # 2. Prerequisites & Resources
    add_section_header("2. Prerequisites & Resources Required")
    p = doc.add_paragraph()
    add_bullet(p, "• Prerequisites: ", "Practical 6 completed working full-stack Task Management application, foundational concepts of password hashing (bcrypt), salted cryptography, and JSON Web Tokens (JWT).")
    p = doc.add_paragraph()
    add_bullet(p, "• Tools & Technologies: ", "Node.js (v18+ / v24.x), npm, bcryptjs, jsonwebtoken, Express.js (v4.21.x), Mongoose ODM (v9.x), MongoDB Server (v8.x local), React 18+, Vite, Postman / Thunder Client.")

    # 3. Problem Definition
    add_section_header("3. Problem Definition")
    for req in [
        "1. Add user registration and login functionality to the Task Management application.",
        "2. Hash passwords using bcryptjs before saving the User document to MongoDB.",
        "3. Implement JWT token generation on successful login, with a reasonable expiry (1 hour).",
        "4. Protect all task routes using an authentication middleware that verifies the JWT from the Authorization header.",
        "5. Add server-side input validation middleware that rejects malformed requests (e.g., missing title) before they reach the database or controller."
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
        ("• Supplementary 1: ", "Implemented a protected /me (and /auth/me) endpoint that returns the currently logged-in user's details (ID, email, creation date) without password exposure, using the decoded JWT payload."),
        ("• Supplementary 2: ", "Implemented frontend token expiry handling: automatically clears stored tokens on receiving HTTP 401 Unauthorized and redirects the client interface back to the login view."),
        ("• Supplementary 3: ", "Added a client-side logout mechanism that purges the JWT token and user profile from localStorage and dispatches an event resetting application state.")
    ]:
        p = doc.add_paragraph()
        add_bullet(p, sup[0], sup[1])

    # 5. Architecture & Middleware Pipeline Flow
    add_section_header("5. Architecture & Middleware Pipeline Flow")
    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.left_indent = Inches(0.25)
    r = p_arch.add_run(
        "POST /register ──► Hash Password (bcrypt) ──► Save User Document ──► 201 Created\n"
        "POST /login    ──► Verify Password (bcrypt) ──► Sign JWT (1h Expiry) ──► Return Token (200 OK)\n\n"
        "Protected Route Middleware Pipeline:\n"
        "Client Request (Authorization: Bearer <token>)\n"
        "       │\n"
        "       ▼\n"
        "[Auth Middleware] ──────── (Missing/Invalid Token) ────────► 401 Unauthorized\n"
        "       │ (Valid Token: Attaches req.user = decoded)\n"
        "       ▼\n"
        "[Validation Middleware] ── (Missing/Blank Title) ─────────► 400 Bad Request\n"
        "       │ (Valid Task Body)\n"
        "       ▼\n"
        "Task Controller Handlers (GET /tasks, POST /tasks, PUT /tasks/:id, DELETE /tasks/:id)\n"
        "       │\n"
        "       ▼\n"
        "MongoDB Database (users & tasks collections)"
    )
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_BLACK

    # 6. Core Implementation Snippets (Concise as requested)
    add_section_header("6. Core Implementation Snippets")

    p_note = doc.add_paragraph()
    r_n = p_note.add_run("Note: In compliance with faculty guidelines, only the newly implemented authentication schema, JWT token generation, auth & validation middlewares, and client auth state handlers are documented below.")
    r_n.font.name = FONT_NAME
    r_n.font.size = Pt(10)
    r_n.font.italic = True
    r_n.font.color.rgb = COLOR_BLACK

    snippets = [
        (
            "Snippet 1: User Model Schema (models/User.js)",
            "Demonstrates defining the User schema with unique email and hashed password (Step 2):",
            """const mongoose = require('mongoose');

const userSchema = new mongoose.Schema({
  email: {
    type: String,
    required: [true, 'Email address is required'],
    unique: true,
    lowercase: true,
    trim: true,
  },
  password: {
    type: String,
    required: [true, 'Password is required'],
    minlength: [6, 'Password must be at least 6 characters long'],
  },
  createdAt: { type: Date, default: Date.now },
});

module.exports = mongoose.model('User', userSchema);"""
        ),
        (
            "Snippet 2: User Registration with bcrypt Hashing (server.js - Step 3)",
            "Demonstrates hashing plaintext password with bcrypt before persisting User in MongoDB:",
            """// POST /register - Hash password before saving to MongoDB
app.post('/register', async (req, res, next) => {
  try {
    const { email, password } = req.body;
    if (!email || !password || password.length < 6) {
      return res.status(400).json({ success: false, message: 'Valid email and password (min 6 chars) required.' });
    }
    // Check duplicate
    if (await User.findOne({ email: email.toLowerCase().trim() })) {
      return res.status(409).json({ success: false, message: 'User already exists with this email.' });
    }
    // Hash password with bcrypt (salt rounds = 10)
    const hashedPassword = await bcrypt.hash(password, 10);
    const user = await User.create({ email: email.toLowerCase().trim(), password: hashedPassword });

    res.status(201).json({ success: true, message: 'User registered successfully', user: { id: user._id, email: user.email } });
  } catch (err) { next(err); }
});"""
        ),
        (
            "Snippet 3: User Login & JWT Token Signing (server.js - Step 4)",
            "Demonstrates validating password with bcrypt.compare and signing JWT with 1-hour expiry:",
            """// POST /login - Verify password and sign JWT
app.post('/login', async (req, res, next) => {
  try {
    const { email, password } = req.body;
    const user = await User.findOne({ email: email.toLowerCase().trim() });
    if (!user) return res.status(401).json({ success: false, message: 'Invalid email or password.' });

    // Compare password with hashed database value
    const isMatch = await bcrypt.compare(password, user.password);
    if (!isMatch) return res.status(401).json({ success: false, message: 'Invalid email or password.' });

    // Sign JWT token valid for 1 hour
    const token = jwt.sign({ id: user._id, email: user.email }, process.env.JWT_SECRET, { expiresIn: '1h' });
    res.status(200).json({ success: true, message: 'Login successful', token, user: { id: user._id, email: user.email } });
  } catch (err) { next(err); }
});"""
        ),
        (
            "Snippet 4: JWT Authentication Middleware (middleware/auth.js - Step 5)",
            "Demonstrates verifying Bearer token and attaching decoded user identity to req.user:",
            """const jwt = require('jsonwebtoken');

const authMiddleware = (req, res, next) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ success: false, message: 'Access denied. No token provided.' });
  }
  const token = authHeader.split(' ')[1];
  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    req.user = decoded; // Attach decoded user { id, email }
    next();
  } catch (err) {
    const message = err.name === 'TokenExpiredError' ? 'Your session has expired.' : 'Invalid token.';
    return res.status(401).json({ success: false, error: 'Unauthorized', message });
  }
};
module.exports = authMiddleware;"""
        ),
        (
            "Snippet 5: Task Input Validation Middleware (middleware/validateTask.js - Step 7)",
            "Demonstrates server-side input validation rejecting empty titles before database interaction:",
            """const validateCreateTask = (req, res, next) => {
  const { title, priority } = req.body;
  if (!title || typeof title !== 'string' || title.trim().length === 0) {
    return res.status(400).json({
      success: false,
      error: 'Validation Error',
      message: 'Task title is required and cannot be empty or blank whitespace.'
    });
  }
  if (priority && !['low', 'medium', 'high'].includes(priority.toLowerCase())) {
    return res.status(400).json({ success: false, message: `Invalid priority '${priority}'. Allowed: low, medium, high` });
  }
  next();
};
module.exports = { validateCreateTask };"""
        ),
        (
            "Snippet 6: Protected /me Profile Route (server.js - Supplementary 1)",
            "Demonstrates retrieving authenticated user details from decoded JWT without exposing password:",
            """// GET /me - Returns authenticated user details from token payload
app.get('/me', authMiddleware, async (req, res, next) => {
  try {
    const user = await User.findById(req.user.id).select('-password');
    if (!user) return res.status(404).json({ success: false, message: 'User not found.' });
    res.status(200).json({
      success: true,
      user: { id: user._id, email: user.email, createdAt: user.createdAt }
    });
  } catch (err) { next(err); }
});"""
        ),
        (
            "Snippet 7: Frontend Auth State & 401 Expiry Interceptor (src/services/api.js - Supplementary 2 & 3)",
            "Demonstrates client-side token attachment, 401 session expiry handling, and logout cleanup:",
            """export function getAuthHeaders() {
  const token = localStorage.getItem('task_manager_jwt');
  const headers = { 'Content-Type': 'application/json' };
  if (token) headers['Authorization'] = `Bearer ${token}`;
  return headers;
}

// Intercepts 401 Unauthorized to trigger automatic login redirect
function handleAuthError(data) {
  localStorage.removeItem('task_manager_jwt');
  localStorage.removeItem('task_manager_user');
  window.dispatchEvent(new CustomEvent('auth:expired', { detail: data.message }));
  throw new Error(data.message || 'Session expired. Please log in again.');
}

// Client logout mechanism
export function logoutUser() {
  localStorage.removeItem('task_manager_jwt');
  localStorage.removeItem('task_manager_user');
  window.dispatchEvent(new CustomEvent('auth:logout'));
}"""
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

    # 7. Output & Verification Results Table
    add_section_header("7. Output & Verification Results")

    p_tbl_intro = doc.add_paragraph()
    r_ti = p_tbl_intro.add_run("Summary of tested and verified authentication, authorization, and validation pipeline flows:")
    r_ti.font.name = FONT_NAME
    r_ti.font.size = Pt(10.5)
    r_ti.font.color.rgb = COLOR_BLACK

    test_table = doc.add_table(rows=11, cols=5)
    test_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_headers = ["Endpoint / Operation", "Method", "Auth Header", "Status", "Observed Behavior & Output"]
    test_data = [
        ("POST /register", "POST", "None", "201 Created", "Password hashed with bcrypt (salt 10); User saved in MongoDB."),
        ("POST /register (Duplicate)", "POST", "None", "409 Conflict", "Duplicate email detected; rejects with structured error."),
        ("POST /login (Bad Password)", "POST", "None", "401 Unauthorized", "bcrypt.compare returns false; access denied."),
        ("POST /login (Valid)", "POST", "None", "200 OK", "Signs JWT valid for 1 hour; returns Bearer token & user info."),
        ("GET /me (Supplementary 1)", "GET", "Bearer <token>", "200 OK", "Decodes JWT; returns user id, email, and created timestamp."),
        ("GET /tasks (No Token)", "GET", "None", "401 Unauthorized", "Auth middleware blocks request: 'No token provided'."),
        ("GET /tasks (Valid Token)", "GET", "Bearer <token>", "200 OK", "Verifies JWT; returns user-specific tasks from MongoDB."),
        ("POST /tasks (No Title)", "POST", "Bearer <token>", "400 Bad Request", "Validation middleware intercepts before controller: missing title."),
        ("POST /tasks (Valid)", "POST", "Bearer <token>", "201 Created", "Input validated, userId attached, task created in MongoDB."),
        ("Frontend 401 Intercept", "Client Event", "Expired Token", "Auto Redirect", "Clears localStorage token; redirects user to Login view."),
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
    r_sh = p_sh.add_run("[ INSERT SCREENSHOTS OF POSTMAN/THUNDER CLIENT REGISTRATION, LOGIN, JWT TOKEN ACQUISITION, PROTECTED /TASKS & /ME ENDPOINT HERE ]")
    r_sh.font.name = FONT_NAME
    r_sh.font.size = Pt(11)
    r_sh.font.bold = True
    r_sh.font.color.rgb = COLOR_BLACK

    # 8. Conclusion
    add_section_header("8. Conclusion")
    p_con = doc.add_paragraph()
    r = p_con.add_run(
        "In Practical 7, JWT-based authentication and input validation were successfully engineered into the Express middleware pipeline. User passwords were encrypted using bcrypt with salt rounds before database persistence, JWT tokens were generated upon valid authentication with 1-hour expiry, all task endpoints were guarded with authorization verification, server-side validation rejected malformed payloads, and the client frontend handled token persistence, 401 session expiry redirects, and logout mechanisms."
    )
    r.font.name = FONT_NAME
    r.font.size = Pt(11)
    r.font.color.rgb = COLOR_BLACK

    output_path = r"D:\SEM-5\Advance Web Development Frameworks\docs\Practical-07-Documentation.docx"
    doc.save(output_path)
    print(f"Successfully generated Practical 7 Word document at: {output_path}")

if __name__ == "__main__":
    create_document()
