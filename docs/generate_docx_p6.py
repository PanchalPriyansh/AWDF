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

    r_pr = p_title.add_run("PRACTICAL REPORT: PRACTICAL NO. 6\n")
    r_pr.font.name = FONT_NAME
    r_pr.font.size = Pt(13)
    r_pr.font.bold = True
    r_pr.font.color.rgb = COLOR_BLACK

    r_pr_name = p_title.add_run("Full-Stack Integration: Wiring React Frontend to Node/Express/MongoDB Backend")
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
    r = p_obj.add_run("To wire the React frontend with the Node.js/Express.js and MongoDB backend into a cohesive, fully functional full-stack web application supporting persistent CRUD operations, optimistic UI updates, confirmation dialogs, and toast notifications.")
    r.font.name = FONT_NAME
    r.font.size = Pt(11)
    r.font.color.rgb = COLOR_BLACK

    # 2. Prerequisites & Resources
    add_section_header("2. Prerequisites & Resources Required")
    p = doc.add_paragraph()
    add_bullet(p, "• Prerequisites: ", "Practicals 1–3 completed React UI with routing and state management, Practicals 4–5 completed Express & MongoDB backend, CORS concepts.")
    p = doc.add_paragraph()
    add_bullet(p, "• Tools & Technologies: ", "Node.js (v24.x), npm (v11.x), React 18+, Vite, Express.js (v4.21.x), Mongoose ODM, MongoDB Server, CORS, Postman / Browser DevTools.")

    # 3. Problem Definition
    add_section_header("3. Problem Definition")
    for req in [
        "1. Connect the React Task UI to the Node+MongoDB backend built in Practicals 4 and 5.",
        "2. Replace the GitHub repository data from Practical 3 with live task data from your own backend.",
        "3. Support creating, viewing, updating, and deleting tasks from the React user interface through API calls.",
        "4. Confirm all data persists in MongoDB across browser reloads.",
        "5. Handle loading and error states for every API interaction (fetch, create, update, delete)."
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
        ("• Supplementary 1: ", "Implemented an optimistic UI update for task creation: immediately renders the new task in the local state list before the server confirms, with automatic rollback if the API fails."),
        ("• Supplementary 2: ", "Added an accessible confirmation modal dialog component (ConfirmModal.jsx) that prompts the user before permanently deleting any task from MongoDB."),
        ("• Supplementary 3: ", "Added a reusable Toast notification component (Toast.jsx) displaying animated success/failure alerts with auto-dismissal after every CRUD operation.")
    ]:
        p = doc.add_paragraph()
        add_bullet(p, sup[0], sup[1])

    # 5. Architecture & Full-Stack Data Flow
    add_section_header("5. Architecture & Full-Stack Data Flow")
    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.left_indent = Inches(0.25)
    r = p_arch.add_run(
        "React Frontend (http://localhost:5173)\n"
        "  ├── src/services/api.js (Centralized API client using Fetch API)\n"
        "  │     ├── getTasks()    → GET    http://localhost:5000/tasks\n"
        "  │     ├── createTask()   → POST   http://localhost:5000/tasks\n"
        "  │     ├── updateTask()   → PUT    http://localhost:5000/tasks/:id\n"
        "  │     └── deleteTask()   → DELETE http://localhost:5000/tasks/:id\n"
        "  │\n"
        "  ▼ Express Backend (http://localhost:5000)\n"
        "  │     ├── CORS Middleware: Allows cross-origin requests from frontend (Port 5173)\n"
        "  │     └── Express Router & Mongoose Controller Handlers\n"
        "  │\n"
        "  ▼ MongoDB Database (tasks collection)\n"
        "        └── Persistent storage confirmed across browser & server restarts\n"
        "  │\n"
        "  ▼ UI State Synchronization (Optimistic Updates, Modals, & Toasts)"
    )
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    r.font.color.rgb = COLOR_BLACK

    # 6. Core Implementation Snippets (Concise as requested)
    add_section_header("6. Core Implementation Snippets")

    p_note = doc.add_paragraph()
    r_n = p_note.add_run("Note: Only newly introduced API services, optimistic state mutations, modal dialogs, and toast notifications are documented below to maintain a concise submission report.")
    r_n.font.name = FONT_NAME
    r_n.font.size = Pt(10)
    r_n.font.italic = True
    r_n.font.color.rgb = COLOR_BLACK

    snippets = [
        (
            "Snippet 1: Centralized API Service (src/services/api.js)",
            "Demonstrates configuring client-side HTTP calls to the Express backend (Step 2):",
            """const BASE_URL = 'http://localhost:5000';

export async function getTasks() {
  const res = await fetch(`${BASE_URL}/tasks`);
  return res.json();
}

export async function createTask(taskData) {
  const res = await fetch(`${BASE_URL}/tasks`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(taskData),
  });
  return res.json();
}

export async function updateTask(id, updateData) {
  const res = await fetch(`${BASE_URL}/tasks/${id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(updateData),
  });
  return res.json();
}

export async function deleteTask(id) {
  const res = await fetch(`${BASE_URL}/tasks/${id}`, { method: 'DELETE' });
  return res.json();
}"""
        ),
        (
            "Snippet 2: Optimistic UI Task Creation (src/pages/Projects.jsx)",
            "Demonstrates immediately rendering the task before server confirmation, with rollback on error (Supplementary 1):",
            """const handleCreateTask = async (e) => {
  e.preventDefault();
  if (!newTitle.trim()) return;

  const tempId = `optimistic-${Date.now()}`;
  const optimisticTask = {
    _id: tempId,
    title: newTitle.trim(),
    description: newDesc.trim(),
    priority: newPriority,
    completed: false,
    createdAt: new Date().toISOString(),
    isOptimistic: true // Marker for visual sync badge
  };

  // 1. Immediately update UI state
  setTasks((prev) => [optimisticTask, ...prev]);

  try {
    // 2. Send POST request to MongoDB backend
    const res = await createTask(optimisticTask);
    // 3. Replace optimistic task with confirmed MongoDB document
    setTasks((prev) => prev.map((t) => (t._id === tempId ? res.data : t)));
    showToast(`Task "${res.data.title}" saved to MongoDB!`, 'success');
  } catch (err) {
    // 4. Rollback on failure
    setTasks((prev) => prev.filter((t) => t._id !== tempId));
    showToast(`Failed: ${err.message}`, 'error');
  }
};"""
        ),
        (
            "Snippet 3: Real-Time Toggle & Delete Mutations (src/pages/Projects.jsx)",
            "Demonstrates updating completion status and invoking delete with loading states (Step 5):",
            """// Toggle Task Completion (PUT /tasks/:id)
const handleToggleCompleted = async (task) => {
  const updatedStatus = !task.completed;
  setTasks((prev) => prev.map((t) => (t._id === task._id ? { ...t, completed: updatedStatus } : t)));
  try {
    await updateTask(task._id, { completed: updatedStatus });
    showToast(updatedStatus ? 'Marked as completed!' : 'Marked as active!', 'success');
  } catch (err) {
    setTasks((prev) => prev.map((t) => (t._id === task._id ? { ...t, completed: task.completed } : t)));
    showToast(err.message, 'error');
  }
};

// Confirm Delete Handler (DELETE /tasks/:id)
const handleConfirmDelete = async () => {
  try {
    await deleteTask(taskToDelete._id);
    setTasks((prev) => prev.filter((t) => t._id !== taskToDelete._id));
    showToast(`Task "${taskToDelete.title}" deleted from MongoDB.`, 'success');
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    setTaskToDelete(null);
  }
};"""
        ),
        (
            "Snippet 4: Delete Confirmation Dialog Component (src/components/ConfirmModal.jsx)",
            "Demonstrates preventing accidental data loss by prompting user confirmation before deletion (Supplementary 2):",
            """function ConfirmModal({ isOpen, title, message, onConfirm, onCancel, isProcessing }) {
  if (!isOpen) return null;

  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true">
      <div className="modal-card">
        <h3>{title}</h3>
        <p>{message}</p>
        <div className="modal-actions">
          <button className="btn-secondary" onClick={onCancel} disabled={isProcessing}>Cancel</button>
          <button className="btn-danger-confirm" onClick={onConfirm} disabled={isProcessing}>
            {isProcessing ? 'Deleting...' : 'Delete Task'}
          </button>
        </div>
      </div>
    </div>
  );
}
export default ConfirmModal;"""
        ),
        (
            "Snippet 5: Toast Notification Feedback Component (src/components/Toast.jsx)",
            "Demonstrates accessible non-intrusive notification feedback auto-dismissed after 4s (Supplementary 3):",
            """function Toast({ message, type = 'success', onClose }) {
  useEffect(() => {
    if (!message) return;
    const timer = setTimeout(() => onClose && onClose(), 4000);
    return () => clearTimeout(timer);
  }, [message, onClose]);

  if (!message) return null;
  return (
    <div className={`toast-notification toast-${type}`} role="status">
      <span>{type === 'success' ? '✅' : '⚠️'}</span>
      <p>{message}</p>
      <button onClick={onClose}>✕</button>
    </div>
  );
}
export default Toast;"""
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
    r_ti = p_tbl_intro.add_run("Summary of verified full-stack CRUD operations and user interactions:")
    r_ti.font.name = FONT_NAME
    r_ti.font.size = Pt(10.5)
    r_ti.font.color.rgb = COLOR_BLACK

    test_table = doc.add_table(rows=7, cols=4)
    test_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_headers = ["Action / Operation", "Underlying API Call", "UI State Behavior", "Observed Outcome"]
    test_data = [
        ("Task Creation", "POST /tasks", "Optimistic UI update", "Task appears instantly with 'Syncing...' badge, confirmed in MongoDB."),
        ("Status Toggle", "PUT /tasks/:id", "Reactive state update", "Checkbox strikes through title and updates MongoDB completed flag."),
        ("Delete Task Click", "None (pre-flight)", "Modal popup opened", "ConfirmModal prompts user with task title before permanent delete."),
        ("Confirm Deletion", "DELETE /tasks/:id", "Document removed", "Task is removed from UI and deleted from MongoDB collection."),
        ("Toast Feedback", "Local Event", "Animated toast alert", "Green success toast or red error toast displayed after each action."),
        ("Browser Refresh", "GET /tasks", "Persistent re-fetch", "All data persists intact in MongoDB and reloads seamlessly."),
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
    r_sh = p_sh.add_run("[ INSERT SCREENSHOTS OF TASK MANAGER UI, OPTIMISTIC UPDATE, DELETE CONFIRMATION MODAL & SUCCESS TOAST HERE ]")
    r_sh.font.name = FONT_NAME
    r_sh.font.size = Pt(11)
    r_sh.font.bold = True
    r_sh.font.color.rgb = COLOR_BLACK

    # 8. Conclusion
    add_section_header("8. Conclusion")
    p_con = doc.add_paragraph()
    r = p_con.add_run(
        "In Practical 6, the React frontend was successfully integrated with the Node.js/Express.js backend and MongoDB database into a complete MERN full-stack application. End-to-end CRUD capabilities were enabled with optimistic UI task creation, accessible delete confirmation dialogs, toast notifications, and persistent database storage."
    )
    r.font.name = FONT_NAME
    r.font.size = Pt(11)
    r.font.color.rgb = COLOR_BLACK

    output_path = r"D:\SEM-5\Advance Web Development Frameworks\docs\Practical-06-Documentation.docx"
    doc.save(output_path)
    print(f"Successfully generated Practical 6 Word document at: {output_path}")

if __name__ == "__main__":
    create_document()
