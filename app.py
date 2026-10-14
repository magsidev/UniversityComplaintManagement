"""
========================================================================================
Smart University Complaint System (SUCS v1.0) - Prototype Phase Backend
========================================================================================
Architecture: Three-Tier Web Architecture (Presentation, Business Logic, Data Storage)
Adopted Methodology: VU Process Model (CS605 - Hybrid Waterfall-Spiral)
Tech Stack: Python 3, Flask Web Framework, SQLite Relational Database, Jinja2 Templating
Security: Werkzeug PBKDF2:SHA256 Password Hashing, Signed Cookie Sessions, RBAC Guards

Core Business Logic Implemented in this File:
1. Role-Based Access Control (RBAC): Strict segregation between Student, Admin, and HOD.
2. Monthly Quota Engine: Strict 5-complaints-per-month limit per student account.
3. Urgency & Emergency Classification Engine: Keyword & Emergency flag priority mapping.
4. Conflict-of-Interest Filter: Routes complaints directed against Admins directly to HOD.
5. Smart Grouping & Cascading Resolution: Atomic bulk-resolution of duplicate complaints.
6. Admin & HOD Response Engine: Attaching official resolution notes to solved tickets.
========================================================================================
"""

import os
import sqlite3
import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
from werkzeug.security import generate_password_hash, check_password_hash

# --------------------------------------------------------------------------------------
# 1. Flask Application Initialization & Configuration
# --------------------------------------------------------------------------------------
app = Flask(__name__)

# Secret key used by Flask to cryptographically sign session cookies (tamper-proof)
app.secret_key = 'sucs_prototype_flask_secret_key_2026'

# Absolute path to the SQLite prototype database file
DB_PATH = os.path.join(os.path.dirname(__file__), 'sucs_prototype.db')


# --------------------------------------------------------------------------------------
# 2. Database Connection Factory & Schema Initialization
# --------------------------------------------------------------------------------------
def get_db():
    """
    Creates and returns an active SQLite database connection.
    - row_factory = sqlite3.Row: Allows accessing columns by name (e.g., row['Username'])
      rather than relying on numeric index tuples.
    - PRAGMA foreign_keys = ON: Enforces relational foreign key constraints in SQLite.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db():
    """
    Initializes the database schema if it doesn't already exist and applies schema migrations.
    - Creates 'User' table: Stores authentication credentials and system role.
    - Creates 'Complaint' table: Stores complaints with categories, targets, urgency, and statuses.
    - Migrates schema safely using ALTER TABLE if new columns were added.
    - Seeds initial demo accounts (student1, student2, admin1, hod1) with hashed passwords.
    """
    conn = get_db()
    cursor = conn.cursor()
    
    # ----------------------------------------------------------------------------------
    # Model 1: User Table (SRS Section 4 / ERD Section 2)
    # Columns: ID (Primary Key), Username, PasswordHash (PBKDF2), Role (RBAC Discriminator)
    # ----------------------------------------------------------------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS User (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            Username TEXT UNIQUE NOT NULL,
            PasswordHash TEXT NOT NULL,
            Role TEXT CHECK(Role IN ('Student', 'Admin', 'HOD')) NOT NULL
        );
    ''')

    # ----------------------------------------------------------------------------------
    # Model 2: Complaint Table (SRS Section 4 / ERD Section 2)
    # Columns: ID, StudentID (Foreign Key), Category, DirectedAgainst, Description,
    #          UrgencyLevel, Status, DateSubmitted, GroupID (for Smart Grouping), AdminComment
    # Note: Status supports 'Pending', 'In-Progress', 'Forwarded', 'Solved'
    # ----------------------------------------------------------------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Complaint (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            StudentID INTEGER NOT NULL,
            Category TEXT NOT NULL,
            DirectedAgainst TEXT NOT NULL,
            Description TEXT DEFAULT '',
            UrgencyLevel TEXT CHECK(UrgencyLevel IN ('Emergency', 'High', 'Medium', 'Low')) NOT NULL,
            Status TEXT CHECK(Status IN ('Pending', 'In-Progress', 'Forwarded', 'Solved')) DEFAULT 'Pending',
            DateSubmitted DATETIME DEFAULT CURRENT_TIMESTAMP,
            GroupID INTEGER DEFAULT NULL,
            AdminComment TEXT DEFAULT NULL,
            FOREIGN KEY (StudentID) REFERENCES User(ID)
        );
    ''')

    # Schema migration safeguard for 'In-Progress' status support in existing databases
    try:
        table_sql = cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='Complaint'").fetchone()
        if table_sql and "'In-Progress'" not in table_sql[0]:
            cursor.execute("PRAGMA foreign_keys = OFF;")
            cursor.execute('''
                CREATE TABLE Complaint_migrated (
                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                    StudentID INTEGER NOT NULL,
                    Category TEXT NOT NULL,
                    DirectedAgainst TEXT NOT NULL,
                    Description TEXT DEFAULT '',
                    UrgencyLevel TEXT CHECK(UrgencyLevel IN ('Emergency', 'High', 'Medium', 'Low')) NOT NULL,
                    Status TEXT CHECK(Status IN ('Pending', 'In-Progress', 'Forwarded', 'Solved')) DEFAULT 'Pending',
                    DateSubmitted DATETIME DEFAULT CURRENT_TIMESTAMP,
                    GroupID INTEGER DEFAULT NULL,
                    AdminComment TEXT DEFAULT NULL,
                    FOREIGN KEY (StudentID) REFERENCES User(ID)
                );
            ''')
            cursor.execute('''
                INSERT INTO Complaint_migrated (ID, StudentID, Category, DirectedAgainst, Description, UrgencyLevel, Status, DateSubmitted, GroupID, AdminComment)
                SELECT ID, StudentID, Category, DirectedAgainst, Description, UrgencyLevel, Status, DateSubmitted, GroupID, AdminComment FROM Complaint;
            ''')
            cursor.execute("DROP TABLE Complaint;")
            cursor.execute("ALTER TABLE Complaint_migrated RENAME TO Complaint;")
            cursor.execute("PRAGMA foreign_keys = ON;")
    except Exception:
        pass

    # Schema migration safeguards: ensures existing database files contain newly added columns
    try:
        cursor.execute("ALTER TABLE Complaint ADD COLUMN Description TEXT DEFAULT '';")
    except Exception:
        pass  # Column already exists in the table

    try:
        cursor.execute("ALTER TABLE Complaint ADD COLUMN AdminComment TEXT DEFAULT NULL;")
    except Exception:
        pass  # Column already exists in the table

    conn.commit()

    # ----------------------------------------------------------------------------------
    # Seed Default User Accounts if the database is newly created
    # All demo accounts share the initial password: 'password123'
    # ----------------------------------------------------------------------------------
    user_count = cursor.execute("SELECT COUNT(*) FROM User").fetchone()[0]
    if user_count == 0:
        # Generate secure hash using PBKDF2 with SHA-256 (standard NIST recommendation)
        pwd = generate_password_hash('password123', method='pbkdf2:sha256')
        cursor.executemany("INSERT INTO User (Username, PasswordHash, Role) VALUES (?, ?, ?)", [
            ('student1', pwd, 'Student'),
            ('student2', pwd, 'Student'),
            ('admin1', pwd, 'Admin'),
            ('hod1', pwd, 'HOD')
        ])
        conn.commit()

    conn.close()

# Execute schema setup once at application module load time
init_db()


# --------------------------------------------------------------------------------------
# 2B. Core Business Service: Auto-Grouping Engine for Similar Complaints
# --------------------------------------------------------------------------------------
def auto_group_similar_complaints(conn=None):
    """
    Auto-combines complaints related to the same person or same issue (Category & DirectedAgainst).
    - Scans all active / unresolved complaints (Status IN ('Pending', 'In-Progress')).
    - Identifies clusters that share the same Category and DirectedAgainst (case-insensitive).
    - Automatically links them under a common GroupID so Admin and HOD can solve all relevant
      complaints with a single click.
    """
    should_close = False
    if conn is None:
        conn = get_db()
        should_close = True

    try:
        clusters = conn.execute('''
            SELECT Category, LOWER(TRIM(DirectedAgainst)) as Target, COUNT(*) as cnt
            FROM Complaint
            WHERE Status IN ('Pending', 'In-Progress')
            GROUP BY Category, LOWER(TRIM(DirectedAgainst))
            HAVING cnt > 1
        ''').fetchall()

        for cluster in clusters:
            cat = cluster['Category']
            target = cluster['Target']

            rows = conn.execute('''
                SELECT ID, GroupID FROM Complaint
                WHERE Category = ? AND LOWER(TRIM(DirectedAgainst)) = ? AND Status IN ('Pending', 'In-Progress')
                ORDER BY ID ASC
            ''', (cat, target)).fetchall()

            if not rows:
                continue

            # Determine master group ID: reuse existing group ID or assign the earliest ticket ID
            existing_group_id = None
            for r in rows:
                if r['GroupID']:
                    existing_group_id = r['GroupID']
                    break

            master_group_id = existing_group_id if existing_group_id is not None else rows[0]['ID']

            # Link all matching complaints under master group ID
            for r in rows:
                if r['GroupID'] != master_group_id:
                    conn.execute("UPDATE Complaint SET GroupID = ? WHERE ID = ?", (master_group_id, r['ID']))

        conn.commit()
    finally:
        if should_close:
            conn.close()


# --------------------------------------------------------------------------------------
# 3. Core Business Service: Monthly Quota Engine (SRS Requirement FR3 / TC-10 to TC-12)
# --------------------------------------------------------------------------------------
def get_monthly_quota(student_id):
    """
    Enforces the student monthly quota policy (SRS FR3):
    - Each student account is allocated exactly 5 complaints per calendar month.
    - Calculates quota usage by counting submissions in the current calendar month ('YYYY-MM').
    - Returns a dictionary with:
        'used': total complaints submitted in the current month
        'limit': maximum allowed per month (5)
        'remaining': non-negative count of remaining submissions (0 to 5)
    """
    conn = get_db()
    current_month = datetime.datetime.now().strftime('%Y-%m')
    count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE StudentID = ? AND strftime('%Y-%m', DateSubmitted) = ?
    ''', (student_id, current_month)).fetchone()[0]
    conn.close()
    return {'used': count, 'limit': 5, 'remaining': max(0, 5 - count)}


# --------------------------------------------------------------------------------------
# 4. Authentication & Session RBAC Routes (Use Case UC-01 / TC-01 to TC-06)
# --------------------------------------------------------------------------------------
@app.route('/')
def index():
    """
    Root URL Dispatcher:
    - If a user is already authenticated in session, redirects directly to their role-specific dashboard.
    - If unauthenticated, redirects to the /login landing page.
    """
    if 'user_id' in session:
        role = session.get('role')
        if role == 'Student':
            return redirect(url_for('student_dashboard'))
        elif role == 'Admin':
            return redirect(url_for('admin_dashboard'))
        elif role == 'HOD':
            return redirect(url_for('hod_dashboard'))
    return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    """
    User Login Endpoint (UC-01):
    - GET: Renders the simple login page with demo quick-fill buttons.
    - POST:
      1. Validates that username and password are provided.
      2. Queries the 'User' table for matching username.
      3. Verifies password against stored PBKDF2 hash using check_password_hash().
      4. On success: Initializes encrypted Flask session (user_id, username, role).
      5. Enforces strict dashboard redirection based on authenticated role.
    """
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        # Input presence check
        if not username or not password:
            flash('Username and Password are required.', 'danger')
            return render_template('login.html')

        conn = get_db()
        user = conn.execute('SELECT * FROM User WHERE Username = ?', (username,)).fetchone()
        conn.close()

        # Validate password hash against provided plaintext password
        if user and check_password_hash(user['PasswordHash'], password):
            # Bind authenticated user state into signed HTTP cookie session
            session['user_id'] = user['ID']
            session['username'] = user['Username']
            session['role'] = user['Role']

            # Role-Based Redirection (RBAC)
            if user['Role'] == 'Student':
                return redirect(url_for('student_dashboard'))
            elif user['Role'] == 'Admin':
                return redirect(url_for('admin_dashboard'))
            elif user['Role'] == 'HOD':
                return redirect(url_for('hod_dashboard'))

        flash('Invalid Username or Password.', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    """
    Session Termination Endpoint:
    - Clears all session keys, invalidating the user's authenticated cookie.
    - Redirects to login with an informational notification message.
    """
    session.clear()
    flash('Logged out successfully.', 'info')
    return redirect(url_for('login'))


# --------------------------------------------------------------------------------------
# 5. Student Workflow Routes (Use Case UC-02 & UC-03 / TC-07 to TC-12)
# --------------------------------------------------------------------------------------
@app.route('/student')
def student_dashboard():
    """
    Student Portal Overview Dashboard:
    - RBAC Guard: Ensures only authenticated 'Student' users can access this page.
    - Displays overall ticket metrics: Open Count, In-Progress Count, Solved Count.
    - Calculates and displays current monthly quota utilization (used, limit, remaining).
    - Contains complaint intake form with 'Others' options and 1-Click Emergency Hazard trigger.
    - Provides link to dedicated Submitted Complaints History page (/student/complaints).
    """
    if session.get('role') != 'Student':
        flash('Unauthorized access to Student Portal. Please log in.', 'danger')
        return redirect(url_for('login'))

    student_id = session['user_id']
    quota = get_monthly_quota(student_id)

    conn = get_db()
    # Metrics breakdown for the student
    open_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE StudentID = ? AND Status IN ('Pending', 'In-Progress', 'Forwarded')
    ''', (student_id,)).fetchone()[0]

    in_progress_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE StudentID = ? AND Status = 'In-Progress'
    ''', (student_id,)).fetchone()[0]

    solved_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE StudentID = ? AND Status = 'Solved'
    ''', (student_id,)).fetchone()[0]

    total_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint WHERE StudentID = ?
    ''', (student_id,)).fetchone()[0]

    # All complaints for backward compatibility with existing tests and cards
    complaints = conn.execute('''
        SELECT * FROM Complaint WHERE StudentID = ? ORDER BY DateSubmitted DESC
    ''', (student_id,)).fetchall()
    conn.close()

    return render_template('student.html', 
                           quota=quota, 
                           open_count=open_count,
                           in_progress_count=in_progress_count,
                           solved_count=solved_count,
                           total_count=total_count,
                           complaints=complaints)


@app.route('/student/complaints')
def student_complaints():
    """
    Dedicated Student Submitted Complaints Page:
    - Displays all active, in-progress, and past resolved complaints filed by this student.
    - Provides status filtering, safe DOM inspection modals, and official administrative remarks.
    """
    if session.get('role') != 'Student':
        flash('Unauthorized access to Student Portal. Please log in.', 'danger')
        return redirect(url_for('login'))

    student_id = session['user_id']
    quota = get_monthly_quota(student_id)

    conn = get_db()
    complaints = conn.execute('''
        SELECT * FROM Complaint WHERE StudentID = ? ORDER BY DateSubmitted DESC
    ''', (student_id,)).fetchall()
    conn.close()

    return render_template('student_complaints.html', quota=quota, complaints=complaints)


@app.route('/submit-complaint', methods=['POST'])
def submit_complaint():
    """
    Complaint Submission Engine (UC-02 & UC-03):
    - RBAC Guard: Restricts action to students.
    - Quota Validation: Checks remaining monthly allowance (strict 5-complaint cap).
    - Custom Option Handling: Supports custom category & directed against if 'Others' is selected.
    - Urgency Classification Engine:
        * If student checked the 'Emergency' flag -> UrgencyLevel = 'Emergency'
        * If category is 'Harassment' or 'Infrastructure' (or text contains hazard/fire) -> 'High'
        * If category is 'Facilities' or 'IT/Network' -> 'Medium'
        * Otherwise -> 'Low'
    - Auto-Combining Engine: Automatically detects similar complaints (same category and target)
      and links them under a common GroupID for 1-click batch resolution.
    """
    if session.get('role') != 'Student':
        flash('Session expired or unauthorized. Please log in.', 'danger')
        return redirect(url_for('login'))

    student_id = session['user_id']
    quota = get_monthly_quota(student_id)

    # Enforce strict 5-complaint monthly cap (SRS FR3 / TC-11)
    if quota['remaining'] <= 0:
        flash('Monthly limit of 5 complaints reached for this account.', 'danger')
        return redirect(url_for('student_dashboard'))

    category = request.form.get('category', '').strip()
    directed_against = request.form.get('directed_against', '').strip()
    description = request.form.get('description', '').strip()
    is_emergency = request.form.get('is_emergency') == 'on'

    # Support for "Others" custom input
    if category == 'Others':
        custom_category = request.form.get('custom_category', '').strip()
        if custom_category:
            category = f"Others: {custom_category}"

    if directed_against == 'Others':
        custom_directed = request.form.get('custom_directed_against', '').strip()
        if custom_directed:
            directed_against = f"Others: {custom_directed}"

    # Validate required input fields
    if not category or not directed_against or not description:
        flash('Category, Directed Against, and Description fields are mandatory.', 'danger')
        return redirect(url_for('student_dashboard'))

    # Urgency & Priority Classification Logic (SRS FR5 & FR7 / TC-13 to TC-17)
    if is_emergency:
        urgency = 'Emergency'
    elif category in ['Harassment', 'Infrastructure'] or 'hazard' in description.lower() or 'fire' in description.lower():
        urgency = 'High'
    elif category in ['Facilities', 'IT/Network']:
        urgency = 'Medium'
    else:
        urgency = 'Low'

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO Complaint (StudentID, Category, DirectedAgainst, Description, UrgencyLevel, Status)
        VALUES (?, ?, ?, ?, ?, 'Pending')
    ''', (student_id, category, directed_against, description, urgency))
    new_complaint_id = cursor.lastrowid
    conn.commit()

    # Automatically group with similar existing active complaints (same category and target)
    auto_group_similar_complaints(conn)

    # Check if newly created complaint was assigned a GroupID
    updated_complaint = conn.execute("SELECT GroupID FROM Complaint WHERE ID = ?", (new_complaint_id,)).fetchone()
    conn.close()

    if updated_complaint and updated_complaint['GroupID']:
        flash(f'Complaint #{new_complaint_id} submitted ({urgency} priority) and auto-combined into Group #{updated_complaint["GroupID"]} for 1-click resolution.', 'success')
    else:
        flash(f'Complaint submitted successfully with {urgency} urgency status.', 'success')

    return redirect(url_for('student_dashboard'))


# --------------------------------------------------------------------------------------
# 6. Administrator Workflow Routes (Use Cases UC-04, UC-05, UC-06 / TC-13 to TC-27)
# --------------------------------------------------------------------------------------
@app.route('/admin')
def admin_dashboard():
    """
    Admin Queue Dashboard:
    - RBAC Guard: Restricts access exclusively to 'Admin' role.
    - AUTO-GROUPING ENGINE: Runs auto-combination for all active similar complaints.
    - OVERVIEW STATS: Calculates Pending, In-Progress, Emergency, and Past/Solved metrics.
    - AUTOMATED CONFLICT-OF-INTEREST FILTER (SRS FR6 / TC-18):
      Filters out complaints directed against 'admin' (routed directly to HOD).
    - ACTIVE QUEUE: Shows Pending & In-Progress tickets sorted Emergency/High first.
    - PAST COMPLAINTS: Fetches resolved tickets for the past complaints sidebar.
    """
    if session.get('role') != 'Admin':
        flash('Unauthorized access to Admin Portal. Please log in.', 'danger')
        return redirect(url_for('login'))

    conn = get_db()
    # Trigger auto-grouping for all similar complaints on queue load
    auto_group_similar_complaints(conn)

    # Overview Metrics for Admin Dashboard
    pending_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE Status = 'Pending' AND LOWER(DirectedAgainst) NOT LIKE '%admin%'
    ''').fetchone()[0]

    in_progress_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE Status = 'In-Progress' AND LOWER(DirectedAgainst) NOT LIKE '%admin%'
    ''').fetchone()[0]

    emergency_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE UrgencyLevel = 'Emergency' AND Status != 'Solved' AND LOWER(DirectedAgainst) NOT LIKE '%admin%'
    ''').fetchone()[0]

    solved_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE Status = 'Solved' AND LOWER(DirectedAgainst) NOT LIKE '%admin%'
    ''').fetchone()[0]

    # Active Complaints (Pending & In-Progress, sorted by Urgency then Date)
    complaints = conn.execute('''
        SELECT c.*, u.Username as StudentUsername 
        FROM Complaint c
        JOIN User u ON c.StudentID = u.ID
        WHERE LOWER(c.DirectedAgainst) NOT LIKE '%admin%' AND c.Status IN ('Pending', 'In-Progress')
        ORDER BY 
            CASE c.UrgencyLevel 
                WHEN 'Emergency' THEN 1 
                WHEN 'High' THEN 2 
                WHEN 'Medium' THEN 3 
                WHEN 'Low' THEN 4 
            END ASC,
            c.DateSubmitted DESC
    ''').fetchall()

    # Past Solved Complaints for Sidebar
    past_complaints = conn.execute('''
        SELECT c.*, u.Username as StudentUsername 
        FROM Complaint c
        JOIN User u ON c.StudentID = u.ID
        WHERE LOWER(c.DirectedAgainst) NOT LIKE '%admin%' AND c.Status = 'Solved'
        ORDER BY c.DateSubmitted DESC
    ''').fetchall()
    conn.close()

    return render_template('admin.html', 
                           complaints=complaints, 
                           past_complaints=past_complaints,
                           pending_count=pending_count,
                           in_progress_count=in_progress_count,
                           emergency_count=emergency_count,
                           solved_count=solved_count)


@app.route('/admin/action/<int:complaint_id>/<action_type>', methods=['POST'])
def admin_action(complaint_id, action_type):
    """
    Single Complaint Action Handler (UC-04 & UC-05):
    - 'solve': Marks complaint status as 'Solved' and records the admin's resolution response note.
    - 'in_progress': Sets status to 'In-Progress' with a MANDATORY update comment.
    - 'forward': Escalates ticket to HOD by setting status to 'Forwarded'.
    """
    if session.get('role') != 'Admin':
        flash('Session expired or unauthorized. Please log in.', 'danger')
        return redirect(url_for('login'))

    conn = get_db()
    if action_type == 'solve':
        comment = request.form.get('admin_comment', '').strip()
        conn.execute("UPDATE Complaint SET Status = 'Solved', AdminComment = ? WHERE ID = ?", (comment, complaint_id))
        flash(f'Complaint #{complaint_id} marked as Solved with response comment.', 'success')
    elif action_type == 'in_progress':
        comment = request.form.get('admin_comment', '').strip()
        if not comment:
            flash('Comment is mandatory when setting a complaint to In-Progress.', 'danger')
            conn.close()
            return redirect(url_for('admin_dashboard'))
        conn.execute("UPDATE Complaint SET Status = 'In-Progress', AdminComment = ? WHERE ID = ?", (comment, complaint_id))
        flash(f'Complaint #{complaint_id} marked as In-Progress with comment.', 'info')
    elif action_type == 'forward':
        conn.execute("UPDATE Complaint SET Status = 'Forwarded' WHERE ID = ?", (complaint_id,))
        flash(f'Complaint #{complaint_id} forwarded to HOD.', 'info')

    conn.commit()
    conn.close()
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/auto-group', methods=['POST'])
def admin_auto_group():
    """
    Manual Trigger to Auto-Combine all similar complaints in the system.
    """
    if session.get('role') != 'Admin':
        flash('Session expired or unauthorized. Please log in.', 'danger')
        return redirect(url_for('login'))

    conn = get_db()
    auto_group_similar_complaints(conn)
    conn.close()
    flash('Auto-combined all complaints related to the same person or same issue.', 'success')
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/group', methods=['POST'])
def group_complaints():
    """
    Smart Grouping Engine (SRS FR8 / UC-06 / TC-23 & TC-24):
    - Receives a list of complaint IDs selected via checkboxes on the Admin dashboard.
    - Requires at least 2 complaints to establish a group.
    - Sets GroupID across all selected complaints, linking them to a common master issue.
    """
    if session.get('role') != 'Admin':
        flash('Session expired or unauthorized. Please log in.', 'danger')
        return redirect(url_for('login'))

    complaint_ids = request.form.getlist('complaint_ids')
    if len(complaint_ids) < 2:
        flash('Select at least 2 complaints to group.', 'warning')
        return redirect(url_for('admin_dashboard'))

    conn = get_db()
    # Uses the primary ticket ID as the master group identifier
    group_id = int(complaint_ids[0])
    for cid in complaint_ids:
        conn.execute("UPDATE Complaint SET GroupID = ? WHERE ID = ?", (group_id, cid))
    
    conn.commit()
    conn.close()

    flash(f'Grouped {len(complaint_ids)} complaints under Group #{group_id}.', 'success')
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/resolve-group/<int:group_id>', methods=['POST'])
def resolve_group(group_id):
    """
    1-Click Cascading Group Resolution for Admin:
    - Solves all complaints relevant to each other in one single click.
    - In a single atomic database operation, cascades Status = 'Solved' and the resolution
      note across all complaints associated with the GroupID (or matching primary ID).
    """
    if session.get('role') != 'Admin':
        flash('Session expired or unauthorized. Please log in.', 'danger')
        return redirect(url_for('login'))

    comment = request.form.get('admin_comment', f'1-Click resolved under Group #{group_id}').strip()

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE Complaint SET Status = 'Solved', AdminComment = ? WHERE GroupID = ? OR ID = ?", (comment, group_id, group_id))
    updated_count = cursor.rowcount
    conn.commit()
    conn.close()

    flash(f'1-Click Solved all {updated_count} related complaints in Group #{group_id}.', 'success')
    return redirect(url_for('admin_dashboard'))


# --------------------------------------------------------------------------------------
# 7. Head of Department (HOD) Workflow Routes (SRS FR6 / UC-04 & UC-05)
# --------------------------------------------------------------------------------------
@app.route('/hod')
def hod_dashboard():
    """
    HOD Executive Desk & Analytics Dashboard:
    - RBAC Guard: Restricts access strictly to the 'HOD' role.
    - AUTO-GROUPING: Automatically clusters similar issues into groups for 1-click solving.
    - EXECUTIVE STATS:
      * Count of active tickets Admin currently has.
      * Count of Conflict-of-Interest cases (direct HOD routing).
      * Count of Escalated tickets forwarded by Admin.
      * Total Solved tickets count across university.
    - LIST OF HIGHEST REASON TICKETS & CHARTS:
      * Aggregates frequency counts by Category (highest reason complaints first).
      * Aggregates distribution by Urgency Level.
    - TWO ACTIVE QUEUES:
      1. Conflict of Interest Queue: Grievances directed against Admin staff.
      2. Escalated Tickets Queue: Complaints forwarded by Admins.
    - PAST COMPLAINTS SIDEBAR:
      * Lists all solved executive complaints.
    """
    if session.get('role') != 'HOD':
        flash('Unauthorized access to HOD Portal. Please log in.', 'danger')
        return redirect(url_for('login'))

    conn = get_db()
    # Auto-combine similar complaints
    auto_group_similar_complaints(conn)

    # Count of active tickets Admin currently has
    admin_ticket_count = conn.execute('''
        SELECT COUNT(*) FROM Complaint 
        WHERE Status IN ('Pending', 'In-Progress') AND LOWER(DirectedAgainst) NOT LIKE '%admin%'
    ''').fetchone()[0]

    # Sensitive complaints targeted against Admin staff (Conflict of Interest)
    conflict_tickets = conn.execute('''
        SELECT c.*, u.Username as StudentUsername 
        FROM Complaint c
        JOIN User u ON c.StudentID = u.ID
        WHERE LOWER(c.DirectedAgainst) LIKE '%admin%' AND c.Status != 'Solved'
        ORDER BY c.DateSubmitted DESC
    ''').fetchall()

    # Complaints forwarded by Admins for executive intervention
    escalated_tickets = conn.execute('''
        SELECT c.*, u.Username as StudentUsername 
        FROM Complaint c
        JOIN User u ON c.StudentID = u.ID
        WHERE c.Status = 'Forwarded' AND LOWER(c.DirectedAgainst) NOT LIKE '%admin%'
        ORDER BY c.DateSubmitted DESC
    ''').fetchall()

    # Past Solved executive complaints for sidebar
    past_complaints = conn.execute('''
        SELECT c.*, u.Username as StudentUsername 
        FROM Complaint c
        JOIN User u ON c.StudentID = u.ID
        WHERE c.Status = 'Solved'
        ORDER BY c.DateSubmitted DESC
    ''').fetchall()

    # Total solved count across university
    solved_count = conn.execute("SELECT COUNT(*) FROM Complaint WHERE Status = 'Solved'").fetchone()[0]

    # Highest Reason Tickets (Category aggregation)
    category_stats = conn.execute('''
        SELECT Category, COUNT(*) as count 
        FROM Complaint 
        GROUP BY Category 
        ORDER BY count DESC
    ''').fetchall()

    # Urgency Level distribution
    urgency_stats = conn.execute('''
        SELECT UrgencyLevel, COUNT(*) as count 
        FROM Complaint 
        GROUP BY UrgencyLevel 
        ORDER BY 
            CASE UrgencyLevel 
                WHEN 'Emergency' THEN 1 
                WHEN 'High' THEN 2 
                WHEN 'Medium' THEN 3 
                WHEN 'Low' THEN 4 
            END ASC
    ''').fetchall()

    conn.close()

    return render_template('hod.html', 
                           conflict_tickets=conflict_tickets, 
                           escalated_tickets=escalated_tickets,
                           past_complaints=past_complaints,
                           admin_ticket_count=admin_ticket_count,
                           conflict_count=len(conflict_tickets),
                           escalated_count=len(escalated_tickets),
                           solved_count=solved_count,
                           category_stats=category_stats,
                           urgency_stats=urgency_stats)


@app.route('/hod/solve/<int:complaint_id>', methods=['POST'])
def hod_solve(complaint_id):
    """
    HOD Executive Resolution Action:
    - Marks a conflict or escalated ticket as 'Solved' and attaches the HOD's official executive decision.
    """
    if session.get('role') != 'HOD':
        flash('Session expired or unauthorized. Please log in.', 'danger')
        return redirect(url_for('login'))

    comment = request.form.get('admin_comment', '').strip()

    conn = get_db()
    conn.execute("UPDATE Complaint SET Status = 'Solved', AdminComment = ? WHERE ID = ?", (comment, complaint_id))
    conn.commit()
    conn.close()

    flash(f'Complaint #{complaint_id} marked as Solved by HOD with response comment.', 'success')
    return redirect(url_for('hod_dashboard'))


@app.route('/hod/resolve-group/<int:group_id>', methods=['POST'])
def hod_resolve_group(group_id):
    """
    1-Click Cascading Group Resolution for HOD:
    - Solves all complaints relevant to each other in one single click directly from HOD Executive Desk.
    """
    if session.get('role') != 'HOD':
        flash('Session expired or unauthorized. Please log in.', 'danger')
        return redirect(url_for('login'))

    comment = request.form.get('admin_comment', f'1-Click resolved by HOD under Group #{group_id}').strip()

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE Complaint SET Status = 'Solved', AdminComment = ? WHERE GroupID = ? OR ID = ?", (comment, group_id, group_id))
    updated_count = cursor.rowcount
    conn.commit()
    conn.close()

    flash(f'HOD 1-Click Solved all {updated_count} related complaints in Group #{group_id}.', 'success')
    return redirect(url_for('hod_dashboard'))


# --------------------------------------------------------------------------------------
# 8. Live Database Inspector Route (Zero-extension web inspection for Vivas & Testing)
# --------------------------------------------------------------------------------------
@app.route('/db-view')
def db_view():
    """
    Live SQLite Database Viewer:
    - Renders a clean web-based table view of all records in the 'User' and 'Complaint' tables.
    - Useful during viva presentations and testing without requiring third-party database viewer software.
    """
    conn = get_db()
    users = conn.execute("SELECT ID, Username, Role, PasswordHash FROM User").fetchall()
    complaints = conn.execute("SELECT * FROM Complaint ORDER BY ID DESC").fetchall()
    conn.close()
    return render_template('db_view.html', users=users, complaints=complaints)


# --------------------------------------------------------------------------------------
# 9. Application Server Entrypoint
# --------------------------------------------------------------------------------------
if __name__ == '__main__':
    # Default to port 5001 (or environment override)
    port = int(os.environ.get('PORT', 5001))
    print(f"SUCS Prototype Flask running on http://0.0.0.0:{port}")
    # Run with debug mode enabled for automatic code reload during development
    app.run(host='0.0.0.0', port=port, debug=True)
