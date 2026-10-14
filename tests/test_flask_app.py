"""
========================================================================================
Smart University Complaint System (SUCS v1.0) - Automated Pytest Test Suite
========================================================================================
Test Scope:
- Authentication & Session Verification (TC-01)
- 5-Complaint Monthly Quota Calculation (TC-10 to TC-12)
- Complaint Submission & Emergency Urgency Classification (TC-07 & TC-17)
- Automated Conflict-of-Interest Filter Verification (TC-18)

Execution Command:
    ./venv/bin/pytest tests/test_flask_app.py
========================================================================================
"""

import os
import pytest
from app import app, init_db, get_db, get_monthly_quota

# --------------------------------------------------------------------------------------
# Pytest Fixture: Test Client Setup
# --------------------------------------------------------------------------------------
@pytest.fixture
def client():
    """
    Test Client Fixture:
    - Sets app.config['TESTING'] = True to propagate exceptions and disable error trapping.
    - Yields a clean test client instance that can send HTTP requests without running a live server.
    - Ensures database schema is initialized before executing test scenarios.
    """
    app.config['TESTING'] = True
    with app.test_client() as client:
        init_db()
        yield client


# --------------------------------------------------------------------------------------
# Test Case 1: User Login & Session Establishment (TC-01)
# --------------------------------------------------------------------------------------
def test_prototype_login(client):
    """
    Verifies that a valid student credential payload successfully logs in:
    - Sends POST /login with username='student1' and password='password123'.
    - Follows redirect to ensure the response status is HTTP 200.
    - Asserts that the response HTML body contains 'Student Portal'.
    """
    res = client.post('/login', data={'username': 'student1', 'password': 'password123'}, follow_redirects=True)
    assert res.status_code == 200
    assert b'Student Portal' in res.data


# --------------------------------------------------------------------------------------
# Test Case 2: 5-Complaint Monthly Quota Calculation (TC-10 to TC-12)
# --------------------------------------------------------------------------------------
def test_prototype_quota(client):
    """
    Verifies the monthly quota business logic:
    - Queries the database to retrieve student1's user ID.
    - Invokes get_monthly_quota(student_id).
    - Asserts total limit is strictly 5.
    - Asserts remaining quota formula: remaining == limit - used.
    - Asserts remaining quota is constrained between 0 and 5.
    """
    conn = get_db()
    student = conn.execute("SELECT ID FROM User WHERE Username = 'student1'").fetchone()
    conn.close()

    quota = get_monthly_quota(student['ID'])
    assert quota['limit'] == 5
    assert quota['remaining'] == max(0, quota['limit'] - quota['used'])
    assert 0 <= quota['remaining'] <= 5


# --------------------------------------------------------------------------------------
# Test Case 3: Complaint Submission & Emergency Priority Tagging (TC-07 & TC-17)
# --------------------------------------------------------------------------------------
def test_prototype_submit_complaint(client):
    """
    Verifies complaint submission with Emergency flag:
    - Simulates an active authenticated student session using session_transaction().
    - Sends POST /submit-complaint with is_emergency='on'.
    - Verifies HTTP 200 response upon redirect to student dashboard.
    - Asserts that 'Emergency' priority badge is present in the rendered HTML.
    """
    with client.session_transaction() as sess:
        sess['user_id'] = 1
        sess['username'] = 'student1'
        sess['role'] = 'Student'

    res = client.post('/submit-complaint', data={
        'category': 'Facilities',
        'directed_against': 'IT Support Desk',
        'description': 'The multimedia projector in CS-204 displays flickering red lines during class.',
        'is_emergency': 'on'
    }, follow_redirects=True)

    assert res.status_code == 200
    assert b'Emergency' in res.data


# --------------------------------------------------------------------------------------
# Test Case 4: Automated Conflict-of-Interest Filter (TC-18)
# --------------------------------------------------------------------------------------
def test_prototype_admin_queue_and_conflict_filter(client):
    """
    Verifies that complaints directed against Admin personnel are filtered out from the Admin queue:
    - Inserts a sample complaint with DirectedAgainst = 'Prof. Asif Khan (Admin)'.
    - Simulates an active Admin session (role='Admin').
    - Requests GET /admin.
    - Asserts HTTP 200 OK.
    - Asserts that 'Prof. Asif Khan (Admin)' does NOT appear anywhere in the Admin's queue HTML.
    """
    # 1. Insert complaint directed against Admin personnel
    conn = get_db()
    conn.execute('''
        INSERT INTO Complaint (StudentID, Category, DirectedAgainst, Description, UrgencyLevel, Status)
        VALUES (2, 'Administrative', 'Prof. Asif Khan (Admin)', 'Testing conflict routing logic', 'High', 'Pending')
    ''')
    conn.commit()
    conn.close()

    # 2. Simulate Admin login
    with client.session_transaction() as sess:
        sess['user_id'] = 3
        sess['username'] = 'admin1'
        sess['role'] = 'Admin'

    # 3. Fetch Admin queue and verify conflict-of-interest exclusion
    res = client.get('/admin')
    assert res.status_code == 200
    # Conflict complaint directed "Against Admin" must NOT appear in Admin view
    assert b'Prof. Asif Khan (Admin)' not in res.data


# --------------------------------------------------------------------------------------
# Test Case 5: Auto-Combining Similar Complaints & 1-Click Admin Group Resolution
# --------------------------------------------------------------------------------------
def test_auto_combining_and_admin_one_click_solve(client):
    """
    Verifies that complaints with the same Category and Target are auto-combined,
    and can be resolved simultaneously with 1-click by the Admin:
    - Submits 2 complaints sharing Category='Facilities' and DirectedAgainst='Campus Maintenance'.
    - Verifies both complaints receive the same GroupID.
    - Admin submits 1-click group resolution.
    - Asserts both complaints are marked 'Solved' with the group resolution comment.
    """
    with client.session_transaction() as sess:
        sess['user_id'] = 1
        sess['username'] = 'student1'
        sess['role'] = 'Student'

    # Submit ticket 1
    client.post('/submit-complaint', data={
        'category': 'Facilities',
        'directed_against': 'Campus Maintenance',
        'description': 'Water dispenser in CS Block is leaking on the floor.',
    }, follow_redirects=True)

    # Submit ticket 2 (matching category and target)
    client.post('/submit-complaint', data={
        'category': 'Facilities',
        'directed_against': 'Campus Maintenance',
        'description': 'Water cooler on 2nd floor has no cooling.',
    }, follow_redirects=True)

    conn = get_db()
    rows = conn.execute('''
        SELECT ID, GroupID, Status FROM Complaint 
        WHERE Category = 'Facilities' AND DirectedAgainst = 'Campus Maintenance'
    ''').fetchall()
    conn.close()

    assert len(rows) >= 2
    # Verify auto-grouping linked them under a shared Master GroupID
    group_id = rows[0]['GroupID']
    assert group_id is not None
    assert rows[1]['GroupID'] == group_id

    # Admin 1-Click Resolves the whole group
    with client.session_transaction() as sess:
        sess['user_id'] = 3
        sess['username'] = 'admin1'
        sess['role'] = 'Admin'

    res = client.post(f'/admin/resolve-group/{group_id}', data={
        'admin_comment': 'Plumbing team replaced all valves across CS Block.'
    }, follow_redirects=True)
    assert res.status_code == 200

    conn = get_db()
    updated_rows = conn.execute('''
        SELECT Status, AdminComment FROM Complaint WHERE GroupID = ?
    ''', (group_id,)).fetchall()
    conn.close()

    for r in updated_rows:
        assert r['Status'] == 'Solved'
        assert 'Plumbing team replaced all valves' in r['AdminComment']


# --------------------------------------------------------------------------------------
# Test Case 6: HOD 1-Click Group Resolution for Relevant Issues
# --------------------------------------------------------------------------------------
def test_hod_one_click_group_resolve(client):
    """
    Verifies that the HOD can resolve all relevant complaints in a group with 1-click:
    - Inserts 2 conflict-of-interest complaints directed against 'Admin Staff'.
    - Links them under a GroupID.
    - HOD triggers 1-click group resolution (/hod/resolve-group/<group_id>).
    - Asserts all tickets under that group are closed as 'Solved'.
    """
    conn = get_db()
    c1 = conn.execute('''
        INSERT INTO Complaint (StudentID, Category, DirectedAgainst, Description, UrgencyLevel, Status)
        VALUES (1, 'Administrative', 'Central Admin Staff', 'Fee voucher delayed', 'Medium', 'Pending')
    ''').lastrowid
    c2 = conn.execute('''
        INSERT INTO Complaint (StudentID, Category, DirectedAgainst, Description, UrgencyLevel, Status)
        VALUES (2, 'Administrative', 'Central Admin Staff', 'Fee challan error', 'Medium', 'Pending')
    ''').lastrowid
    # Link them in a group
    conn.execute("UPDATE Complaint SET GroupID = ? WHERE ID IN (?, ?)", (c1, c1, c2))
    conn.commit()
    conn.close()

    # HOD session
    with client.session_transaction() as sess:
        sess['user_id'] = 4
        sess['username'] = 'hod1'
        sess['role'] = 'HOD'

    res = client.post(f'/hod/resolve-group/{c1}', data={
        'admin_comment': 'Accounts portal system bug patched by IT.'
    }, follow_redirects=True)
    assert res.status_code == 200

    conn = get_db()
    solved_tickets = conn.execute("SELECT Status, AdminComment FROM Complaint WHERE GroupID = ?", (c1,)).fetchall()
    conn.close()

    assert len(solved_tickets) == 2
    for t in solved_tickets:
        assert t['Status'] == 'Solved'
        assert 'Accounts portal system bug patched' in t['AdminComment']


# --------------------------------------------------------------------------------------
# Test Case 7: In-Progress Status with Mandatory Comment Enforcement
# --------------------------------------------------------------------------------------
def test_in_progress_status_and_mandatory_comment(client):
    """
    Verifies that:
    - Setting In-Progress status requires a mandatory comment.
    - If comment is missing, the status update is rejected.
    - If comment is provided, Status becomes 'In-Progress'.
    """
    conn = get_db()
    cid = conn.execute('''
        INSERT INTO Complaint (StudentID, Category, DirectedAgainst, Description, UrgencyLevel, Status)
        VALUES (1, 'Academic', 'Exam Cell', 'Grades discrepancy in Physics mid-term', 'Low', 'Pending')
    ''').lastrowid
    conn.commit()
    conn.close()

    with client.session_transaction() as sess:
        sess['user_id'] = 3
        sess['username'] = 'admin1'
        sess['role'] = 'Admin'

    # Attempt without comment -> must reject
    res_fail = client.post(f'/admin/action/{cid}/in_progress', data={'admin_comment': ''}, follow_redirects=True)
    assert b'Comment is mandatory' in res_fail.data

    # Submit with mandatory comment
    res_success = client.post(f'/admin/action/{cid}/in_progress', data={
        'admin_comment': 'Forwarded exam sheet to Dr. Tariq for grade re-check.'
    }, follow_redirects=True)
    assert res_success.status_code == 200

    conn = get_db()
    ticket = conn.execute("SELECT Status, AdminComment FROM Complaint WHERE ID = ?", (cid,)).fetchone()
    conn.close()

    assert ticket['Status'] == 'In-Progress'
    assert 'Dr. Tariq' in ticket['AdminComment']


# --------------------------------------------------------------------------------------
# Test Case 8: Custom Category and Directed Against Options ("Others")
# --------------------------------------------------------------------------------------
def test_custom_others_options(client):
    """
    Verifies that selecting 'Others' allows the student to specify custom
    category and target names which are stored properly in the database.
    """
    conn = get_db()
    # Reset complaints for student2 to ensure clean monthly quota
    conn.execute("DELETE FROM Complaint WHERE StudentID = 2")
    conn.commit()
    conn.close()

    with client.session_transaction() as sess:
        sess['user_id'] = 2
        sess['username'] = 'student2'
        sess['role'] = 'Student'

    res = client.post('/submit-complaint', data={
        'category': 'Others',
        'custom_category': 'Robotics Lab Equipment',
        'directed_against': 'Others',
        'custom_directed_against': 'Robotics Lab Assistant',
        'description': 'Soldering irons and 3D printer are malfunctioning.',
    }, follow_redirects=True)

    assert res.status_code == 200

    conn = get_db()
    row = conn.execute('''
        SELECT Category, DirectedAgainst FROM Complaint 
        WHERE Description LIKE '%Soldering irons%'
    ''').fetchone()
    conn.close()

    assert row is not None
    assert row['Category'] == 'Others: Robotics Lab Equipment'
    assert row['DirectedAgainst'] == 'Others: Robotics Lab Assistant'


# --------------------------------------------------------------------------------------
# Test Case 9: Dedicated Student Complaints History View
# --------------------------------------------------------------------------------------
def test_dedicated_student_complaints_page(client):
    """
    Verifies the dedicated student complaints route (/student/complaints).
    """
    with client.session_transaction() as sess:
        sess['user_id'] = 1
        sess['username'] = 'student1'
        sess['role'] = 'Student'

    res = client.get('/student/complaints')
    assert res.status_code == 200
    assert b'My Submitted Complaints' in res.data


