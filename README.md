# Smart University Complaint System (SUCS)

> **A Role-Governed Web Platform for Campus Grievance Redressal, Conflict-of-Interest Firewall & Automated Triage**  
> **Course:** Virtual University of Pakistan — CS619 Final Project  
> **Group ID:** `S26PROJECT4EE01`  
> **Supervisor:** Muhammad Raheel  

---

## 📌 Project Overview
The **Smart University Complaint System (SUCS)** replaces traditional physical suggestion boxes and unstructured email threads with a transparent, role-governed digital grievance platform. It eliminates lost tickets, prevents queue flooding, enforces institutional conflict-of-interest isolation, and provides automated issue clustering with 1-Click batch resolution.

---

## 🚀 Key Business Rules & Features

| Rule ID | Name | Operational Logic & Implementation |
| :---: | :--- | :--- |
| **BR-01** | **Monthly Quota Limit** | Enforces a strict maximum of **5 complaints per student per calendar month** to eliminate frivolous spam and preserve administrative triage focus. |
| **BR-02** | **Conflict-of-Interest (COI) Firewall** | Grievances directed against **Administration** bypass the standard Admin triage queue completely and route exclusively to the **Head of Department (HOD)** desk to eliminate retaliation and bias. |
| **BR-03** | **Emergency Keyword Priority** | Description text is scanned for hazard terms (`urgent`, `fire`, `hazard`, `shock`, `leak`, `broken wire`, `emergency`). If found, priority auto-escalates to **High (P1)**. |
| **BR-04** | **Intelligent Auto-Grouping** | Correlated grievances sharing identical **Category** and **Target** are automatically clustered under a shared `group_id` token. |
| **BR-05** | **1-Click Batch Resolution** | Admin and HOD officers can resolve all correlated complaints in a cluster simultaneously with a single click and cascading audit note. |
| **BR-06** | **Mandatory Progress Remarks** | Status transitions to **In-Progress** strictly require an explicit investigative note before being saved. |

---

## 👥 Role-Based Portals

1. **Student Portal (`/student` & `/student-complaints`)**
   - Live monthly quota counter (X / 5 used).
   - Category and Target dropdown taxonomy pickers.
   - Real-time grievance tracker with color-coded status badges (`Pending`, `In-Progress`, `Resolved`) and official investigative remarks.

2. **Administrator Triage Desk (`/admin`)**
   - Active grievance triage queue with Category, Priority, and Status filters.
   - Clustered complaint cards with **1-Click Batch Resolution**.
   - Modal dialogs requiring mandatory remarks when moving tickets to `In-Progress`.

3. **HOD Executive Console (`/hod`)**
   - Campus-wide complaint statistics and SLA monitoring.
   - Exclusive **Conflict-of-Interest (COI)** tab displaying grievances lodged against Administration staff.
   - Executive override and batch resolution authority.

4. **In-Browser Database Inspector (`/db-view`)**
   - Built-in administrative database viewer displaying live tables (`User`, `Complaint`), record counts, and schema definitions without third-party tools.

---

## 🛠️ Technology Stack

- **Backend:** Python 3.9+ with Flask microframework
- **Database:** SQLite 3 relational database (ACID compliant, parameterized SQL)
- **Security:** Werkzeug Security (salted PBKDF2-SHA256 password hashing, CSRF session protection)
- **Frontend:** Jinja2 Template Engine + Tailwind CSS (responsive component styling)
- **Quality Assurance:** Pytest (100% automated test pass rate across 9 core integration scenarios)

---

## ⚙️ Quick Start Guide

### 1. Clone the Repository
```bash
git clone https://github.com/magsidev/UniversityComplaintManagement.git
cd UniversityComplaintManagement
```

### 2. Set Up Virtual Environment & Dependencies
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Run the Application
```bash
python app.py
```
Open your web browser and navigate to:
```
http://127.0.0.1:5001
```

### 4. Demo Login Credentials
| Role | Username | Password |
| :--- | :--- | :--- |
| **Student** | `student1` | `student123` |
| **Student** | `student2` | `student123` |
| **Administrator** | `admin1` | `admin123` |
| **Head of Department (HOD)** | `hod1` | `hod123` |

---

## 🧪 Automated Testing
Run the comprehensive Pytest integration suite:
```bash
pytest tests/test_flask_app.py -v
```
All **9 integration tests (TC-01 through TC-09)** pass with 100% verification coverage.

---

## 📄 Project Deliverables
- 📄 **Final Project Report:** `SUCS_Final_Project_Report.docx` (Complete academic report following VU CS619 guidelines with all embedded diagrams, tables, and viva guide).
- 📊 **Final Project Presentation:** `SUCS_VU_Template_Presentation.pptx` (PowerPoint presentation following official VU template with all embedded SRS diagrams and screenshots).
