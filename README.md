# TeamForge — Intelligent Student Team Formation Platform

TeamForge is a complete, secure, responsive university collaboration platform that helps college students discover teammates, form balanced project teams using classical AI algorithms, participate in collegiate hackathons, manage project task boards, and communicate in real-time with zero persistent message storage for privacy.

---

## 1. Technology Stack & Decoupled Architecture

- **Frontend**: Pure **HTML5**, **CSS3**, and **JavaScript (ES6+)**:
  - `index.html` — Landing page with AI heuristics breakdown & hackathon spotlights
  - `login.html` & `register.html` — Authentication with 1-click pre-seeded demo login helpers
  - `dashboard.html` — Role-aware student dashboard with live metrics and team overview
  - `students.html` — Student directory with instant search, department filters & invite modals
  - `events.html` — Collegiate hackathons with registrations & event publisher
  - `teams.html` — My Teams & Projects with roster inspection & skill-gap alerts
  - `ai-generator.html` — Classical AI Team Generator with CSP solver, radar charts & trace console
  - `tasks.html` — Project Kanban board with drag & drop column status transitions
  - `chat.html` — Real-time ephemeral chat with zero persistent database message retention
  - `sessions.html` — Standups & mentor review scheduling with meeting links
  - `profile.html` & `settings.html` — Student profile, technical skill proficiencies & settings
  - `notifications.html` & `admin.html` — Notifications center, campus analytics & audit logs
  - Styling: Vanilla CSS3 design system (`assets/css/main.css`) + Bootstrap 5.3 + Bootstrap Icons
  - Logic: Pure JavaScript (`assets/js/api.js`, `assets/js/main.js`) communicating with backend via `fetch()`
- **Backend / Server-Side & Database**: **PHP 8.2+ & MySQL 8+ / MariaDB**:
  - Pure PHP API endpoints in `api/` (`auth.php`, `dashboard.php`, `students.php`, `events.php`, `teams.php`, `ai.php`, `tasks.php`, `chat.php`, `sessions.php`, `profile.php`, `notifications.php`, `admin.php`)
  - Normalized 3NF MySQL schema (`database/schema.sql`) with foreign keys, prepared statements, and transaction locks
  - 100% Deterministic Classical AI Algorithms in PHP (`ai/TeamForgeCspEngine.php` implementing CSP, Backtracking, MRV Heuristic, Forward Checking, Greedy Initial, and Hill Climbing)
- **Local Development**: XAMPP (Apache on port 80, MySQL on port 3306).
- **Strict Compliance**: Zero external AI APIs, zero machine learning, zero paid third-party dependencies.

---

## 2. Quick Demo Credentials (Pre-Seeded)

All demo accounts use password: **`Password123!`**

| Role | Account Name | Email | Primary Responsibilities |
|---|---|---|---|
| **Faculty / Admin** | Dr. Sarah Mitchell | `faculty@university.edu` | Publish events, verify student accounts & projects, view university metrics, audit trails |
| **Team Leader (TL)**| Aarav Sharma | `aarav.sharma@student.edu` | Lead team *ByteCrafters*, send invites, manage tasks, run AI Generator, schedule sessions |
| **Team Leader 2** | Priya Patel | `priya.patel@student.edu` | Lead team *AlgoKnights*, CSP and Discrete Optimization specialist |
| **Student (AI/Algo)**| anaya | `ananya.iyer@student.edu` | Algorithmic Problem Solving, Graph search, available for hackathon matching |
| **Student (Frontend)**| Rohan Verma | `rohan.verma@student.edu` | UI/UX & Responsive Web Developer |
| **Student (Database)**| Vikram Nair | `vikram.nair@student.edu` | MySQL Normalization & High-Throughput Databases |

---

## 3. Local Installation & XAMPP Setup

### Prerequisites
1. **XAMPP for Windows** installed with PHP 8.2+ and MySQL.
2. Start **Apache** and **MySQL** from the XAMPP Control Panel.

### Linking Project Directory to XAMPP
The project resides in `c:\AI & Web project`. To serve it through Apache, a Windows Directory Junction is linked to `C:\xampp\htdocs\teamforge`:
```cmd
mklink /J "C:\xampp\htdocs\teamforge" "c:\AI & Web project"
```
You can access TeamForge in your browser at:
**`http://localhost/teamforge/`**

### Database Setup
The database setup script can be run automatically:
```powershell
php "c:\AI & Web project\database\setup_db.php"
```
Or manually via **phpMyAdmin** (`http://localhost/phpmyadmin`):
1. Create database: `teamforge_db` (Collation: `utf8mb4_unicode_ci`).
2. Import `database/schema.sql`.
3. Import `database/seed_data.sql`.

---

## 4. Classical AI Algorithm Specifications & Implementation

All AI logic is implemented purely in PHP in `ai/TeamForgeCspEngine.php`:

### A. Constraint Satisfaction Problem (CSP)
- **Variables ($V_1, V_2, \dots, V_k$)**: Unassigned team role slots (e.g., $V_1$ = Team Leader, $V_2$ = Frontend, $V_3$ = Database, $V_4$ = QA/Algorithms).
- **Domains ($D(V_i)$)**: Eligible candidate students retrieved from MySQL matching project event eligibility criteria.
- **Hard Constraints**:
  1. **Team Size Constraint**: Exactly $k$ members (where $k \in [\text{min\_size}, \text{max\_size}]$).
  2. **Uniqueness Constraint**: $\forall i \neq j, V_i \neq V_j$ (No student can hold multiple slots in the same team).
  3. **Event Eligibility Constraint**: If an event ID is selected, candidate must be enrolled in that specific event.
  4. **Active Team Limit**: Disallows students exceeding active commitment thresholds.

### B. Minimum Remaining Values (MRV) Heuristic
- Implemented in `solveWithBacktrackingMRV()`.
- Prioritizes assigning the variable with the smallest remaining compatible candidate domain first.
- Greatly prunes the search tree branching factor by failing early on highly constrained technical roles (e.g., dedicated database architects or classical algorithm leads).

### C. Constraint Propagation / Forward Checking
- Implemented in `backtrack()`.
- After every tentative assignment of candidate $C$ to role $V_i$, $C$ is immediately removed from all other unassigned variables' domains ($D(V_j) \leftarrow D(V_j) \setminus \{C\}$).
- Detects domain wipeout ($|D(V_j)| = 0$) immediately and initiates backtracking before searching downstream dead ends.

### D. Greedy Initial Assignment
- Implemented in `generateGreedyInitial()`.
- Rapidly assembles a baseline feasible team by picking candidates with highest cumulative skill proficiencies and positive availability status.

### E. Hill Climbing Local Search (Member Swaps)
- Implemented in `optimizeWithHillClimbing()`.
- Explores the 1-opt neighborhood by iteratively swapping unpinned team members with candidates in the pool.
- Retains swaps that strictly increase the multi-objective balance score $S(\text{Team})$.
- Terminates upon reaching a local optimum where no neighbor swap yields a higher score.

### Multi-Objective Objective Function
$$S(\text{Team}) = W_{\text{skill}} \cdot \text{SkillCov} + W_{\text{role}} \cdot \text{RoleCov} + W_{\text{exp}} \cdot \text{ExpBal} + W_{\text{avail}} \cdot \text{AvailScore} + W_{\text{interest}} \cdot \text{Interest}$$

---

## 5. Security & Privacy Implementations

1. **Ephemeral Real-Time Chat (Section 15 & 19)**:
   - Chat messages are deliberately **NOT stored in MySQL**.
   - Messages are kept in a rolling ephemeral in-memory buffer that auto-expires.
   - Server-side authorization ensures only accepted team members (`isUserMemberOf()`) can access the channel.
2. **Prepared SQL Statements**: 100% of queries use PDO parameterized statements with disabled emulation to eliminate SQL injection.
3. **Session Hardening**: Session IDs are regenerated upon authentication (`session_regenerate_id(true)`) to prevent session fixation.
4. **CSRF Protection**: State-changing forms and AJAX calls validate a cryptographically secure token (`csrf_token()`).
5. **Output Escaping**: All dynamic rendering passes through `e()` (`htmlspecialchars(..., ENT_QUOTES, 'UTF-8')`).
6. **Bcrypt Password Hashing**: Passwords stored via PHP `password_hash()` and verified via `password_verify()`.

---

## 6. Directory Structure

```
c:\AI & Web project/
├── index.php                 # Modern Landing Page & Algorithm Showcase
├── login.php                 # Secure Authentication with Quick Demo Autofill
├── register.php              # Comprehensive Student Registration
├── logout.php                # Session Invalidation
├── dashboard.php             # Role-Aware Command Center
├── profile.php               # Student Profile, Skills & Verified Projects
├── notifications.php         # In-App Notification Center
├── settings.php              # Password & Privacy Preferences
├── 404.php                   # Error & Access Denied Page
├── config/
│   ├── constants.php         # Global constants, roles, and AI weights
│   └── db.php                # PDO Singleton Database Connection
├── includes/
│   ├── auth.php              # Role-Based Access Control (RBAC) & Sessions
│   ├── helpers.php           # Sanitization, CSRF, Flash alerts, Formatting
│   ├── header.php            # HTML Header & CSS Links
│   ├── navbar.php            # Top Navigation Bar with Unread Counter
│   ├── sidebar.php           # Categorized Navigation Sidebar
│   └── footer.php            # Footer & JavaScript Scripts
├── assets/
│   ├── css/main.css          # Design System (Navy, Violet accents, White cards)
│   └── js/main.js            # CSRF fetch helper, Mobile sidebar, Toasts
├── database/
│   ├── schema.sql            # Normalized MySQL Schema (19 tables)
│   ├── seed_data.sql         # Realistic University Demo Dataset
│   └── setup_db.php          # 1-Click Database Setup & Seeder Script
├── students/
│   └── index.php             # Student Directory & Manual Team Builder Tray
├── events/
│   ├── index.php             # College Events & Hackathons Directory
│   ├── view.php              # Event Details, Participant Roster & Enrollment
│   └── create.php            # Faculty Event Publisher
├── teams/
│   ├── create.php            # Create Team with Code Generation
│   ├── my_teams.php          # My Teams Hub & Join by Code
│   ├── view.php              # Team Management Dashboard & Leadership Transfer
│   ├── send_invite.php       # Transactional Invitation Sender
│   ├── respond_invite.php    # Transactional Accept/Reject Invitation Handler
│   └── smart_replace.php     # Surgical Skill Replacement Engine
├── ai/
│   ├── TeamForgeCspEngine.php # Core Classical AI Engine (CSP, MRV, Hill Climbing)
│   ├── generator.php         # AI Team Generator Interface & Proposals
│   ├── optimization_lab.php  # Visual Academic Execution Trace Lab
│   └── simulator.php         # What-If Hypothetical Scenario Simulator
├── tasks/
│   └── index.php             # Kanban Task Board & AI Task Allocation
├── chat/
│   ├── index.php             # Real-Time Team Group Chat Interface
│   ├── relay.php             # Ephemeral Zero-Storage PHP Relay
│   └── server.js             # Optional Node.js WebSocket Relay Server
├── sessions/
│   ├── index.php             # Scheduled Timed Sessions Management
│   └── room.php              # Live Timed Session Room with Attendance
└── admin/
    ├── index.php             # Faculty & Admin Center (Verifications & Events)
    └── audit_logs.php        # Immutable System Security Audit Logs
```
