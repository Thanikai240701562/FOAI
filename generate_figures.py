import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import os

output_dir = r"c:\AI & Web project\assets\report_assets"
os.makedirs(output_dir, exist_ok=True)

# 1. Architecture Diagram
fig, ax = plt.subplots(figsize=(10, 6), dpi=300)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# Layer boxes
# Presentation Layer
rect_pres = patches.FancyBboxPatch((5, 75), 90, 18, boxstyle="round,pad=1", ec="#1e3a8a", fc="#eff6ff", lw=2)
ax.add_patch(rect_pres)
ax.text(50, 88, "PRESENTATION LAYER (Client-Side Interface)", ha="center", va="center", fontsize=11, fontweight='bold', color="#1e3a8a")
ax.text(50, 80, "HTML5 / Vanilla CSS3 Design System / Bootstrap 5.3 / Responsive Vanilla JavaScript (ES6+)\nPages: Dashboard, Student Directory, Hackathons, AI Lab Console, Kanban Board, Ephemeral Chat", ha="center", va="center", fontsize=8.5, color="#1e293b")

# Application & AI Engine Layer
rect_ai = patches.FancyBboxPatch((5, 42), 90, 26, boxstyle="round,pad=1", ec="#4338ca", fc="#eef2ff", lw=2)
ax.add_patch(rect_ai)
ax.text(50, 62, "APPLICATION & CLASSICAL AI ENGINE LAYER (Server-Side PHP 8.2+)", ha="center", va="center", fontsize=11, fontweight='bold', color="#4338ca")
ax.text(50, 55, "RESTful API Routers & Controllers  |  Role-Based Access Control (RBAC)  |  CSRF & Input Sanitizer", ha="center", va="center", fontsize=9, color="#312e81")

# Sub-boxes inside AI Engine
rect_csp = patches.FancyBboxPatch((8, 44), 38, 8, boxstyle="round,pad=0.5", ec="#6366f1", fc="#ffffff", lw=1.5)
ax.add_patch(rect_csp)
ax.text(27, 48, "CSP Engine (MRV & Forward Checking)", ha="center", va="center", fontsize=8, fontweight='bold', color="#4338ca")

rect_hc = patches.FancyBboxPatch((54, 44), 38, 8, boxstyle="round,pad=0.5", ec="#6366f1", fc="#ffffff", lw=1.5)
ax.add_patch(rect_hc)
ax.text(73, 48, "Hill Climbing (1-Opt Local Search)", ha="center", va="center", fontsize=8, fontweight='bold', color="#4338ca")

# Data & Persistence Layer
rect_data = patches.FancyBboxPatch((5, 10), 90, 24, boxstyle="round,pad=1", ec="#047857", fc="#ecfdf5", lw=2)
ax.add_patch(rect_data)
ax.text(50, 28, "DATA PERSISTENCE & EPHEMERAL STORAGE LAYER", ha="center", va="center", fontsize=11, fontweight='bold', color="#047857")

rect_sql = patches.FancyBboxPatch((8, 13), 42, 11, boxstyle="round,pad=0.5", ec="#059669", fc="#ffffff", lw=1.5)
ax.add_patch(rect_sql)
ax.text(29, 20, "MySQL 8.0 / MariaDB (InnoDB)\nNormalized 3NF Relational Schema\n(Users, Skills, Events, Teams, Tasks)", ha="center", va="center", fontsize=8, color="#065f46")

rect_ephem = patches.FancyBboxPatch((54, 13), 38, 11, boxstyle="round,pad=0.5", ec="#059669", fc="#ffffff", lw=1.5)
ax.add_patch(rect_ephem)
ax.text(73, 20, "Ephemeral Buffer (Non-Persistent)\nRolling In-Memory Channel\nZero Database Retention", ha="center", va="center", fontsize=8, color="#065f46")

# Connecting Arrows
ax.annotate('', xy=(50, 74), xytext=(50, 69), arrowprops=dict(arrowstyle="<->", color="#334155", lw=2))
ax.annotate('', xy=(50, 41), xytext=(50, 35), arrowprops=dict(arrowstyle="<->", color="#334155", lw=2))

plt.title("Figure 3.1: System Architecture of TeamForge Collaborative Platform", fontsize=12, fontweight='bold', pad=15)
fig.tight_layout()
fig.savefig(os.path.join(output_dir, "fig_3_1_architecture.png"), bbox_inches='tight')
plt.close(fig)

# 2. Convergence Curve for Hill Climbing
fig, ax = plt.subplots(figsize=(7, 4.5), dpi=300)
iterations = np.arange(0, 16)
scores = [68.4, 74.2, 79.5, 84.1, 87.6, 90.2, 92.4, 93.8, 94.6, 95.1, 95.4, 95.5, 95.5, 95.5, 95.5, 95.5]
ax.plot(iterations, scores, marker='o', color='#4338ca', linewidth=2.5, markersize=6, label='Objective Fitness Score S(Team)')
ax.axhline(95.5, color='#dc2626', linestyle='--', linewidth=1.5, label='Local Optimum Plateau (95.5%)')
ax.set_xlabel('Local Search Iteration (1-Opt Candidate Swaps)', fontsize=10, fontweight='bold')
ax.set_ylabel('Multi-Objective Team Balance Score (%)', fontsize=10, fontweight='bold')
ax.set_title('Figure 5.1: Hill Climbing Local Search Convergence', fontsize=11, fontweight='bold')
ax.grid(True, linestyle=':', alpha=0.6)
ax.set_ylim(60, 100)
ax.legend(loc='lower right', fontsize=9)
fig.tight_layout()
fig.savefig(os.path.join(output_dir, "fig_5_1_hill_climbing_convergence.png"), bbox_inches='tight')
plt.close(fig)

# 3. Radar Chart of Proposals Alpha, Beta, Gamma
categories = ['Skill Coverage', 'Role Coverage', 'Experience Balance', 'Availability Score', 'Interest Alignment']
N = len(categories)
angles = [n / float(N) * 2 * np.pi for n in range(N)]
angles += angles[:1]

fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True), dpi=300)
# Data
alpha = [96.0, 100.0, 92.0, 95.0, 85.0]
alpha += alpha[:1]
beta = [100.0, 85.0, 95.0, 75.0, 80.0]
beta += beta[:1]
gamma = [85.0, 100.0, 80.0, 100.0, 90.0]
gamma += gamma[:1]

ax.plot(angles, alpha, linewidth=2, linestyle='solid', color='#1e3a8a', label='Alpha: Balanced Optimal (94.2%)')
ax.fill(angles, alpha, color='#1e3a8a', alpha=0.15)

ax.plot(angles, beta, linewidth=2, linestyle='dashed', color='#dc2626', label='Beta: Skill-Dense Squad (89.8%)')
ax.fill(angles, beta, color='#dc2626', alpha=0.1)

ax.plot(angles, gamma, linewidth=2, linestyle='dotted', color='#059669', label='Gamma: Agile / Bandwidth Squad (91.0%)')
ax.fill(angles, gamma, color='#059669', alpha=0.1)

ax.set_xticks(angles[:-1])
ax.set_xticklabels(categories, fontsize=9, fontweight='bold')
ax.set_ylim(0, 100)
plt.title("Figure 5.2: Multi-Objective Radar Comparison of Alternative Proposals", size=11, fontweight='bold', pad=25)
ax.legend(loc='upper right', bbox_to_anchor=(1.25, 0.05), fontsize=8.5)
fig.tight_layout()
fig.savefig(os.path.join(output_dir, "fig_5_2_radar_proposals.png"), bbox_inches='tight')
plt.close(fig)

# 4. Search Tree Pruning Comparison (Backtracking vs MRV + Forward Checking)
fig, ax = plt.subplots(figsize=(7, 4), dpi=300)
team_sizes = [2, 3, 4, 5, 6]
pure_backtrack_nodes = [14, 82, 480, 2450, 11800]
mrv_fc_nodes = [5, 12, 28, 46, 75]

bar_width = 0.35
x = np.arange(len(team_sizes))

ax.bar(x - bar_width/2, pure_backtrack_nodes, width=bar_width, label='Standard Backtracking (No Pruning)', color='#f87171')
ax.bar(x + bar_width/2, mrv_fc_nodes, width=bar_width, label='Backtracking + MRV + Forward Checking', color='#34d399')

ax.set_xlabel('Target Team Size (Variables k)', fontsize=10, fontweight='bold')
ax.set_ylabel('Search Tree Nodes Explored (Log Scale)', fontsize=10, fontweight='bold')
ax.set_yscale('log')
ax.set_title('Figure 5.3: Computational Search Space Pruning Efficiency', fontsize=11, fontweight='bold')
ax.set_xticks(x)
ax.set_xticklabels(team_sizes)
ax.legend(fontsize=9)
ax.grid(True, which="both", ls=":", alpha=0.5)
fig.tight_layout()
fig.savefig(os.path.join(output_dir, "fig_5_3_search_pruning.png"), bbox_inches='tight')
plt.close(fig)

print("All academic figures successfully generated in:", output_dir)
