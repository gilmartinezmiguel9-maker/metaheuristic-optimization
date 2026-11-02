# metaheuristic-optimization
# PRACTIPLAN: Metaheuristic Optimization for Resource Allocation

## 📌 Executive Summary
An automated, Python-based scheduling and resource allocation engine developed to solve complex combinatorial optimization problems in academic operations. The system utilizes Constructive Greedy heuristics and Tabu Search to allocate human resources (students) and physical spaces (laboratories) under strict operational constraints, generating an automated BI dashboard for performance auditing.

## 🏢 Business Problem
Manual scheduling of laboratory sessions and physical spaces is a highly combinatorial, error-prone task that scales poorly. The challenge was to assign hundreds of students to specific lab groups and allocate those sessions to physical laboratories while strictly adhering to:
* **Hard Constraints:** Capacity limits, software/hardware availability, and overlapping schedule conflicts.
* **Soft Constraints:** Maintaining teacher continuity across consecutive sessions and balancing the load between groups.

## ⚙️ Methodology & Approach
The optimization was divided into sequential phases to ensure fast computation times and near-optimal assignments:
1. **Human Resource Allocation (Phase 2):** Implemented a Constructive Greedy algorithm with load-balancing logic to assign students to specific groups, minimizing immediate schedule incompatibilities.
2. **Space & Asset Allocation (Phase 3):** Engineered a **Tabu Search** metaheuristic with short-term memory and aspiration criteria. The algorithm explores local neighborhoods by swapping sessions between physical laboratories, avoiding cyclical traps and escaping local optima.
3. **Tech Stack:** Python (Standard Library) and `openpyxl` (for automated I/O operations and visual reporting).

## 📊 Impact & Results
* **Quality Improvement:** Achieved a **93.6% reduction** in objective function penalties (schedule clashes, capacity breaches, software mismatch) compared to a randomized baseline assignment.
* **Operational Efficiency:** The Tabu Search algorithm converges to near-optimal solutions in under 3 seconds.
* **Business Intelligence Integration:** The script automatically generates a formatted Excel Dashboard featuring capacity utilization KPIs, clash reports, and algorithmic evolution tracking, bridging the gap between mathematical optimization and actionable business intelligence.

## 🚀 How to Run
1. Clone the repository.
2. Install dependencies: `pip install -r requirements.txt` (Requires `openpyxl`).
3. Run the solver from the root directory: `python src/PRACTIPLAN_solver.py`.
4. The script will automatically read the input data and overwrite the Excel file with the optimized assignment and the visual dashboard.
