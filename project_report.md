# PRACTIPLAN GIOI-DOE: Metaheuristic Optimization for Laboratory Scheduling

**Author:** Miguel Gil Martínez  
**Academic Institution:** Universitat Politècnica de València (UPV)  
**Context:** Industrial Management Engineering — Industrial Organization Department  

---

## 1. Executive Summary

`PRACTIPLAN GIOI-DOE` is an automated, Python-based decision-support tool designed to solve complex combinatorial optimization problems in academic operations. The system bridges quantitative operations research and business execution by automating the allocation of human resources (students to laboratory groups) and physical assets (laboratory sessions to rooms) under strict operational constraints.

The solution integrates a two-stage approach: a **Constructive Greedy heuristic** with load-balancing for student assignment (Phase 2) and an advanced **Tabu Search metaheuristic** for physical laboratory allocation (Phase 3). Evaluated against a randomized baseline, the optimization engine achieves a **93.6% reduction in objective function penalties**, converging to near-optimal schedules in under 3 seconds. Furthermore, the system outputs an automated executive BI dashboard in Excel, turning raw mathematical outputs into actionable business intelligence.

---

## 2. Project Objectives

The primary objective is to build a robust, reproducible, and executable optimization pipeline that removes manual administrative bottlenecks. Specific functional goals include:
* **Automated Data Pipeline:** Ingesting raw academic datasets directly from Excel (`data/PRACTIPLAN DATOS.xlsx`) without intermediate manual manipulation.
* **Greedy Load-Balancing:** Assigning third-year Industrial Management Engineering (GIOI) students to practice groups while minimizing schedule clashes and equalizing group sizes.
* **Constrained Space Allocation:** Mapping DOE (Departamento de Organización de Empresas) laboratory sessions to physical infrastructure while respecting capacity ceilings, software prerequisites, time overlaps, and teacher continuity.
* **Visual Auditing:** Generating dynamic KPIs, performance charts, and conditional formatting directly within the workbook.
* **Experimental Benchmarking:** Quantifying algorithmic performance improvements against a randomized allocation baseline.

---

## 3. System Architecture & Repository Structure

The project follows a production-grade repository structure to ensure clean separation of concerns, reproducibility, and ease of review for technical recruiters and engineering managers:

```text
📁 metaheuristic-optimization/
├── 📁 data/            # Clean input dataset (PRACTIPLAN DATOS.xlsx)
├── 📁 docs/            # Technical documentation and this formal report
├── 📁 src/             # Core optimization engine (PRACTIPLAN_solver.py)
├── 📁 notebooks/       # Exploratory analysis and verification notebooks
├── 📄 .gitignore       # System and cache exclusion rules
├── 📄 requirements.txt # Project dependencies (openpyxl)
└── 📄 README.md        # Executive project summary and instructions
```

---

## 4. Methodology & Algorithmic Design

### 4.1 Phase 2: Constructive Greedy for Student Grouping
The student allocation phase processes enrollment records and student-specific time incompatibilities (e.g., afternoon restrictions). 
* **Heuristic Logic:** Students are sorted by academic load and constraint complexity. For each enrolled course, the algorithm evaluates available groups, selecting the one that minimizes immediate schedule overlaps.
* **Load-Balancing Tie-Breaker:** When multiple groups yield equivalent clash counts, the algorithm allocates the student to the group with the lowest current headcount, preventing severe capacity imbalances.

### 4.2 Phase 3: Tabu Search for Laboratory Scheduling
Allocating sessions to physical spaces represents a highly constrained combinatorial problem across six specialized laboratories (`0.01`, `0.02`, `0.03`, `S.01`, `S.02`, `S.03`) distributed across ground floor and basement levels.
* **Solution Representation:** A dictionary mapping each DOE session key `(course_code, start_time)` to a physical laboratory string.
* **Neighborhood Exploration:** Moves are defined by reallocating a single session to a different laboratory.
* **Tabu Memory & Aspiration:** A short-term attribute memory prohibits reversing recent reassignments for a defined tenure (`TABU_TENURE = 11`), preventing infinite cyclical loops. An aspiration criterion overrides tabu restrictions if a neighbor solution surpasses the global optimum.
* **Intensification:** The search dynamically prioritizes conflict-prone sessions to maximize marginal improvement efficiency.

---

## 5. Objective Function & Penalty Modeling

Solution quality is quantified via a multi-component objective function ($\text{FO}$). By default, weights are unified ($\text{weight} = 1$), allowing $\text{FO}$ to represent total operational infractions:

$$\text{FO}_{\text{Phase 3}} = w_1(\text{Lab Overlaps}) + w_2(\text{Capacity Excess}) + w_3(\text{Software Mismatch}) + w_4(\text{Teacher Discontinuity}) + w_5(\text{Unmapped GIOI Groups})$$

* **Lab Overlaps:** Two concurrent sessions assigned to the same physical room.
* **Capacity Excess:** Session headcount exceeding room maximum capacity ($20$ or $30$ seats).
* **Software Mismatch:** Sessions requiring dedicated computing infrastructure assigned to non-compliant labs.
* **Teacher Continuity:** Penalizes unnecessary room switches for instructors teaching consecutive sessions.
* **Unmapped GIOI Groups:** Tracks missing group mappings (maintained at $0$ in the final verified iteration).

---

## 6. Experimental Results

The solver was tested across multiple independent runs with randomized scenario generation. 

| Metric | Random Baseline | Tabu Search Optimized | Improvement (%) |
| :--- | :---: | :---: | :---: |
| **Objective Function ($\text{FO}$)** | 951.0 (avg) | 60.3 (avg) | **93.66%** |
| **Execution Time** | N/A | ~2.33 seconds | High Efficiency |
| **Unmapped GIOI Sessions** | Variable | 0 | 100% Resolution |

The experimental data confirms that while random allocation results in heavy penalties due to capacity and overlap violations, the Tabu Search engine systematically navigates the solution space to achieve stable, high-quality schedules instantly.

---

## 7. Business Value & Intelligence Integration

Beyond mathematical rigor, `PRACTIPLAN` delivers immediate value through automated reporting (`src/PRACTIPLAN_solver.py`):
* **Automated BI Dashboard:** Generates executive KPI blocks, comparative bar charts, and penalty breakdown pie charts.
* **Self-Contained Execution:** The script reads from `data/PRACTIPLAN DATOS.xlsx` and safely writes visual outputs back into the workbook without requiring external BI software.
* **Traceability:** The `Historial Tabu` sheet records optimization convergence step-by-step for auditing purposes.

---

## 8. Conclusions & Future Work

The project successfully demonstrates how metaheuristics can solve operational scheduling bottlenecks in academic and enterprise environments. The hybrid model (Greedy + Tabu Search) guarantees fast execution times and exceptional solution quality.

**Proposed Future Enhancements:**
1. Implementing a deterministic `RANDOM_SEED` parameter toggle directly from the CLI interface.
2. Benchmarking Tabu Search against alternative metaheuristics such as Simulated Annealing (SA) or Genetic Algorithms (GA).
3. Introducing multi-objective Pareto optimization to balance execution speed against teacher preference constraints dynamically.

---
## 9. References
* Romano, C. A. *PRACTIPLAN: Herramienta para el cálculo de horarios de prácticas y calendario de laboratorios DOE*. Universitat Politècnica de València.
* Romano, C. A. *Optimización Combinatoria: Heurísticas Constructivas, Búsqueda Local y Metaheurísticas*.