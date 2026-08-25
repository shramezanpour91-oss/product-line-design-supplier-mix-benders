# Product Line Design with Many Components and Supplier Mix: A Tailored Benders Approach

This repository contains the computational implementations accompanying the research paper.

The study integrates **product-line design**, **customer choice**, and **capacitated multi-supplier procurement** in a single profit-maximization framework. The resulting mixed-integer linear programming (MILP) model becomes computationally challenging when the number of components is large. To improve scalability, the paper develops two tailored Benders decomposition formulations with closed-form subproblem calculations and greedy initialization.

The repository provides implementations of:

- the **Compact Model (CM)**;
- the **first Benders decomposition (BD1)**, with customer-choice variables in the master problem;
- the **second Benders decomposition (BD2)**, with customer-choice decisions transferred to the subproblem;
- the **Forward-Addition Greedy Heuristic (FAGH)** used to generate a strong initial solution/lower bound.

---

## 1. Research problem

A manufacturer must decide which product versions to offer while simultaneously accounting for customer preferences and component sourcing.

Each candidate product version:

- generates revenue when selected by a customer segment;
- requires a subset of components;
- may compete with other versions under a **First-Choice** customer-choice rule.

Each component can be purchased from multiple suppliers with:

- supplier-specific unit procurement costs;
- supplier-specific capacity limits.

The model jointly determines:

1. which product versions to include in the product line;
2. which version is chosen by each customer segment;
3. the resulting component requirements;
4. how component demand is allocated across available suppliers.

The objective is to maximize:

**total customer revenue − total component procurement cost**

subject to product-line, customer-choice, supplier-capacity, and sourcing constraints.

Because component demand depends on customer choices, and customer choices depend on the selected product line, product design and procurement decisions are tightly coupled.

---

## 2. Methodological contribution

The compact MILP can become difficult to solve when the number of components reaches the thousands. The paper therefore develops two tailored Benders reformulations.

### Compact Model (CM)

The compact formulation includes the product-selection, customer-choice, and supplier-allocation decisions in a single MILP.

Main decision variables include:

- `x_i`: 1 if product version `i` is offered;
- `y_i,k`: 1 if customer segment `k` selects version `i`;
- `Q_p,s`: quantity of component `p` purchased from supplier `s`.

This formulation is useful as the benchmark for evaluating the decomposition approaches.

### BD1 — Customer choices in the master problem

BD1 keeps product-selection and customer-choice decisions in the master problem and separates the component-procurement recourse structure.

For a fixed customer-choice solution, total demand for component `p` is

\[
A_p=\sum_k\sum_i \lambda_k a_{ip}y_{ik}.
\]

Suppliers are ordered by increasing unit procurement cost. The procurement subproblem can therefore be evaluated through the pivotal supplier and cumulative supplier capacities.

The implementation exploits this structure to construct Benders optimality cuts **without repeatedly solving a separate LP subproblem**.

### BD2 — Customer choices in the subproblem

BD2 uses a sparser master problem centered primarily on product-selection decisions and transfers the customer-choice structure to the subproblem.

The code computes the required dual quantities directly through closed-form expressions and generates Benders cuts from these values. No explicit CPLEX LP subproblem is built or solved inside the callback.

This formulation is particularly attractive as the number of candidate product versions increases because the `n × K` customer-choice structure is moved away from the master problem.

---

## 3. Greedy initialization

The uploaded implementations use the **Forward-Addition Greedy Heuristic (FAGH)** to construct an initial feasible product set.

Starting from an empty product line, FAGH repeatedly:

1. considers each unselected candidate product;
2. temporarily adds that product to the current product line;
3. evaluates customer choices, component requirements, procurement cost, and profit;
4. permanently selects the candidate producing the largest improvement;
5. stops when no improving addition exists or the product-line limit `V` is reached.

The resulting solution is used as an initial lower bound and/or MIP start.

The paper also studies swap-local-search and revenue-based greedy alternatives, but the three solver files included here focus on the forward-addition initialization.

---

## 4. Repository structure

A recommended GitHub structure is:

```text
product-line-design-supplier-mix-benders/
│
├── README.md
├── code/
│   ├── CM.py
│   ├── BD1.py
│   └── BD2.py
│
├── data/
│   └── Main Dataset.xlsx
```

Suggested mapping of the supplied scripts:

| File | Purpose |
|---|---|
| `CM.py` | Direct compact MILP solved with the CPLEX Python API |
| `BD1.py` | Benders decomposition with customer-choice variables in the master |
| `BD2.py` | Closed-form BD2 implementation with customer choices handled in the subproblem |

If different filenames are used in the repository, replace the commands below accordingly.

---

## 5. Input data

The scripts read the instance data from a single Excel workbook.

The workbook must contain the following sheets **with exactly these names**:

| Excel sheet | Model parameter | Description |
|---|---|---|
| `Lambda^k` | \(\lambda_k\) | Number/weight of customers in segment `k` |
| `Pi_i` | \(\pi_i\) | Revenue/selling-price coefficient of product version `i` |
| `a_ip` | \(a_{ip}\) | Component requirement: quantity of component `p` required by version `i` |
| `Sigma_i^k` | \(\sigma_i^k\) | Preference rank of version `i` for customer segment `k` |
| `r_p^s` | \(r_{ps}\) | Unit procurement cost of component `p` from supplier `s` |
| `Cap_p^s` | \(cap_{ps}\) | Available capacity of component `p` from supplier `s` |

The code uses product index `0` for the no-purchase/outside option where applicable.

### Artificial supplier

To preserve feasibility when the capacities of the real suppliers are insufficient, the implementations add an artificial supplier internally.

The current scripts use:

```python
ARTIFICIAL_SUPPLIER_COST = 1000.0
ARTIFICIAL_SUPPLIER_CAP = 50000.0
```

The artificial supplier represents an expensive fallback source and is not part of the original supplier data in the Excel workbook.

---

## 6. Experimental instance sizes

The computational study considers the following parameter grid:

```python
N_values = [10, 20, 40]
P_values = [1000, 2000, 4000, 8000, 16000]
K_values = [200, 400]
V_values = [3, 5]
S_values = [2, 4]
```

where:

- `n` = number of candidate product versions;
- `P` = number of components;
- `K` = number of customer segments;
- `V` = maximum number of product versions that can be selected;
- `S` = number of real suppliers per component.

The full Cartesian product gives **120 computational instances**.

The scripts are currently configured with a smaller active parameter set for convenient testing, while the full experimental grid is included in the user-settings section as commented code.

---

## 7. Data generation and scaling

In the paper, the test parameters are generated from synthetic distributions. In particular:

- customer-segment sizes \(\lambda_k\) are sampled from `[1, 10]`;
- supplier unit costs \(r_{ps}\) are sampled from `[100, 200]`;
- product preference rankings are generated from a uniform distribution;
- \(a_{ip}\) is binary;
- product revenues and supplier capacities are generated using ranges adjusted to instance size.

The supplied scripts read a base Excel workbook and apply automatic scaling when subsets of different sizes are solved.

In the current implementations:

- product revenue is scaled as the number of components increases;
- supplier capacity is adjusted according to the number of real suppliers;
- capacity is also adjusted when the number of customer segments changes.

Researchers wishing to reproduce a specific table from the paper should therefore preserve both the source workbook and the scaling rules in the scripts.

---

## 8. Requirements

### Software

- Python 3
- IBM ILOG CPLEX Optimization Studio with a valid license
- CPLEX Python API
- pandas
- XlsxWriter
- an Excel-compatible application for inspecting output files

Install the open Python dependencies with:

```bash
pip install pandas xlsxwriter
```

The `cplex` Python package should be installed using the installation method corresponding to your IBM ILOG CPLEX distribution and Python environment.

> **Important:** CPLEX is proprietary software. Running these scripts requires access to a valid CPLEX installation/license.

The scripts were developed using the CPLEX Python API. The paper reports experiments run in Python within the Spyder/Anaconda environment on Windows 11 Pro.

---

## 9. Configuration

Each script contains a **USER SETTINGS** section near the top.

Before running a script, update the Excel path:

```python
address = r"C:\path\to\your\Main Dataset.xlsx"
```

For example:

```python
address = r"C:\Users\YourName\Documents\product-line-design-supplier-mix-benders\data\Main Dataset.xlsx"
```

Then select the instance sizes to solve:

```python
N_values = [10]
P_values = [1000]
K_values = [200]
V_values = [3]
S_values = [2]
```

To execute the full experimental grid:

```python
N_values = [10, 20, 40]
P_values = [1000, 2000, 4000, 8000, 16000]
K_values = [200, 400]
V_values = [3, 5]
S_values = [2, 4]
```

The Benders scripts use:

```python
TIME_LIMIT = 3600
THREADS = 1
```

so each selected instance has a one-hour time limit.

---

## 10. Running the models

From the repository root, run one formulation at a time.

### Compact model

```bash
python code/CM.py
```

### BD1

```bash
python code/BD1.py
```

### BD2

```bash
python code/BD2.py
```

The scripts automatically iterate over the Cartesian product of the parameter values specified in their user-settings sections.

For initial verification, it is recommended to run a small instance such as:

```python
N_values = [10]
P_values = [1000]
K_values = [200]
V_values = [3]
S_values = [2]
```

before launching the full computational experiment.

---

## 11. Output files

All implementations export detailed Excel results.

### Compact model

The compact implementation creates a `CPLEX_subset_results` directory next to the input workbook.

Typical output:

```text
CPLEX_subset_results/
├── CPLEX_solution_sample_1_n=10_P=1000_K=200_V=3_S=2.xlsx
└── CPLEX_all_subset_runs_summary.xlsx
```

Detailed workbooks contain sheets for:

- summary statistics;
- final offered products;
- greedy offered products;
- greedy steps;
- customer choices;
- component demand;
- complete and nonzero `X`, `Y`, and `Q` values.

### BD1

BD1 creates a `Benders_subset_results` directory next to the input workbook.

Typical output:

```text
Benders_subset_results/
├── Benders_solution_sample_1_n=10_P=1000_K=200_V=3_S=2.xlsx
└── Benders_all_subset_runs_summary.xlsx
```

Detailed workbooks include:

- run summary;
- final and greedy offered products;
- greedy steps;
- customer choices;
- component demand;
- complete and nonzero `x` and `y` values;
- Benders-related performance information.

### BD2

BD2 exports detailed closed-form results next to the input workbook, including files of the form:

```text
BD2_closed_form_beta_values_n10_P1000_K200_V3_S2_sample1.xlsx
BD2_closed_form_summary_all_runs.xlsx
```

The detailed BD2 workbook records, among other outputs:

- final product-selection decisions;
- customer choices;
- component requirements;
- pivotal-supplier information;
- closed-form `beta1`–`beta5` values;
- matrix/long-format beta tables;
- Benders cut information;
- run-level summary statistics.

---

## 12. Computational benchmark reported in the paper

The paper reports experiments performed on:

- Intel(R) Core(TM) Ultra 7 155U @ 1.70 GHz;
- 12 cores / 14 logical processors;
- 16 GB RAM;
- Windows 11 Pro;
- Python with the CPLEX Python API;
- Spyder IDE from the Anaconda distribution;
- 3600-second time limit per instance.

Reported solution times include the complete computational workflow, including data loading, model construction, initialization, and optimization.

Runtime results on another machine, CPLEX version, operating system, or solver configuration may differ.

---

## 13. Main computational findings

The computational study shows that the advantage of Benders decomposition grows as the number of components increases.

Key findings reported in the paper include:

- both BD1 and BD2 scale substantially better than the compact formulation on large-component instances;
- the Forward-Addition Greedy Heuristic provides an effective initialization strategy and achieved the shortest solution time in **34 of 48** large-instance heuristic comparisons;
- at `P = 16000`, the average computation-time improvement over the compact formulation is approximately **80% for BD1** and **69% for BD2**;
- for every large-scale instance in which the compact model failed to prove optimality within one hour, at least one of BD1 or BD2 obtained an optimal solution in less than **15% of the one-hour time limit**;
- BD1 performs particularly strongly for smaller and medium values of `n`, while BD2 becomes increasingly attractive as `n` grows because its master problem avoids the `n × K` customer-choice structure.

These results illustrate the value of exploiting the procurement structure analytically rather than repeatedly solving large LP subproblems inside the decomposition procedure.

---

## 14. Reproducing the paper's experiments

For reproducible computational comparisons:

1. use the same `Main Dataset.xlsx`;
2. preserve the Excel sheet names exactly;
3. use the full parameter grid shown above;
4. preserve the automatic scaling rules in each script;
5. retain the artificial-supplier parameters unless intentionally performing a sensitivity test;
6. use the same 3600-second time limit;
7. record the CPLEX version and hardware used;
8. run CM, BD1, and BD2 on the same instance definitions;
9. compare objective values before comparing solution times;
10. use the generated `summary_all_runs` workbooks for aggregate analysis.

Because solver runtimes depend on hardware, operating system, CPLEX version, solver settings, and numerical tolerances, exact runtimes should not be expected to be identical across computing environments.

---

## 15. Model notation

| Symbol | Meaning |
|---|---|
| \(i=0,\ldots,n\) | Product versions (`0` is the outside/no-purchase option) |
| \(p=1,\ldots,P\) | Components |
| \(k=1,\ldots,K\) | Customer segments |
| \(s\) | Suppliers |
| \(V\) | Maximum number of offered product versions |
| \(\lambda_k\) | Number/weight of customers in segment `k` |
| \(\pi_i\) | Revenue/selling-price coefficient of version `i` |
| \(a_{ip}\) | Quantity of component `p` required by version `i` |
| \(\sigma_i^k\) | Preference rank of version `i` for segment `k` |
| \(r_{ps}\) | Unit procurement cost of component `p` from supplier `s` |
| \(cap_{ps}\) | Supplier capacity for component `p` |
| \(x_i\) | Product-selection decision |
| \(y_{ik}\) | Customer-choice decision |
| \(Q_{ps}\) | Quantity purchased from supplier `s` for component `p` |
| \(A_p\) | Aggregate demand for component `p` |

---

## 17. Authors

**Shaghayegh Ramezanpour**  
**Laurent Alfandari**  
ESSEC Business School, France
