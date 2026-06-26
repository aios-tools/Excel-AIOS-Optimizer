# AIOS-Optimizer: Complete User Manual

This manual explains how to configure and run optimization problems using AIOS-Optimizer. The system is designed to adapt to two levels of complexity: from simple mathematical problems resolved entirely with Excel formulas, to complex integrations with third-party engineering software (ETABS, SAP2000, Python) via VBA.

## 1. Ribbon Description (Optimizer Tab)


The Ribbon on the **AIOS Optimizer** tab is organized into four functional groups:

### Group 1: Problem Optimization
Contains the essential elements to define the optimization problem on the active sheet:
- **Objective:** Opens the configuration window for the objective function to minimize or maximize.
- **Variables:** Allows selecting the cells that the optimizer will modify. These cells must be empty (without formulas or initial data), as the optimizer will overwrite their contents during execution.
- **Constraints:** Defines the physical constraints that must be satisfied.

### Group 2: Configuration
Manages general configuration and algorithm parameters:
- **Parameters:** Adjusts optimizer parameters.
- **Options:** Additional system configuration options.

### Group 3: Solve
Controls the execution of the optimization process:
- **Iterations, Generations, Population:** Fields to define the evolutionary configuration of the algorithm.
- **Run:** Starts the optimization using the native Excel recalculation engine and User Defined Functions (UDFs).
- **RunMacro:** Runs the optimization utilizing a VBA macro for external control.
- **SyncMacro:** Forces the execution of the linked macro to synchronize results.

### Group 4: Help
Support resources and software settings:
- **Documentation:** Opens this user manual.
- **Information:** Version and system information.
- **License:** Management of permissions and user license.

## 2. System Forms Description

The system uses several interface forms to facilitate the configuration of optimization components.

### 2.1. Objective Function Configuration 
Allows defining the cell containing the objective formula and whether you want to minimize or maximize.

### 2.2. Variables Configuration 
Manages the list of system variables, allowing you to add, edit, or delete variables.

### 2.3. Variables Editor
Child form to edit the details of a specific variable (name, type, limits).

### 2.4. Constraints Configuration 
Manages the list of physical constraints of the problem.

### 2.5. Constraints Editor
Child form to edit the details of a constraint.

### 2.6. Algorithm Parameters 
Allows selecting the optimization algorithm and adjusting its specific parameters.

### 2.7. Options 
Configuration of general system options.

### 2.8. Information
Displays software version information.

### 2.9. License Configuration
Manages user license configuration for the selected algorithm.

---

## 3. Basic Optimization (Excel-Only Sheets)

This is the fastest method, as it leverages Excel's native recalculation engine coupled with the optimizer's pure C DLL.

### Problem Preparation
All relationships between the variables and your objective function must be resolved using standard Excel formulas (`=SUM()`, `=IF()`, arithmetic calculations).

1. **Variables:** Independent cells that the optimizer will change (e.g., beam dimensions). These cells must be empty (no formulas or initial data), as the optimizer will overwrite their contents.
2. **Objective Function:** A cell with a formula that depends on your variables. This is what you want to Minimize or Maximize.
3. **Constraints (Optional):** Cells with formulas evaluating physical conditions (`<= 0` or `== 0`).

### Execution
Once the parameters (Iterations, Generations, Population) are configured in the **Optimizer** tab, click the **Run** button.
The evaluation will be virtually instantaneous, injecting thousands of values per second.

---

## 4. Advanced Optimization (Macros and External Software)

When you need to evaluate your designs using external software, the optimizer delegates execution to a VBA macro.

### Recommended Architecture (Multiple Sheets)
- **Dashboard (Active Sheet):** ALWAYS configure the optimizer from this sheet. Your Variables, Objective, and Constraints cells will be located here.
- **Temporary Data Sheets:** Create separate sheets where your macro will paste the raw data returned by the external program.
- **Connection:** In your Dashboard, use Excel formulas that read from the temporary data sheets (e.g., `='ETABS_Results'!C10`).

### VBA Code Configuration
You must define two elements in your VBA module:
1. **The Connection Function:** A mandatory function that returns the name of your main macro.
   ```vba
   Public Function AiosOptimizationMacro() As String
       AiosOptimizationMacro = "EvaluacionParaAIOS"
   End Function
   ```
2. **The Evaluation Subroutine:** Your main macro (e.g., `EvaluacionParaAIOS`) that reads variables from the sheet, runs your external software silently, and writes the returned results to the secondary sheets.

### Execution
Click the **RunMacro** button. The DLL will inject the variables and automatically execute your macro in each iteration.

---

## 5. Clarification: When to use "Run" and when "RunMacro" with VBA?

If you have VBA code in your project, the choice of button depends strictly on **how you are using it**:

- **Use "Run" if you use User Defined Functions (UDFs):**
  If your VBA code is a Function (created with `Function`) and you call it by typing a formula directly into a cell (e.g., `=MyVBAFunction(A1, B1)`), you must use **Run**. The optimizer's internal command forces Excel to recalculate all formulas, which automatically runs your VBA functions. *(Note: VBA functions are slower than native Excel functions).*

- **Use "RunMacro" if you use Complex Subroutines:**
  If your VBA code is a process (created with `Sub`) that cannot be written as a formula inside a cell because it controls external programs (ETABS), modifies cells in bulk, or changes formatting, you must use **RunMacro**.

---

## 6. Safe Interruption and Synchronization (SyncMacro Button)

Regardless of the method you choose, you can abort the execution at any time by holding down the **`ESC`** key.

- **Progress Retention:** The optimizer will abort safely and **write the best values found up to that point** into your Dashboard variable cells.
- **Synchronization in "Run":** If you were using Excel only (Run button), the algorithm recalculates the sheet one last time. Everything will remain perfectly synchronized.
- **Desynchronization in "RunMacro":** If you were using external software, the sheet **does not execute the macro one last time upon cancelling**. The printed physical results (e.g., forces) will belong to the last trial before cancelling, not to the displayed optimal variables.
- **Solution:** Click the **SyncMacro** button on the Ribbon. This will execute your macro using the already printed optimal variables, forcing ETABS/External Software to synchronize with your Excel.

*(**Additional Tip:** The **SyncMacro** button is also extremely useful **before** starting the optimizer. You can enter values manually in your variable cells and press SyncMacro to quickly test that communication with your external software works correctly, ensuring that the model yields valid results before launching a long optimization run).*
