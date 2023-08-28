
// SolverTest.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Uses a nonlinear solver (DEPS or SCO) on a linear problem with
   three variables.
     * DEPS (default): com.sun.star.comp.Calc.NLPSolver.DEPSSolverImpl
     * SCO: com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl
   SCO is used.

   Or we could use one of the linear solvers.

   Docs on the nonlinear solvers:
      - see https://wiki.openoffice.org/wiki/NLPSolver
*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;




public class SolverTest
{

  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    // GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    // specify the variable cells
    CellAddress xPos = Calc.getCellAddress(sheet, "B1");  // x
    CellAddress yPos = Calc.getCellAddress(sheet, "B2");  // y
    CellAddress zPos = Calc.getCellAddress(sheet, "B3");  // z

    CellAddress[] vars = new CellAddress[]{ xPos, yPos, zPos };


    // set up equation formula without inequality
    Calc.setVal(sheet, "B4", "=B1+B2-B3"); 
    CellAddress objective = Calc.getCellAddress(sheet, "B4");


    // create three constraints (using the 3 variables)
    SolverConstraint sc1 = Calc.makeConstraint(sheet, "B1", "<=", 6);
       // x <= 6
             
    SolverConstraint sc2 = Calc.makeConstraint(sheet, "B2", "<=", 8);
       // y <= 8
             
    SolverConstraint sc3 = Calc.makeConstraint(sheet, "B3", ">=", 4);
       // z >= 4
             
    SolverConstraint[] constraints = new SolverConstraint[]{ sc1, sc2, sc3 };


    // initialize the nonlinear solver (SCO)
    XSolver solver = Lo.createInstanceMCF(XSolver.class,  // "com.sun.star.sheet.Solver");
                // uses com.sun.star.comp.Calc.NLPSolver.DEPSSolverImpl by default
                   "com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl");

    // System.out.println("Solver: " + solver);
    solver.setDocument(doc);
    solver.setObjective(objective);
    solver.setVariables(vars);
    solver.setConstraints(constraints);
    solver.setMaximize(true);

    Props.showObjProps("Solver", solver);
    Props.setProperty(solver, "EnhancedSolverStatus", false);
              // switch off nonlinear dialog about current progress

    // execute the solver
    solver.solve();
    Calc.solverReport(solver);
        //  Profit max == 10; vars are very close to 6, 8, and 4, but off by 6-7 dps

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of SolverTest class
