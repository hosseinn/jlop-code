
// SolverTest2.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Uses a nonlinear solver (DEPS or SCO) on a nonlinear problem.
   This example is from http://en.wikipedia.org/wiki/Nonlinear_programming,
   the 2D example
     * DEPS (default): com.sun.star.comp.Calc.NLPSolver.DEPSSolverImpl
     * SCO: com.sun.star.comp.Calc.NLPSolver.SCOSolverImpl
   DEPS is used.

   Docs on the nonlinear solvers:
      - see https://wiki.openoffice.org/wiki/NLPSolver
*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;




public class SolverTest2
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

    CellAddress[] vars = new CellAddress[]{ xPos, yPos };


    // specify profit equation
    Calc.setVal(sheet, "B3", "=B1+B2");    // x + y
    CellAddress objective = Calc.getCellAddress(sheet, "B3");

    // set up equation formula without inequality (only one needed)
    Calc.setVal(sheet, "B4", "=B1*B1 + B2*B2");   // x^2 + y^2



    // create two constraints from one equation
    SolverConstraint sc1 =  Calc.makeConstraint(sheet, "B4", ">=", 1);
          // x^2 + y^2 >= 1
             
    SolverConstraint sc2 = Calc.makeConstraint(sheet, "B4", "<=", 2);
          // x^2 + y^2 <= 2
             
    SolverConstraint[] constraints = new SolverConstraint[]{ sc1, sc2 };


    // initialize the nonlinear solver (DEPS or SCO)
    XSolver solver = Lo.createInstanceMCF(XSolver.class,  // "com.sun.star.sheet.Solver");
                // uses "com.sun.star.comp.Calc.NLPSolver.DEPSSolverImpl" by default
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

    Props.setProperty(solver, "AssumeNonNegative", true);
             // restrict the search to the top-right quadrant of the graph

    // execute the solver
    solver.solve();
    Calc.solverReport(solver);
       //  Profit max == 2; x == 1, y == 1

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of SolverTest2 class
