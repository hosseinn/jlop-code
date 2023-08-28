
// LinearSolverTest.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* I took the example from "Formulation of an LP problem in lpsolve"
   at http://lpsolve.sourceforge.net/5.5/

   Using the linear solvers
     * basic linear solver: com.sun.star.comp.Calc.LpsolveSolver
     * CoinMP: com.sun.star.comp.Calc.CoinMPSolver
*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;




public class LinearSolverTest
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
    Calc.listSolvers();

    // specify the variable cells
    CellAddress xPos = Calc.getCellAddress(sheet, "B1");   // x
    CellAddress yPos = Calc.getCellAddress(sheet, "B2");   // y

    CellAddress[] vars = new CellAddress[]{ xPos, yPos };


    // specify profit equation
    Calc.setVal(sheet, "B3", "=143*B1 + 60*B2");   // P = 143x + 60y, maximize
    CellAddress profitEqu = Calc.getCellAddress(sheet, "B3");


    // set up equation formulae without inequalities
    Calc.setVal(sheet, "B4", "=120*B1 + 210*B2"); 
    Calc.setVal(sheet, "B5", "=110*B1 + 30*B2"); 
    Calc.setVal(sheet, "B6", "=B1 + B2"); 

    // create the constraints
    // constraints are equations and their inequalities
    SolverConstraint sc1 = Calc.makeConstraint(sheet, "B4","<=", 15000); 
         // 120x + 210y <= 15000
         // B4 is the address of the cell that is constrained; 

    SolverConstraint sc2 = Calc.makeConstraint(sheet, "B5", "<=", 4000);
         // 110x + 30y <= 4000
             
    SolverConstraint sc3 = Calc.makeConstraint(sheet, "B6", "<=", 75);
         // x + y <= 75
             
    // could also include x >= 0 and y >= 0

    SolverConstraint[] constraints = new SolverConstraint[]{ sc1, sc2, sc3 };


    // initialize the linear solver (CoinMP or basic linear)
    XSolver solver = Lo.createInstanceMCF(XSolver.class, 
                               "com.sun.star.comp.Calc.CoinMPSolver");
                              //   "com.sun.star.comp.Calc.LpsolveSolver");
    // System.out.println("Solver: " + solver);
    solver.setDocument(doc);
    solver.setObjective(profitEqu);
    solver.setVariables(vars);
    solver.setConstraints(constraints);
    solver.setMaximize(true);

    Props.showObjProps("Solver", solver);
    Props.setProperty(solver, "NonNegative", true);
              // restrict the search to the top-right quadrant of the graph

    // execute the solver
    solver.solve();
    Calc.solverReport(solver);
    //  Profit max == $6315.625, with x == 21.875 and y == 53.125

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



}  // end of LinearSolverTest class
