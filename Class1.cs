using ExcelDna.Integration;
using Google.OrTools.LinearSolver;

public static class MyFunctions
{
    [ExcelFunction(Description = "OR Tools")]
    public static double TestOrTools(double x_test)
    {
        // Create the linear solver with the GLOP backend.
        Solver solver = Solver.CreateSolver("GLOP");

        // Create the variables x and y.
        Variable x = solver.MakeNumVar(0.0, x_test, "x");
        Variable y = solver.MakeNumVar(0.0, 2.0, "y");

        Console.WriteLine("Number of variables = " + solver.NumVariables());

        // Create a linear constraint, 0 <= x + y <= 2.
        Constraint ct = solver.MakeConstraint(0.0, 2.0, "ct");
        ct.SetCoefficient(x, 1);
        ct.SetCoefficient(y, 1);

        Console.WriteLine("Number of constraints = " + solver.NumConstraints());

        // Create the objective function, 3 * x + y.
        Objective objective = solver.Objective();
        objective.SetCoefficient(x, 3);
        objective.SetCoefficient(y, 1);
        objective.SetMaximization();

        solver.Solve();

        Console.WriteLine("Solution:");
        Console.WriteLine("Objective value = " + solver.Objective().Value());
        Console.WriteLine("x = " + x.SolutionValue());
        Console.WriteLine("y = " + y.SolutionValue());

        return x.SolutionValue();
    }
}