using System;
using System.Collections.Generic;
using System.Text;
using DynamicExpresso;

namespace El2Core.Services
{
    internal class CodeSolver
    {
        public void Solve()
        {
            var interpreter = new Interpreter();

            var context = new RuleContext
            {
                Temperatur = 28,
                Bewegung = false,
                Stunde = 14
            };

            bool result = interpreter.Eval<bool>(
                "Temperatur > 25",
                new Parameter("Temperatur", context.Temperatur)
            );

            Console.WriteLine(result);
        }

        private class RuleContext
        {
            public int Temperatur { get; set; }
            public bool Bewegung { get; set; }
            public int Stunde { get; set; }
        }
    }
}
