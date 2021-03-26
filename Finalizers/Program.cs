using System;

namespace Finalizers
{
    class Program
    {
        static void Run()
        {
            using MixedClass mc = new MixedClass();

            mc.StartWriting();
        }

        static void Main()
        {
            Run();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
