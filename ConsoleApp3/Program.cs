using ExcelLibrary;
using System;
using System.Linq;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            Class1 class1 = new Class1();
            class1.Method();


        }

        static private string[] GetPosition(string Position)
        {
            var Index = GetAlphabet(Position);
            return new string[] { Position.Substring(0, Index ), Position.Substring(Index) };
        }

        static private int GetAlphabet(string Text)
        {
            var Sym = Text.First(i => int.TryParse(i.ToString(), out int result));
            var Index = Text.IndexOf(Sym.ToString());
            return Index;
        }
    }
}
