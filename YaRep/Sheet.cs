using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YaRep
{
    /// <summary>
    /// Лист отчета
    /// </summary>
    public abstract class Sheet
    {
        internal abstract object[,] GetArray();
    }
}
