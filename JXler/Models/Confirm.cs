using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JXler.Models
{
    public class Confirm
    {
        public ExecPtn execPtn { get; set; }
        public string Path { get; set; }
        public string Message { get; set; }
    }

    public enum ExecPtn
    {
        SameInput,
        SpecifyPath,
        SetIndividually,
        Unselected
    }
}
