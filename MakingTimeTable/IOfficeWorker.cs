using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MakingTimeTable
{
    public interface IOfficeWorker
    {
        void Make(params string[] urls);
    }
}
