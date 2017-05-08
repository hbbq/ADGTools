using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADGTools.App
{

    interface ISecretInformation
    {

        string iotUser { get; }
        string iotPassword { get; }
        string ftpUser { get; }
        string ftpPassword { get; }

    }

}
