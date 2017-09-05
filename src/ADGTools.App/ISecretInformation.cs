using System.Diagnostics.CodeAnalysis;

namespace ADGTools.App
{

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    interface ISecretInformation
    {

        string iotUser { get; }
        string iotPassword { get; }
        string ftpUser { get; }
        string ftpPassword { get; }

    }

}
