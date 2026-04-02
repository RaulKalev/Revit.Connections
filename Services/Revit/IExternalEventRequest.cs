using Autodesk.Revit.UI;

namespace Connections.Services.Revit
{
    /// <summary>
    /// Represents a request to be executed within a Revit External Event.
    /// </summary>
    public interface IExternalEventRequest
    {
        /// <summary>
        /// Executes the request within the Revit API context.
        /// </summary>
        /// <param name="app">The Revit UI Application.</param>
        void Execute(UIApplication app);
    }
}
