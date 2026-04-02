using Autodesk.Revit.UI;
using Connections.Services.Core;
using System;
using System.Collections.Concurrent;

namespace Connections.Services.Revit
{
    /// <summary>
    /// Manages the execution of external events in Revit.
    /// Implementation of the 'Queue' pattern for safe modeless interaction.
    /// </summary>
    public class RevitExternalEventService : IExternalEventHandler
    {
        private readonly ExternalEvent _externalEvent;
        private readonly ConcurrentQueue<IExternalEventRequest> _requestQueue;
        private readonly ILogger _logger;

        public RevitExternalEventService(ILogger logger)
        {
            _logger = logger;
            _requestQueue = new ConcurrentQueue<IExternalEventRequest>();
            _externalEvent = ExternalEvent.Create(this);
            _logger.Info("RevitExternalEventService initialized.");
        }

        /// <summary>
        /// Queues a request and signals the external event.
        /// Safe to call from any thread (UI thread).
        /// </summary>
        /// <param name="request">The request to execute.</param>
        public void Raise(IExternalEventRequest request)
        {
            _requestQueue.Enqueue(request);
            _externalEvent.Raise();
        }

        /// <summary>
        /// Executed by Revit in a valid API context.
        /// </summary>
        public void Execute(UIApplication app)
        {
            try
            {
                // Process all queued requests
                while (_requestQueue.TryDequeue(out var request))
                {
                    try
                    {
                        request.Execute(app);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"Error executing external event request: {request.GetType().Name}", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Critical error in external event loop", ex);
            }
        }

        public string GetName()
        {
            return "Connections External Event Service";
        }
    }
}
