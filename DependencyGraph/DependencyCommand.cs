using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace DependencyGraph
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class DependencyCommand
    {
        private const int S_OK = 0;

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3f46cd8e-102d-44b4-8963-d760b5f7e196");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="DependencyCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private DependencyCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static DependencyCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Verify the current thread is the UI thread - the call to AddCommand in DependencyCommand's constructor requires
            // the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new DependencyCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Task.Run(async () => await ExecuteAsync());
        }

        private async Task ExecuteAsync()
        {
            var _applicationObject = await ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE2;
            var serviceProvider = (Microsoft.VisualStudio.OLE.Interop.IServiceProvider)_applicationObject;
            var solutionService = GetService<SVsSolution, IVsSolution>(serviceProvider);

            var projectDependencies = LoadProjectDependencies(_applicationObject);
            var graph = CreateGraphFrom(projectDependencies);
            var form = CreateFormFor(graph);

            form.ShowDialog();
        }

        private static IEnumerable<KeyValuePair<string, string>> LoadProjectDependencies(DTE2 _applicationObject)
        {
            var solutionBuild = _applicationObject.Solution.SolutionBuild;
            var buildDependencies = solutionBuild.BuildDependencies
                .OfType<BuildDependency>()
                .ToDictionary(bd => bd.Project.UniqueName);

            var startupProjects = (solutionBuild.StartupProjects as object[])?.OfType<string>().ToArray();

            return CreateDependencyList(buildDependencies, startupProjects);
        }

        private static IEnumerable<KeyValuePair<string, string>> CreateDependencyList(Dictionary<string, BuildDependency> buildDependencies, string[] startupProjects)
        {
            var processedProjects = new List<string>();
            var pendingProjects = new Stack<string>(startupProjects);

            while (pendingProjects.Any())
            {
                var next = pendingProjects.Pop();
                if (!processedProjects.Contains(next))
                {
                    processedProjects.Add(next);

                    var dep = buildDependencies[next];
                    var requiredProjects = dep.RequiredProjects as object[];
                    foreach (var req in requiredProjects)
                    {
                        var reqProj = req as Project;
                        pendingProjects.Push(reqProj.UniqueName);
                        yield return new KeyValuePair<string, string>(dep.Project.Name, reqProj.Name);
                    }
                }
            }
        }

        private static Microsoft.Msagl.Drawing.Graph CreateGraphFrom(IEnumerable<KeyValuePair<string, string>> dependencies)
        {
            var graph = new Microsoft.Msagl.Drawing.Graph("graph");
            foreach (var edge in dependencies)
            {
                graph.AddEdge(edge.Key, edge.Value);
            }

            //graph.AddEdge("A", "B");
            //graph.AddEdge("B", "C");
            //graph.AddEdge("A", "C").Attr.Color = Microsoft.Msagl.Drawing.Color.Green;
            //graph.FindNode("A").Attr.FillColor = Microsoft.Msagl.Drawing.Color.Magenta;
            //graph.FindNode("B").Attr.FillColor = Microsoft.Msagl.Drawing.Color.MistyRose;
            //Microsoft.Msagl.Drawing.Node c = graph.FindNode("C");
            //c.Attr.FillColor = Microsoft.Msagl.Drawing.Color.PaleGreen;
            //c.Attr.Shape = Microsoft.Msagl.Drawing.Shape.Diamond;
            return graph;
        }

        private static System.Windows.Forms.Form CreateFormFor(Microsoft.Msagl.Drawing.Graph graph)
        {
            var form = new System.Windows.Forms.Form();
            form.SuspendLayout();

            var viewer = new Microsoft.Msagl.GraphViewerGdi.GViewer
            {
                Graph = graph,
                Dock = System.Windows.Forms.DockStyle.Fill
            };
            form.Controls.Add(viewer);

            form.ResumeLayout();
            return form;
        }

        private I GetService<T, I>(Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider)
        {
            return (I)GetService(serviceProvider, typeof(T), typeof(I));
        }

        private object GetService(Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider, Type serviceType, Type interfaceType)
        {
            object service = null;
            IntPtr servicePointer;
            int hr = 0;
            Guid serviceGuid;
            Guid interfaceGuid;

            serviceGuid = serviceType.GUID;
            interfaceGuid = interfaceType.GUID;

            hr = serviceProvider.QueryService(ref serviceGuid, ref interfaceGuid, out servicePointer);
            if (hr != S_OK)
            {
                System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr);
            }
            else if (servicePointer != IntPtr.Zero)
            {
                service = System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(servicePointer);
                System.Runtime.InteropServices.Marshal.Release(servicePointer);
            }
            return service;
        }
    }
}
