using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace AsyncToolWindowSample.Services
{
    /// <summary>
    /// Provides access to VS Document &amp; File APIs via DTE.
    /// All public methods must be called on the UI thread.
    /// </summary>
    public sealed class DocumentService
    {
        private readonly AsyncPackage _package;
        private readonly IServiceProvider _serviceProvider;

        public DocumentService(AsyncPackage package)
        {
            _package         = package ?? throw new ArgumentNullException(nameof(package));
            _serviceProvider = package;
        }

        // ------------------------------------------------------------------ //
        //  Internal helpers                                                    //
        // ------------------------------------------------------------------ //

        private DTE2 GetDte()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return _serviceProvider.GetService(typeof(DTE)) as DTE2;
        }

        /// <summary>
        /// Tries to read the CodePage/Encoding of a document via its TextDocument.
        /// Returns -1 when unavailable (binary, designer, etc.).
        /// </summary>
        private static int TryGetEncoding(Document doc)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var td = doc.Object("TextDocument") as TextDocument;
                // TextDocument does not expose Encoding directly either in DTE 8.0.
                // Use the underlying ProjectItem property if available.
                // Fallback: return 0 (unknown).
                return 0;
            }
            catch
            {
                return -1;
            }
        }

        // ================================================================== //
        //  Document Properties                                                 //
        // ================================================================== //

        /// <summary>
        /// Returns a snapshot of the active document's properties,
        /// or <c>null</c> when no document is active.
        /// </summary>
        public DocumentInfo GetActiveDocumentInfo()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = GetDte();
            var doc = dte?.ActiveDocument;
            if (doc == null) return null;

            string projName = null;
            string filePath = null;
            try
            {
                if (doc.ProjectItem != null)
                {
                    projName = doc.ProjectItem.ContainingProject?.Name;
                    filePath = doc.ProjectItem.FileNames[1];
                }
            }
            catch { /* ProjectItem may throw for misc files */ }

            return new DocumentInfo
            {
                Name            = doc.Name,
                FullName        = doc.FullName,
                Language        = doc.Language,
                Kind            = doc.Kind,
                Saved           = doc.Saved,
                ReadOnly        = doc.ReadOnly,
                ProjectName     = projName,
                ProjectFilePath = filePath
            };
        }

        /// <summary>Returns info for all currently open documents.</summary>
        public IReadOnlyList<DocumentInfo> GetAllOpenDocuments()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var result = new List<DocumentInfo>();
            var dte    = GetDte();
            if (dte == null) return result;

            foreach (Document doc in dte.Documents)
            {
                string projName = null;
                try { projName = doc.ProjectItem?.ContainingProject?.Name; } catch { }

                result.Add(new DocumentInfo
                {
                    Name        = doc.Name,
                    FullName    = doc.FullName,
                    Language    = doc.Language,
                    Kind        = doc.Kind,
                    Saved       = doc.Saved,
                    ReadOnly    = doc.ReadOnly,
                    ProjectName = projName
                });
            }

            return result;
        }

        // ================================================================== //
        //  Document Operations                                                 //
        // ================================================================== //

        /// <summary>Saves the active document. Returns <c>false</c> if none active.</summary>
        public bool SaveActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var doc = GetDte()?.ActiveDocument;
            if (doc == null) return false;
            doc.Save();
            return true;
        }

        /// <summary>Save-As the active document to <paramref name="newPath"/>.</summary>
        public bool SaveActiveDocumentAs(string newPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var doc = GetDte()?.ActiveDocument;
            if (doc == null) return false;
            doc.Save(newPath);
            return true;
        }

        /// <summary>
        /// Closes the active document.
        /// <paramref name="saveChanges"/>: true = save; false = discard.
        /// </summary>
        public bool CloseActiveDocument(bool saveChanges = true)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var doc = GetDte()?.ActiveDocument;
            if (doc == null) return false;
            doc.Close(saveChanges
                ? vsSaveChanges.vsSaveChangesYes
                : vsSaveChanges.vsSaveChangesNo);
            return true;
        }

        /// <summary>Undoes the last change in the active document.</summary>
        public bool UndoActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var doc = GetDte()?.ActiveDocument;
            if (doc == null) return false;
            doc.Undo();
            return true;
        }

        // ================================================================== //
        //  Open File / Navigate                                               //
        // ================================================================== //

        /// <summary>Opens a file in the VS code editor.</summary>
        public void OpenFile(string filePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            GetDte()?.ItemOperations.OpenFile(filePath, Constants.vsViewKindCode);
        }

        /// <summary>Navigates to a URL in VS or the default browser.</summary>
        public void NavigateUrl(string url)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            GetDte()?.ItemOperations.Navigate(url);
        }

        // ================================================================== //
        //  TextDocument & EditPoint                                            //
        // ================================================================== //

        /// <summary>
        /// Returns a <see cref="TextDocumentInfo"/> snapshot for the active document.
        /// Returns <c>null</c> when the active document is not a text document.
        /// </summary>
        public TextDocumentInfo GetTextDocumentInfo()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var doc = GetDte()?.ActiveDocument;
            if (doc == null) return null;

            TextDocument textDoc;
            try   { textDoc = doc.Object("TextDocument") as TextDocument; }
            catch { return null; }
            if (textDoc == null) return null;

            int firstLine  = textDoc.StartPoint.Line;
            int lastLine   = textDoc.EndPoint.Line;
            int totalChars = textDoc.EndPoint.AbsoluteCharOffset
                           - textDoc.StartPoint.AbsoluteCharOffset;

            EditPoint ep      = textDoc.StartPoint.CreateEditPoint();
            string    preview = ep.GetText(Math.Min(200, totalChars));

            return new TextDocumentInfo
            {
                FirstLine  = firstLine,
                LastLine   = lastLine,
                TotalChars = totalChars,
                Preview    = preview
            };
        }

        /// <summary>
        /// Reads a range of lines (1-based, endLine exclusive) from the active document.
        /// Returns <c>null</c> if no active text document.
        /// </summary>
        public string ReadLines(int startLine, int endLine)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var doc = GetDte()?.ActiveDocument;
            if (doc == null) return null;

            TextDocument textDoc;
            try   { textDoc = doc.Object("TextDocument") as TextDocument; }
            catch { return null; }
            if (textDoc == null) return null;

            EditPoint ep = textDoc.StartPoint.CreateEditPoint();
            return ep.GetLines(startLine, endLine);
        }

        /// <summary>
        /// Inserts <paramref name="text"/> at the beginning of the active document.
        /// Returns <c>false</c> if no active text document.
        /// </summary>
        public bool InsertAtStart(string text)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var doc = GetDte()?.ActiveDocument;
            if (doc == null) return false;

            TextDocument textDoc;
            try   { textDoc = doc.Object("TextDocument") as TextDocument; }
            catch { return false; }
            if (textDoc == null) return false;

            EditPoint ep = textDoc.StartPoint.CreateEditPoint();
            ep.Insert(text);
            return true;
        }

        // ================================================================== //
        //  ExecuteCommand                                                      //
        // ================================================================== //

        /// <summary>Executes a built-in VS command by name.</summary>
        public void ExecuteCommand(string commandName, string args = "")
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            GetDte()?.ExecuteCommand(commandName, args);
        }

        /// <summary>Formats the active document (Edit.FormatDocument).</summary>
        public void FormatDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ExecuteCommand("Edit.FormatDocument");
        }

        /// <summary>Saves all open documents (File.SaveAll).</summary>
        public void SaveAll()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ExecuteCommand("File.SaveAll");
        }

        /// <summary>Navigates to the specified 1-based line (Edit.GoToLine).</summary>
        public void GoToLine(int line)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ExecuteCommand("Edit.GoToLine", line.ToString());
        }
    }

    // ====================================================================== //
    //  Data transfer objects                                                   //
    // ====================================================================== //

    /// <summary>Snapshot of a VS document's properties.</summary>
    public sealed class DocumentInfo
    {
        public string Name            { get; set; }
        public string FullName        { get; set; }
        public string Language        { get; set; }
        public string Kind            { get; set; }
        public bool   Saved           { get; set; }
        public bool   ReadOnly        { get; set; }
        public string ProjectName     { get; set; }
        public string ProjectFilePath { get; set; }
    }

    /// <summary>Snapshot of a TextDocument's structure info.</summary>
    public sealed class TextDocumentInfo
    {
        public int    FirstLine  { get; set; }
        public int    LastLine   { get; set; }
        public int    TotalChars { get; set; }
        public string Preview    { get; set; }
    }
}
