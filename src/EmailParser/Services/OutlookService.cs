using System.Runtime.InteropServices;
using EmailParser.Models;
using Microsoft.Office.Interop.Outlook;
using Serilog;

namespace EmailParser.Services;

/// <summary>
/// Reads emails from a specified Outlook folder using the Outlook COM interop API.
/// </summary>
public class OutlookService
{
    private static readonly ILogger Log = Serilog.Log.ForContext<OutlookService>();

    /// <summary>
    /// Returns all <see cref="MailItem"/> objects found in the specified Outlook folder,
    /// with their attachments saved to temporary files.
    /// </summary>
    /// <param name="folderPath">
    /// Folder name or path. Use '/' or '\' to separate nested levels,
    /// e.g. "Inbox", "Inbox/Projects", or "Archive/2024/Q1".
    /// The first segment is matched against the root folder names across all
    /// Outlook data stores (mailboxes / PST files).
    /// </param>
    public IEnumerable<EmailData> GetEmailsFromFolder(string folderPath)
    {
        Log.Information("Connecting to Outlook to read folder {FolderPath}", folderPath);

        Application? outlookApp = null;
        NameSpace? nameSpace = null;

        try
        {
            outlookApp = GetOrCreateOutlookApplication();
            nameSpace = outlookApp.GetNamespace("MAPI");
            nameSpace.Logon(Type.Missing, Type.Missing, false, false);

            MAPIFolder folder = ResolveFolder(nameSpace, folderPath);
            Log.Information("Resolved Outlook folder: {FolderName}", folder.Name);

            Items items = folder.Items;
            Log.Information("Found {Count} item(s) in folder", items.Count);

            foreach (object item in items)
            {
                if (item is not MailItem mailItem)
                {
                    Marshal.ReleaseComObject(item);
                    continue;
                }

                EmailData email;
                try
                {
                    email = ExtractEmailData(mailItem);
                }
                finally
                {
                    Marshal.ReleaseComObject(mailItem);
                }

                yield return email;
            }

            Marshal.ReleaseComObject(items);
            Marshal.ReleaseComObject(folder);
        }
        finally
        {
            if (nameSpace != null) Marshal.ReleaseComObject(nameSpace);
            if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
    }

    // Private helpers

    // P/Invoke declarations to access the COM Running Object Table (ROT),
    // replicating Marshal.GetActiveObject which was removed in .NET 5+.
    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid pclsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    /// <summary>
    /// Attempts to attach to a running Outlook instance before creating a new one.
    /// Reusing an existing instance avoids additional authentication prompts.
    /// </summary>
    private static Application GetOrCreateOutlookApplication()
    {
        try
        {
            int hr = CLSIDFromProgID("Outlook.Application", out Guid clsid);
            if (hr == 0)
            {
                GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
                if (obj is Application existing)
                {
                    Log.Debug("Attached to existing Outlook instance");
                    return existing;
                }
            }
        }
        catch
        {
            // Fall through to create a new instance.
        }

        try
        {
            Log.Debug("Creating new Outlook application instance");
            return new Application();
        }
        catch (System.Exception ex) when (OfficeAvailability.IsOfficeUnavailableException(ex))
        {
            throw new InvalidOperationException(
                "Microsoft Outlook does not appear to be installed on this machine. " +
                "Please install Microsoft Office (including Outlook), ensure it is " +
                "configured with a mail profile, and try again.",
                ex);
        }
    }

    /// <summary>
    /// Converts a <see cref="MailItem"/> into an <see cref="EmailData"/> record and
    /// saves all attachments to temporary files.
    /// </summary>
    private static EmailData ExtractEmailData(MailItem mailItem)
    {
        var email = new EmailData
        {
            Subject = mailItem.Subject ?? "(No Subject)",
            HtmlBody = mailItem.HTMLBody ?? string.Empty,
            TextBody = mailItem.Body ?? string.Empty,
            ReceivedTime = mailItem.ReceivedTime,
            From = mailItem.SenderName ?? string.Empty,
            To = mailItem.To ?? string.Empty,
        };

        Attachments attachments = mailItem.Attachments;
        for (int i = 1; i <= attachments.Count; i++)
        {
            Attachment att = attachments[i];
            try
            {
                string ext = Path.GetExtension(att.FileName);
                string tempPath = Path.Combine(
                    Path.GetTempPath(),
                    $"ep_{Path.GetRandomFileName()}{ext}");

                att.SaveAsFile(tempPath);

                email.Attachments.Add(new AttachmentData
                {
                    FileName = att.FileName,
                    TempFilePath = tempPath,
                });
            }
            finally
            {
                Marshal.ReleaseComObject(att);
            }
        }

        Marshal.ReleaseComObject(attachments);
        return email;
    }

    /// <summary>
    /// Resolves a folder path (e.g. "Inbox/Projects") against all Outlook stores.
    /// </summary>
    private static MAPIFolder ResolveFolder(NameSpace nameSpace, string folderPath)
    {
        string[] parts = folderPath.Split(
            new[] { '/', '\\' },
            StringSplitOptions.RemoveEmptyEntries);

        if (parts.Length == 0)
            throw new ArgumentException("Folder path cannot be empty.", nameof(folderPath));

        // First try special (default) folders for single-segment paths like "Inbox".
        if (parts.Length == 1)
        {
            MAPIFolder? special = TryGetSpecialFolder(nameSpace, parts[0]);
            if (special != null)
                return special;
        }

        // Walk every configured store (mailbox / PST) looking for a match.
        Stores stores = nameSpace.Stores;
        for (int s = 1; s <= stores.Count; s++)
        {
            Store store = stores[s];
            MAPIFolder rootFolder = store.GetRootFolder();
            try
            {
                MAPIFolder? found = FindFolder(rootFolder, parts, 0);
                if (found != null)
                    return found;
            }
            finally
            {
                Marshal.ReleaseComObject(rootFolder);
                Marshal.ReleaseComObject(store);
            }
        }

        Marshal.ReleaseComObject(stores);

        throw new InvalidOperationException(
            $"The folder '{folderPath}' was not found in any Outlook data store. " +
            "Please verify the folder name and try again.");
    }

    /// <summary>
    /// Attempts to map a well-known name (e.g. "Inbox", "Sent Items") to the
    /// corresponding <see cref="OlDefaultFolders"/> value.
    /// </summary>
    private static MAPIFolder? TryGetSpecialFolder(NameSpace nameSpace, string name)
    {
        var map = new Dictionary<string, OlDefaultFolders>(StringComparer.OrdinalIgnoreCase)
        {
            ["Inbox"]        = OlDefaultFolders.olFolderInbox,
            ["Sent Items"]   = OlDefaultFolders.olFolderSentMail,
            ["Sent Mail"]    = OlDefaultFolders.olFolderSentMail,
            ["Drafts"]       = OlDefaultFolders.olFolderDrafts,
            ["Deleted Items"]= OlDefaultFolders.olFolderDeletedItems,
            ["Junk Email"]   = OlDefaultFolders.olFolderJunk,
            ["Outbox"]       = OlDefaultFolders.olFolderOutbox,
        };

        if (map.TryGetValue(name, out OlDefaultFolders defaultFolder))
        {
            try { return nameSpace.GetDefaultFolder(defaultFolder); }
            catch { /* fall through */ }
        }

        return null;
    }

    /// <summary>
    /// Recursively descends into <paramref name="parent"/> matching each segment in
    /// <paramref name="parts"/> starting at <paramref name="index"/>.
    /// </summary>
    private static MAPIFolder? FindFolder(MAPIFolder parent, string[] parts, int index)
    {
        // If the current folder's name matches parts[index], descend further.
        if (!parent.Name.Equals(parts[index], StringComparison.OrdinalIgnoreCase))
        {
            // This branch of the tree doesn't start with the right name; search children.
            Folders subfolders = parent.Folders;
            for (int i = 1; i <= subfolders.Count; i++)
            {
                MAPIFolder sub = subfolders[i];
                if (sub.Name.Equals(parts[index], StringComparison.OrdinalIgnoreCase))
                {
                    Marshal.ReleaseComObject(subfolders);
                    if (index == parts.Length - 1)
                        return sub;

                    MAPIFolder? deeper = FindFolder(sub, parts, index + 1);
                    if (deeper != sub)
                        Marshal.ReleaseComObject(sub);
                    return deeper;
                }
                Marshal.ReleaseComObject(sub);
            }

            Marshal.ReleaseComObject(subfolders);
            return null;
        }

        // Current folder matches parts[index].
        if (index == parts.Length - 1)
            return parent;

        // Descend into children looking for parts[index + 1].
        Folders children = parent.Folders;
        for (int i = 1; i <= children.Count; i++)
        {
            MAPIFolder child = children[i];
            if (child.Name.Equals(parts[index + 1], StringComparison.OrdinalIgnoreCase))
            {
                Marshal.ReleaseComObject(children);
                if (index + 1 == parts.Length - 1)
                    return child;

                MAPIFolder? deeper = FindFolder(child, parts, index + 1);
                if (deeper != child)
                    Marshal.ReleaseComObject(child);
                return deeper;
            }

            Marshal.ReleaseComObject(child);
        }

        Marshal.ReleaseComObject(children);
        return null;
    }
}
