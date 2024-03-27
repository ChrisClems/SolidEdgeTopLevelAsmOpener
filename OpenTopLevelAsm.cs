using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using Application = SolidEdge.Framework.Interop.Application;

namespace SolidEdgeTopLevelAsmOpener
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Directory)]
    public class OpenTopLevelAsm : SharpContextMenu
    {
        private ContextMenuStrip menu = new ContextMenuStrip();

        protected override bool CanShowMenu()
        {
            // Check that only one item is selected and that it is a directory
            return SingleItem() && IsDirectory();
            bool SingleItem() => SelectedItemPaths.Count() == 1;
            bool IsDirectory() => File.GetAttributes(SelectedItemPaths.First()).HasFlag(FileAttributes.Directory);
        }

        protected override ContextMenuStrip CreateMenu()
        {
            // Mostly adapted from SharpShell example:
            // https://www.codeproject.com/Articles/1035998/NET-Shell-Extensions-Adding-submenus-to-Shell-Cont
            menu.Items.Clear();
            // TODO: Find way in SharpShell to handle pre-existing dropdown and add to it.
            ToolStripMenuItem SEMenu;
            SEMenu = new ToolStripMenuItem
            {
                Text = "Solid Edge"
            };

            var openParentAssemblies = new ToolStripMenuItem
            {
                Text = "Open Top-Level Assemblies"
            };
            openParentAssemblies.Click += (sender, args) => openAsmFiles();
            SEMenu.DropDownItems.Add(openParentAssemblies);
            menu.Items.Clear();
            menu.Items.Add(SEMenu);
            return menu;
        }

        private void openAsmFiles()
        {
            // Check if any .asm files exist in selected path before launching SE.
            var path = SelectedItemPaths.First();
            var asmFiles = Directory.GetFiles(path, "*.asm");
            Application SEApp = null;
            if (asmFiles.Any())
            {
                // Find running SE instance or launch new one.
                OleMessageFilter.Register();
                if (SEAppHelper.SEIsRunningForeground(out SEApp))
                {
                    SEApp.DoIdle();
                }
                else
                {
                    SEApp = SEAppHelper.SEStart();
                    SEApp.DoIdle();
                }
                // !! Must initialize Array with new operator or will throw NullRefEx
                Array topLevelAsms = new string[] { };
                SEApp.GetListOfTopLevelAssembliesFromFolder(path, out topLevelAsms);
                SEApp.DoIdle();
                if (topLevelAsms != null)
                {
                    foreach (string asm in topLevelAsms)
                    {
                        SEApp.Documents.Open(asm);
                        SEApp.DoIdle();
                    }
                }
            }

            OleMessageFilter.Revoke();
            if (SEApp != null)
            {
                Marshal.FinalReleaseComObject(SEApp);
            }
        }
    }
}